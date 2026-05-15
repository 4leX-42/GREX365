#requires -Version 7.4
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
try { $Host.UI.RawUI.WindowTitle = 'GREX365 · Grupos · workflow unificado' } catch {}

# Single CSV-driven workflow for groups:
#  - Resolve / create groups (DL, Mail-enabled security, M365 group)
#  - Skip if exists
#  - Set owners
#  - Add members
#  - Validate end state
#  - Generate HTML report
#
# CSV schema (semicolon):
#   Action;Type;Name;DisplayName;Alias;Owners;Members;HiddenFromGAL
#
#   Action       = ensure | create | members-only | owners-only
#   Type         = DL | M365 | MailSecurity
#   Name         = mail alias / primary smtp local part (or full smtp)
#   DisplayName  = visible name
#   Alias        = optional, defaults to Name
#   Owners       = comma-separated UPNs (only used by M365)
#   Members      = comma-separated UPNs / SMTP
#   HiddenFromGAL = true | false (optional)

Assert-RequiredServicesReady
Assert-Role -Required operator -Operation 'gestión de grupos'

Show-Header -Title 'GREX365' -Subtitle 'Workflow unificado de grupos'

$session = Start-LogSession -Name 'GroupsWorkflow'
$started = Get-Date
$results = New-Object System.Collections.Generic.List[object]

function Add-Result {
    param(
        [string]$Group,[string]$Phase,[string]$Detail,
        [string]$Estado = 'OK',[string]$Type = '',[string]$Action = ''
    )
    $results.Add([PSCustomObject]@{
        Grupo  = $Group
        Tipo   = $Type
        Action = $Action
        Phase  = $Phase
        Estado = $Estado
        Detalle = $Detail
    })
    $level = switch ($Estado) { 'OK' { 'OK' } 'ERROR' { 'ERROR' } default { 'WARN' } }
    Write-Log ("[{0}] {1} · {2}{3}" -f $Estado, $Group, $Phase, $(if ($Detail) { ' — ' + $Detail } else { '' })) -Level $level -Source 'GroupsWorkflow'
}

function Resolve-ExistingGroup {
    param([string]$Identity)
    try {
        $r = Get-Recipient -Identity $Identity -ErrorAction Stop | Select-Object -First 1
        if ($r) { return $r }
    } catch {}
    return $null
}

function Split-CsvList {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return @() }
    return @($Value -split '[,;|\s]+' | Where-Object { $_ -and -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() })
}

try {
    $csvPath = Read-ValidatedCsvPath -Prompt 'CSV (Action;Type;Name;DisplayName;Alias;Owners;Members;HiddenFromGAL)'
    $rows = @(Import-Csv -Path $csvPath -Delimiter ';' -Encoding UTF8)
    if (-not $rows -or $rows.Count -eq 0) { throw 'CSV vacío.' }

    # Action column is OPTIONAL now. If missing, the script assumes 'ensure' (create-if-missing-and-update).
    # Type/Name/DisplayName remain required.
    $required = @('Type','Name','DisplayName')
    foreach ($col in $required) {
        if ($rows[0].PSObject.Properties.Name -notcontains $col) { throw "Columna requerida ausente: $col" }
    }
    $hasActionColumn = ($rows[0].PSObject.Properties.Name -contains 'Action')
    if (-not $hasActionColumn) {
        Write-Log "Columna 'Action' no encontrada · usando default 'ensure' (crear si no existe + aplicar owners/miembros/GAL)." -Level WARN -Source 'GroupsWorkflow'
    }

    Show-Section -Title 'Vista previa'
    Write-KeyValue -Key 'CSV'   -Value $csvPath
    Write-KeyValue -Key 'Filas' -Value $rows.Count

    if (-not (Confirm-DestructiveAction -Operation 'Aplicar workflow de grupos' -Target ("$($rows.Count) entradas"))) {
        throw 'INPUT_EMPTY'
    }

    $idx = 0
    foreach ($row in $rows) {
        $idx++
        $action  = if ($hasActionColumn) { ([string]$row.Action).Trim().ToLowerInvariant() } else { 'ensure' }
        if ([string]::IsNullOrWhiteSpace($action)) { $action = 'ensure' }
        $type    = ([string]$row.Type).Trim()
        $name    = ([string]$row.Name).Trim()
        $disp    = ([string]$row.DisplayName).Trim()
        $alias   = if ($row.PSObject.Properties.Name -contains 'Alias') { ([string]$row.Alias).Trim() } else { '' }
        $owners  = @()
        if ($row.PSObject.Properties.Name -contains 'Owners')  { $owners  = @(Split-CsvList -Value ([string]$row.Owners)) }
        $members = @()
        if ($row.PSObject.Properties.Name -contains 'Members') { $members = @(Split-CsvList -Value ([string]$row.Members)) }
        $hide    = $false
        if ($row.PSObject.Properties.Name -contains 'HiddenFromGAL') {
            $hide = [bool]([string]$row.HiddenFromGAL).Trim().ToLowerInvariant().StartsWith('t')
        }

        Write-Progress -Activity 'Procesando grupos' -Status "$idx / $($rows.Count) · $name" -PercentComplete ([math]::Round(($idx / $rows.Count) * 100, 1))

        if (-not $alias) { $alias = $name -replace '@.*$','' }
        if ([string]::IsNullOrWhiteSpace($name) -or [string]::IsNullOrWhiteSpace($type)) {
            Add-Result -Group $name -Type $type -Action $action -Phase 'validate' -Estado 'ERROR' -Detail 'Name/Type vacíos'
            continue
        }
        if ($action -notin @('ensure','create','members-only','owners-only')) {
            Add-Result -Group $name -Type $type -Action $action -Phase 'validate' -Estado 'ERROR' -Detail "Action no soportada: $action"
            continue
        }
        if ($type -notin @('DL','M365','MailSecurity')) {
            Add-Result -Group $name -Type $type -Action $action -Phase 'validate' -Estado 'ERROR' -Detail "Type no soportado: $type"
            continue
        }

        # --- resolve or create ---
        $existing = Resolve-ExistingGroup -Identity $name
        if ($action -eq 'create' -and $existing) {
            Add-Result -Group $name -Type $type -Action $action -Phase 'create' -Estado 'SKIP' -Detail 'Ya existe'
            $existing = $existing
        }
        elseif (-not $existing -and ($action -in @('ensure','create'))) {
            try {
                switch ($type) {
                    'DL' {
                        Invoke-WithRetry -OperationName 'New-DistributionGroup' -Quiet -ScriptBlock {
                            New-DistributionGroup -Name $disp -Alias $alias -PrimarySmtpAddress $name -Type Distribution -ErrorAction Stop | Out-Null
                        }
                    }
                    'MailSecurity' {
                        Invoke-WithRetry -OperationName 'New-DistributionGroup Security' -Quiet -ScriptBlock {
                            New-DistributionGroup -Name $disp -Alias $alias -PrimarySmtpAddress $name -Type Security -ErrorAction Stop | Out-Null
                        }
                    }
                    'M365' {
                        Invoke-WithRetry -OperationName 'New-UnifiedGroup' -Quiet -ScriptBlock {
                            New-UnifiedGroup -DisplayName $disp -Alias $alias -PrimarySmtpAddress $name -AccessType Private -ErrorAction Stop | Out-Null
                        }
                    }
                }
                Add-Result -Group $name -Type $type -Action $action -Phase 'create' -Estado 'OK' -Detail "Creado ($type)"
                $existing = Resolve-ExistingGroup -Identity $name
            } catch {
                Add-Result -Group $name -Type $type -Action $action -Phase 'create' -Estado 'ERROR' -Detail $_.Exception.Message
                continue
            }
        }
        elseif (-not $existing) {
            Add-Result -Group $name -Type $type -Action $action -Phase 'resolve' -Estado 'ERROR' -Detail 'No existe y action != create/ensure'
            continue
        }
        else {
            Add-Result -Group $name -Type $type -Action $action -Phase 'resolve' -Estado 'OK' -Detail 'Existe'
        }

        # --- owners (M365 only) ---
        if ($action -in @('ensure','create','owners-only') -and $type -eq 'M365' -and $owners.Count -gt 0) {
            $addedOwners = 0
            foreach ($own in $owners) {
                try {
                    Invoke-WithRetry -OperationName 'Add-UnifiedGroupLinks Owners' -Quiet -ScriptBlock {
                        Add-UnifiedGroupLinks -Identity $name -LinkType Owners -Links $own -ErrorAction Stop
                    }
                    $addedOwners++
                } catch {
                    Add-Result -Group $name -Type $type -Action $action -Phase 'owners' -Estado 'WARN' -Detail "$own — $($_.Exception.Message)"
                }
            }
            Add-Result -Group $name -Type $type -Action $action -Phase 'owners' -Estado 'OK' -Detail "$addedOwners owners aplicados"
        }

        # --- members ---
        if ($action -in @('ensure','create','members-only') -and $members.Count -gt 0) {
            $addedMembers = 0
            $skipMembers  = 0
            foreach ($m in $members) {
                try {
                    if ($type -eq 'M365') {
                        Invoke-WithRetry -OperationName 'Add-UnifiedGroupLinks Members' -Quiet -ScriptBlock {
                            Add-UnifiedGroupLinks -Identity $name -LinkType Members -Links $m -ErrorAction Stop
                        }
                    } else {
                        Invoke-WithRetry -OperationName 'Add-DistributionGroupMember' -Quiet -ScriptBlock {
                            Add-DistributionGroupMember -Identity $name -Member $m -ErrorAction Stop
                        }
                    }
                    $addedMembers++
                } catch {
                    if ($_.Exception.Message -match 'already a member|exists|duplicate') { $skipMembers++ }
                    else { Add-Result -Group $name -Type $type -Action $action -Phase 'members' -Estado 'WARN' -Detail "$m — $($_.Exception.Message)" }
                }
            }
            Add-Result -Group $name -Type $type -Action $action -Phase 'members' -Estado 'OK' -Detail "$addedMembers añadidos · $skipMembers ya existían"
        }

        # --- hide from GAL ---
        if ($action -in @('ensure','create') -and $hide) {
            try {
                if ($type -eq 'M365') {
                    Invoke-WithRetry -OperationName 'Set-UnifiedGroup hide' -Quiet -ScriptBlock {
                        Set-UnifiedGroup -Identity $name -HiddenFromAddressListsEnabled $true -ErrorAction Stop
                    }
                } else {
                    Invoke-WithRetry -OperationName 'Set-DistributionGroup hide' -Quiet -ScriptBlock {
                        Set-DistributionGroup -Identity $name -HiddenFromAddressListsEnabled $true -ErrorAction Stop
                    }
                }
                Add-Result -Group $name -Type $type -Action $action -Phase 'gal' -Estado 'OK' -Detail 'Hidden=true'
            } catch {
                Add-Result -Group $name -Type $type -Action $action -Phase 'gal' -Estado 'WARN' -Detail $_.Exception.Message
            }
        }

        # --- validate ---
        $finalCount = 0
        try {
            if ($type -eq 'M365') {
                $finalCount = @(Get-UnifiedGroupLinks -Identity $name -LinkType Members -ResultSize Unlimited -ErrorAction Stop).Count
            } else {
                $finalCount = @(Get-DistributionGroupMember -Identity $name -ResultSize Unlimited -ErrorAction Stop).Count
            }
            Add-Result -Group $name -Type $type -Action $action -Phase 'validate' -Estado 'OK' -Detail "Miembros finales: $finalCount"
        } catch {
            Add-Result -Group $name -Type $type -Action $action -Phase 'validate' -Estado 'WARN' -Detail $_.Exception.Message
        }
    }
    Write-Progress -Activity 'Procesando grupos' -Completed

    $okCount   = @($results | Where-Object Estado -eq 'OK').Count
    $warnCount = @($results | Where-Object Estado -eq 'WARN').Count
    $errCount  = @($results | Where-Object Estado -eq 'ERROR').Count
    $skipCount = @($results | Where-Object Estado -eq 'SKIP').Count

    Show-Section -Title 'Resumen'
    Write-KeyValue -Key 'Eventos OK'    -Value $okCount -ValueColor Green
    Write-KeyValue -Key 'Eventos WARN'  -Value $warnCount -ValueColor Yellow
    Write-KeyValue -Key 'Eventos ERROR' -Value $errCount -ValueColor Red
    Write-KeyValue -Key 'Eventos SKIP'  -Value $skipCount -ValueColor DarkGray

    $outFolder = Split-Path -Path $csvPath -Parent
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $resCsv = Join-Path $outFolder ("groups_workflow_$stamp.csv")
    $results | Export-Csv -Path $resCsv -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-KeyValue -Key 'CSV resultado' -Value $resCsv

    $logFile = Stop-LogSession -Persist
    $report = Publish-OperationReport `
        -Title 'Workflow de grupos · informe' `
        -Operation 'Invoke-GroupsWorkflow' `
        -StartTime $started -EndTime (Get-Date) `
        -Summary ([ordered]@{
            'Filas CSV'    = $rows.Count
            'Eventos OK'   = $okCount
            'Eventos WARN' = $warnCount
            'Eventos ERR'  = $errCount
            'Eventos SKIP' = $skipCount
        }) `
        -Items $results `
        -OkFields @('Eventos OK') -ErrorFields @('Eventos ERR') -WarnFields @('Eventos WARN','Eventos SKIP') `
        -LogFile $logFile
    Write-KeyValue -Key 'Informe HTML' -Value $report

} catch {
    if ($_.Exception.Message -in @('INPUT_EMPTY','INPUT_INVALID','INPUT_NOTFOUND')) {
        Show-WarningBlock -Title 'Cancelado'
    } else {
        Show-ErrorBlock -Title 'Workflow fallido' -Detail $_.Exception.Message
    }
    Stop-LogSession -Persist | Out-Null
    throw
}


