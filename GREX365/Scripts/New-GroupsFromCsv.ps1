#requires -Version 7.4
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [string[]]$FolderPaths,
    [string]$Domain
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
try { $Host.UI.RawUI.WindowTitle = 'GREX365 · Crear grupos/DL desde CSV' } catch {}

# Helpers come from GREX365/Modules.

$script:LogRows      = New-Object System.Collections.Generic.List[object]
$script:CreatedLinks = New-Object System.Collections.Generic.List[object]
$script:GroupType    = ''
$script:Domain       = ''
$script:WhatIfMode   = $WhatIfPreference -ne [System.Management.Automation.ActionPreference]::SilentlyContinue

function Add-LogRow {
    param(
        [string]$CsvFile = '', [string]$GroupName = '', [string]$GroupEmail = '',
        [string]$Action = '', [string]$UserEmail = '', [string]$Detail = ''
    )
    [void]$script:LogRows.Add([PSCustomObject]@{
        Timestamp  = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        CsvFile    = $CsvFile
        GroupName  = $GroupName
        GroupEmail = $GroupEmail
        Action     = $Action
        UserEmail  = $UserEmail
        Detail     = $Detail
    })
}

# --- M365 GROUP OPS ---

function Get-OrCreate-M365Group {
    param([string]$DisplayName, [string]$GroupEmail, [string]$CsvFile)

    $existing = @(Get-MgGroup -Filter "mail eq '$GroupEmail'" -ErrorAction SilentlyContinue)
    if ($existing.Count -gt 0) {
        Write-Log ('[SKIP] M365 Group ya existe: ' + $GroupEmail) -Level WARN -Source 'NewGroups'
        Add-LogRow -CsvFile $CsvFile -GroupName $DisplayName -GroupEmail $GroupEmail -Action 'Skipped' -Detail 'Ya existía'
        return $existing[0]
    }

    if ($script:WhatIfMode) {
        Write-Log ('[WhatIf] Crearía M365 Group: ' + $GroupEmail) -Level WARN -Source 'NewGroups'
        Add-LogRow -CsvFile $CsvFile -GroupName $DisplayName -GroupEmail $GroupEmail -Action 'WhatIf-Created'
        return $null
    }

    if ($PSCmdlet.ShouldProcess($GroupEmail, 'Crear M365 Group')) {
        $alias = ($GroupEmail -split '@')[0]
        $body = @{
            DisplayName     = $DisplayName
            MailNickname    = $alias
            MailEnabled     = $true
            SecurityEnabled = $false
            GroupTypes      = @('Unified')
            Visibility      = 'Private'
        }
        $g = New-MgGroup -BodyParameter $body -ErrorAction Stop
        Write-Log ('[CREATED] M365 Group: ' + $GroupEmail) -Level OK -Source 'NewGroups'
        Add-LogRow -CsvFile $CsvFile -GroupName $DisplayName -GroupEmail $GroupEmail -Action 'Created'

        if ($g.Id) {
            $script:CreatedLinks.Add([PSCustomObject]@{
                Type        = 'Microsoft365Group'
                Id          = [string]$g.Id
                DisplayName = $DisplayName
                Email       = $GroupEmail
            })
            Write-Host ''
            Show-AdminLink -Type 'Microsoft365Group' -Id $g.Id -DisplayName $DisplayName
        }
        return $g
    }
    return $null
}

function Get-M365GroupMembers {
    param([string]$GroupId)

    $map = @{}
    $members = @(Get-MgGroupMember -GroupId $GroupId -All -ErrorAction SilentlyContinue)
    foreach ($m in $members) {
        $props = $m.AdditionalProperties
        if ($props) {
            foreach ($key in @('mail','userPrincipalName')) {
                $val = $props[$key]
                if ($val -and -not [string]::IsNullOrWhiteSpace([string]$val)) {
                    $map[([string]$val).ToLowerInvariant()] = $m.Id
                }
            }
        }
        if ($m.Id) { $map[$m.Id.ToLowerInvariant()] = $m.Id }
    }
    return $map
}

function Add-M365Member {
    param(
        [string]$GroupId, [string]$UserEmail, [string]$GroupEmail,
        [string]$GroupName, [string]$CsvFile, [hashtable]$MemberMap
    )

    $key = $UserEmail.ToLowerInvariant()
    if ($MemberMap.ContainsKey($key)) {
        Write-Log ('[SKIP] Ya es miembro: ' + $UserEmail) -Level WARN -Source 'NewGroups'
        Add-LogRow -CsvFile $CsvFile -GroupName $GroupName -GroupEmail $GroupEmail -Action 'MemberSkipped' -UserEmail $UserEmail -Detail 'Ya pertenece'
        return
    }

    $mgUsers = @(Get-MgUser -Filter "mail eq '$UserEmail' or userPrincipalName eq '$UserEmail'" -Property Id -ConsistencyLevel eventual -All -ErrorAction SilentlyContinue)
    if ($mgUsers.Count -eq 0) {
        Write-Log ('Usuario no encontrado en Graph: ' + $UserEmail) -Level ERROR -Source 'NewGroups'
        Add-LogRow -CsvFile $CsvFile -GroupName $GroupName -GroupEmail $GroupEmail -Action 'Error' -UserEmail $UserEmail -Detail 'No encontrado en Graph'
        return
    }
    $mgUser = $mgUsers[0]

    if ($script:WhatIfMode) {
        Write-Log ('[WhatIf] Añadiría: ' + $UserEmail) -Level WARN -Source 'NewGroups'
        Add-LogRow -CsvFile $CsvFile -GroupName $GroupName -GroupEmail $GroupEmail -Action 'WhatIf-MemberAdded' -UserEmail $UserEmail
        return
    }

    if ($PSCmdlet.ShouldProcess($UserEmail, "Añadir a M365 Group $GroupEmail")) {
        New-MgGroupMember -GroupId $GroupId -DirectoryObjectId $mgUser.Id -ErrorAction Stop
        $MemberMap[$key] = $mgUser.Id
        Write-Log ('Añadido: ' + $UserEmail) -Level OK -Source 'NewGroups'
        Add-LogRow -CsvFile $CsvFile -GroupName $GroupName -GroupEmail $GroupEmail -Action 'MemberAdded' -UserEmail $UserEmail
    }
}

# --- DL OPS ---

function Get-OrCreate-DL {
    param([string]$DisplayName, [string]$GroupEmail, [string]$CsvFile)

    $existing = @(Get-DistributionGroup -Identity $GroupEmail -ErrorAction SilentlyContinue)
    if ($existing.Count -gt 0) {
        Write-Log ('[SKIP] DL ya existe: ' + $GroupEmail) -Level WARN -Source 'NewGroups'
        Add-LogRow -CsvFile $CsvFile -GroupName $DisplayName -GroupEmail $GroupEmail -Action 'Skipped' -Detail 'Ya existía'
        return $existing[0]
    }

    if ($script:WhatIfMode) {
        Write-Log ('[WhatIf] Crearía DL: ' + $GroupEmail) -Level WARN -Source 'NewGroups'
        Add-LogRow -CsvFile $CsvFile -GroupName $DisplayName -GroupEmail $GroupEmail -Action 'WhatIf-Created'
        return $null
    }

    if ($PSCmdlet.ShouldProcess($GroupEmail, 'Crear DL')) {
        $dl = New-DistributionGroup -Name $DisplayName -PrimarySmtpAddress $GroupEmail -Type Distribution -ErrorAction Stop
        Write-Log ('[CREATED] DL: ' + $GroupEmail) -Level OK -Source 'NewGroups'
        Add-LogRow -CsvFile $CsvFile -GroupName $DisplayName -GroupEmail $GroupEmail -Action 'Created'

        $dlId = ''
        if ($dl.ExternalDirectoryObjectId) { $dlId = [string]$dl.ExternalDirectoryObjectId }
        elseif ($dl.Guid)                  { $dlId = [string]$dl.Guid }
        if ($dlId) {
            $script:CreatedLinks.Add([PSCustomObject]@{
                Type        = 'DistributionList'
                Id          = $dlId
                DisplayName = $DisplayName
                Email       = $GroupEmail
            })
            Write-Host ''
            Show-AdminLink -Type 'DistributionList' -Id $dlId -DisplayName $DisplayName
        }

        Start-Sleep -Seconds 3
        return $dl
    }
    return $null
}

function Get-DLMemberMap {
    param([string]$Identity)
    $map = @{}
    $members = @(Get-DistributionGroupMember -Identity $Identity -ResultSize Unlimited -ErrorAction SilentlyContinue)
    foreach ($m in $members) {
        if ($m.PrimarySmtpAddress) {
            $smtp = $m.PrimarySmtpAddress.ToString().Trim()
            if ($smtp) { $map[$smtp.ToLowerInvariant()] = $true }
        }
    }
    return $map
}

function Add-DLMember {
    param(
        [string]$Identity, [string]$UserEmail, [string]$GroupEmail,
        [string]$GroupName, [string]$CsvFile, [hashtable]$MemberMap
    )

    $key = $UserEmail.ToLowerInvariant()
    if ($MemberMap.ContainsKey($key)) {
        Write-Log ('[SKIP] Ya es miembro: ' + $UserEmail) -Level WARN -Source 'NewGroups'
        Add-LogRow -CsvFile $CsvFile -GroupName $GroupName -GroupEmail $GroupEmail -Action 'MemberSkipped' -UserEmail $UserEmail -Detail 'Ya pertenece'
        return
    }

    if ($script:WhatIfMode) {
        Write-Log ('[WhatIf] Añadiría: ' + $UserEmail) -Level WARN -Source 'NewGroups'
        Add-LogRow -CsvFile $CsvFile -GroupName $GroupName -GroupEmail $GroupEmail -Action 'WhatIf-MemberAdded' -UserEmail $UserEmail
        return
    }

    if ($PSCmdlet.ShouldProcess($UserEmail, "Añadir a DL $GroupEmail")) {
        Add-DistributionGroupMember -Identity $Identity -Member $UserEmail -ErrorAction Stop
        $MemberMap[$key] = $true
        Write-Log ('Añadido: ' + $UserEmail) -Level OK -Source 'NewGroups'
        Add-LogRow -CsvFile $CsvFile -GroupName $GroupName -GroupEmail $GroupEmail -Action 'MemberAdded' -UserEmail $UserEmail
    }
}

# --- PER-CSV PROCESSING ---

function Invoke-ProcessCsv {
    param([string]$CsvPath)

    $csvName = Split-Path $CsvPath -Leaf
    Show-Section -Title $csvName

    $rows = @()
    try {
        $rows = @(Import-FlexibleCsv -Path $CsvPath -Schema EmailGroupName)
        $lastGroupName = ''
        foreach ($r in $rows) {
            if (-not [string]::IsNullOrWhiteSpace($r.GroupName)) { $lastGroupName = $r.GroupName }
            elseif ($lastGroupName) { $r.GroupName = $lastGroupName }
        }
    } catch {
        Write-Log ('CSV ignorado: ' + $_) -Level ERROR -Source 'NewGroups'
        Add-LogRow -CsvFile $csvName -Action 'Error' -Detail ('Lectura fallida: ' + $_)
        return
    }

    if ($rows.Count -eq 0) { Write-Log 'CSV sin filas.' -Level WARN -Source 'NewGroups'; return }

    $validGroups = @($rows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.GroupName) })
    $ignored = $rows.Count - $validGroups.Count
    if ($ignored -gt 0) { Write-Log ("$ignored fila(s) sin GroupName — omitidas") -Level WARN -Source 'NewGroups' }
    if ($validGroups.Count -eq 0) { Write-Log 'Sin filas válidas.' -Level WARN -Source 'NewGroups'; return }

    $grouped = $validGroups | Group-Object -Property GroupName

    foreach ($grp in $grouped) {
        $groupName  = $grp.Name.Trim()
        $groupEmail = "$groupName@$($script:Domain)"
        Write-Log ('Grupo: ' + $groupEmail + ' (' + $grp.Group.Count + ' miembro(s))') -Source 'NewGroups'

        $validMembers = @($grp.Group | Where-Object {
            $em = $_.Email.Trim()
            if (-not $em) {
                Write-Log 'Fila vacía en Email — omitida' -Level WARN -Source 'NewGroups'
                Add-LogRow -CsvFile $csvName -GroupName $groupName -GroupEmail $groupEmail -Action 'Error' -Detail 'Email vacío'
                return $false
            }
            if (-not (Test-Email -Value $em)) {
                Write-Log ('Email inválido: ' + $em) -Level WARN -Source 'NewGroups'
                Add-LogRow -CsvFile $csvName -GroupName $groupName -GroupEmail $groupEmail -Action 'Error' -UserEmail $em -Detail 'Formato inválido'
                return $false
            }
            return $true
        })

        if ($script:GroupType -eq 'M365') {
            $mg = $null
            try { $mg = Get-OrCreate-M365Group -DisplayName $groupName -GroupEmail $groupEmail -CsvFile $csvName }
            catch {
                Write-Log ('Error creando M365 Group: ' + $_) -Level ERROR -Source 'NewGroups'
                Add-LogRow -CsvFile $csvName -GroupName $groupName -GroupEmail $groupEmail -Action 'Error' -Detail ('Creación fallida: ' + $_)
                continue
            }
            if (-not $mg) { continue }

            $memberMap = Get-M365GroupMembers -GroupId $mg.Id
            foreach ($r in $validMembers) {
                try { Add-M365Member -GroupId $mg.Id -UserEmail $r.Email.Trim() -GroupEmail $groupEmail -GroupName $groupName -CsvFile $csvName -MemberMap $memberMap }
                catch {
                    Write-Log ('Error añadiendo ' + $r.Email + ': ' + $_) -Level ERROR -Source 'NewGroups'
                    Add-LogRow -CsvFile $csvName -GroupName $groupName -GroupEmail $groupEmail -Action 'Error' -UserEmail $r.Email -Detail ('Adición fallida: ' + $_)
                }
            }
        } else {
            $dl = $null
            try { $dl = Get-OrCreate-DL -DisplayName $groupName -GroupEmail $groupEmail -CsvFile $csvName }
            catch {
                Write-Log ('Error creando DL: ' + $_) -Level ERROR -Source 'NewGroups'
                Add-LogRow -CsvFile $csvName -GroupName $groupName -GroupEmail $groupEmail -Action 'Error' -Detail ('Creación fallida: ' + $_)
                continue
            }
            if (-not $dl) { continue }

            $memberMap = Get-DLMemberMap -Identity $dl.PrimarySmtpAddress.ToString()
            foreach ($r in $validMembers) {
                try { Add-DLMember -Identity $dl.PrimarySmtpAddress.ToString() -UserEmail $r.Email.Trim() -GroupEmail $groupEmail -GroupName $groupName -CsvFile $csvName -MemberMap $memberMap }
                catch {
                    Write-Log ('Error añadiendo ' + $r.Email + ': ' + $_) -Level ERROR -Source 'NewGroups'
                    Add-LogRow -CsvFile $csvName -GroupName $groupName -GroupEmail $groupEmail -Action 'Error' -UserEmail $r.Email -Detail ('Adición fallida: ' + $_)
                }
            }
        }
    }
}

function Show-Summary {
    param([string]$LogCsvPath)

    $created  = @($script:LogRows | Where-Object { $_.Action -eq 'Created' }).Count
    $skipped  = @($script:LogRows | Where-Object { $_.Action -eq 'Skipped' }).Count
    $added    = @($script:LogRows | Where-Object { $_.Action -eq 'MemberAdded' }).Count
    $mSkipped = @($script:LogRows | Where-Object { $_.Action -eq 'MemberSkipped' }).Count
    $errors   = @($script:LogRows | Where-Object { $_.Action -eq 'Error' }).Count

    Show-Section -Title 'Resumen'
    Write-KeyValue -Key 'Grupos creados'    -Value $created  -ValueColor Green
    Write-KeyValue -Key 'Grupos omitidos'   -Value $skipped  -ValueColor Yellow
    Write-KeyValue -Key 'Miembros añadidos' -Value $added    -ValueColor Green
    Write-KeyValue -Key 'Miembros existentes' -Value $mSkipped -ValueColor Yellow
    Write-KeyValue -Key 'Errores'           -Value $errors   -ValueColor Red
    Write-KeyValue -Key 'Log CSV'           -Value $LogCsvPath

    if ($script:WhatIfMode) {
        Show-WarningBlock -Title 'Modo WhatIf' -Detail 'No se realizaron cambios reales.'
    }
}

# --- ENTRYPOINT ---

Assert-RequiredServicesReady
Show-Header -Title 'GREX365' -Subtitle 'Creación de grupos/DL desde CSV'

$session = Start-LogSession -Name 'New-Groups'

try {
    Show-Section -Title 'Tipo de grupo'
    Write-Indent -Level 2; Write-Host '1   Microsoft 365 Group (Graph)' -ForegroundColor Gray
    Write-Indent -Level 2; Write-Host '2   Distribution List   (Exchange Online)' -ForegroundColor Gray
    do {
        $raw = Read-Input -Prompt 'Selección (1 o 2)'
    } while ((Normalize-Input -Value $raw) -notin '1','2')
    $script:GroupType = if ((Normalize-Input -Value $raw) -eq '1') { 'M365' } else { 'DL' }
    Write-Log ('Tipo: ' + $script:GroupType) -Source 'NewGroups'

    if (-not $Domain) {
        $dom = Read-Input -Prompt 'Dominio para email del grupo (ej: contoso.com)'
        $dom = Normalize-Input -Value $dom
        if (-not $dom) { throw 'Dominio requerido.' }
        $script:Domain = $dom.TrimStart('@')
    } else {
        $script:Domain = $Domain.TrimStart('@')
    }
    Write-Log ('Dominio: @' + $script:Domain) -Source 'NewGroups'

    $allFolders = New-Object System.Collections.Generic.List[string]
    if ($FolderPaths -and $FolderPaths.Count -gt 0) {
        foreach ($fp in $FolderPaths) { [void]$allFolders.Add((Normalize-Input -Value $fp)) }
    } else {
        Show-Section -Title 'Carpetas con CSVs'
        Write-Indent -Level 2
        Write-Host 'Introduce rutas (una por línea). Línea vacía para terminar.' -ForegroundColor DarkGray
        if (Get-Command Show-CsvFormatHint -ErrorAction SilentlyContinue) { Show-CsvFormatHint }
        $idx = 1
        while ($true) {
            $raw = Read-Input -Prompt ("Carpeta $idx")
            $val = Normalize-Input -Value $raw
            if (-not $val) { break }
            if (-not (Test-Path -LiteralPath $val -PathType Container)) {
                Show-WarningBlock -Title 'Ruta no es carpeta' -Detail $val
                continue
            }
            [void]$allFolders.Add($val); $idx++
        }
    }

    if ($allFolders.Count -eq 0) { throw 'Ninguna carpeta válida. Operación cancelada.' }

    $allCsvs = New-Object System.Collections.Generic.List[string]
    foreach ($folder in $allFolders) {
        $found = @(Get-ChildItem -LiteralPath $folder -Filter '*.csv' -File | Where-Object { $_.Name -notmatch '^Log_' })
        if ($found.Count -eq 0) { Write-Log ('Sin CSVs en: ' + $folder) -Level WARN -Source 'NewGroups' }
        else {
            foreach ($f in $found) { [void]$allCsvs.Add($f.FullName) }
            Write-Log ($found.Count.ToString() + ' CSV(s) en: ' + $folder) -Source 'NewGroups'
        }
    }

    if ($allCsvs.Count -eq 0) { throw 'No hay CSVs para procesar.' }
    Write-Log ('Total CSVs: ' + $allCsvs.Count) -Source 'NewGroups'

    foreach ($csvPath in $allCsvs) { Invoke-ProcessCsv -CsvPath $csvPath }

    $logFolder = Split-Path $allCsvs[0] -Parent
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $logCsvPath = Join-Path $logFolder ("new_groups_log_$stamp.csv")
    if ($script:LogRows.Count -gt 0) {
        $script:LogRows | Export-Csv -Path $logCsvPath -NoTypeInformation -Encoding UTF8
        Write-Log ('Log CSV: ' + $logCsvPath) -Level OK -Source 'NewGroups'
    }

    Show-Summary -LogCsvPath $logCsvPath

    if ($script:CreatedLinks.Count -gt 0) {
        Show-Section -Title ('Grupos creados · ' + $script:CreatedLinks.Count)
        foreach ($link in $script:CreatedLinks) {
            Show-AdminLink -Type $link.Type -Id $link.Id -DisplayName ($link.DisplayName + ' · ' + $link.Email)
        }
        Write-Host ''
    }

    Stop-LogSession -Persist | Out-Null
} catch {
    Show-ErrorBlock -Title 'Operación fallida' -Detail $_.Exception.Message
    Stop-LogSession -Persist | Out-Null
}

