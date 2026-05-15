#requires -Version 7.4
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
try { $Host.UI.RawUI.WindowTitle = 'GREX365 · Agregar miembros a grupo/DL' } catch {}

# All helpers (Write-Log, Show-Header, CSV parser, Validation, Resolve-GroupBySearch)
# live in GREX365/Modules and are dot-sourced by Main.ps1 before invocation.

function Resolve-GraphIdentity {
    param(
        [string]$Email,
        [string]$Id,
        [hashtable]$Cache
    )

    $email = if ($null -ne $Email) { $Email.Trim() } else { '' }
    $id    = if ($null -ne $Id)    { $Id.Trim() }    else { '' }

    if ($email) {
        $k = 'EMAIL::' + $email.ToLowerInvariant()
        if ($Cache.ContainsKey($k)) { return $Cache[$k] }
    }
    if ($id) {
        $k = 'ID::' + $id.ToLowerInvariant()
        if ($Cache.ContainsKey($k)) { return $Cache[$k] }
    }

    try {
        if ($id -and (Test-Guid -Value $id)) {
            $u = Get-MgUser -UserId $id -Property Id,Mail,UserPrincipalName -ErrorAction Stop
            if ($u) {
                if (-not $email) {
                    $email = @($u.Mail, $u.UserPrincipalName) | Where-Object { $_ } | Select-Object -First 1
                }
                $r = [PSCustomObject]@{ Email = [string]$email; Id = [string]$u.Id; Source = 'GraphById' }
                if ($r.Email) { $Cache['EMAIL::' + $r.Email.ToLowerInvariant()] = $r }
                if ($r.Id)    { $Cache['ID::'    + $r.Id.ToLowerInvariant()]    = $r }
                return $r
            }
        }
    } catch {}

    try {
        if ($email) {
            $safe = $email.Replace("'", "''")
            $users = @(Get-MgUser -Filter "mail eq '$safe' or userPrincipalName eq '$safe'" -Property Id,Mail,UserPrincipalName -ConsistencyLevel eventual -All -ErrorAction Stop)
            $u = $users | Select-Object -First 1
            if ($u) {
                $finalEmail = @($u.Mail, $u.UserPrincipalName, $email) | Where-Object { $_ } | Select-Object -First 1
                $r = [PSCustomObject]@{ Email = [string]$finalEmail; Id = [string]$u.Id; Source = 'GraphByEmail' }
                if ($r.Email) { $Cache['EMAIL::' + $r.Email.ToLowerInvariant()] = $r }
                if ($r.Id)    { $Cache['ID::'    + $r.Id.ToLowerInvariant()]    = $r }
                return $r
            }
        }
    } catch {}

    $fallback = [PSCustomObject]@{ Email = [string]$email; Id = [string]$id; Source = 'Unchanged' }
    if ($email) { $Cache['EMAIL::' + $email.ToLowerInvariant()] = $fallback }
    if ($id)    { $Cache['ID::'    + $id.ToLowerInvariant()]    = $fallback }
    return $fallback
}

function Get-ExistingMemberMap {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('DistributionGroup','UnifiedGroup')][string]$GroupType,
        [Parameter(Mandatory = $true)][string]$Identity
    )

    $map = @{}
    if ($GroupType -eq 'DistributionGroup') {
        $members = @(Get-DistributionGroupMember -Identity $Identity -ResultSize Unlimited -ErrorAction Stop)
        foreach ($m in $members) {
            if ($m.PrimarySmtpAddress) {
                $smtp = $m.PrimarySmtpAddress.ToString().Trim()
                if ($smtp) { $map[$smtp.ToLowerInvariant()] = $true }
            }
        }
    } else {
        $members = @(Get-UnifiedGroupLinks -Identity $Identity -LinkType Members -ResultSize Unlimited -ErrorAction Stop)
        foreach ($m in $members) {
            if ($m.PrimarySmtpAddress) {
                $smtp = $m.PrimarySmtpAddress.ToString().Trim()
                if ($smtp) { $map[$smtp.ToLowerInvariant()] = $true }
            }
            elseif ($m.WindowsLiveID) {
                $smtp = $m.WindowsLiveID.ToString().Trim()
                if ($smtp) { $map[$smtp.ToLowerInvariant()] = $true }
            }
        }
    }
    return $map
}

function Add-MemberToTargetGroup {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('DistributionGroup','UnifiedGroup')][string]$GroupType,
        [Parameter(Mandatory = $true)][string]$Identity,
        [Parameter(Mandatory = $true)][string]$MemberSmtp
    )

    if ($GroupType -eq 'DistributionGroup') {
        Add-DistributionGroupMember -Identity $Identity -Member $MemberSmtp -ErrorAction Stop
    } else {
        Add-UnifiedGroupLinks -Identity $Identity -LinkType Members -Links $MemberSmtp -ErrorAction Stop
    }
}

function Resolve-RecipientByIdOrEmail {
    param([string]$Id, [string]$Email)

    $r = @()
    if ($Id) {
        $safe = $Id.Replace("'", "''")
        $r = @(Get-EXORecipient -ResultSize Unlimited -Filter "ExternalDirectoryObjectId -eq '$safe'" -ErrorAction SilentlyContinue)
    }
    if ((-not $r -or @($r).Count -eq 0) -and $Email) {
        try { $r = @(Get-EXORecipient -Identity $Email -ErrorAction Stop) } catch {}
    }
    if ($r -and @($r).Count -gt 0) { return @($r)[0] }
    return $null
}

function New-ResultObject {
    param(
        [string]$InputEmail, [string]$InputId,
        [string]$Email, [string]$Id, [string]$Estado, [string]$Detalle = '',
        [string]$RecipientType = '', [string]$PrimarySmtpAddress = '',
        [string]$DisplayName = '', [string]$ResolutionSource = ''
    )
    [PSCustomObject]@{
        InputEmail         = $InputEmail
        InputId            = $InputId
        ResolvedEmail      = $Email
        ResolvedId         = $Id
        Estado             = $Estado
        Detalle            = $Detalle
        DisplayName        = $DisplayName
        PrimarySmtpAddress = $PrimarySmtpAddress
        RecipientType      = $RecipientType
        ResolutionSource   = $ResolutionSource
    }
}

# --- ENTRYPOINT ---

Assert-RequiredServicesReady
Show-Header -Title 'GREX365' -Subtitle 'Agregar miembros a grupo/DL'

$session = Start-LogSession -Name 'Add-GroupMembers'

try {
    $InputCsv = Read-ValidatedCsvPath -Prompt 'Ruta completa del CSV de entrada'

    Show-Section -Title 'Selección de grupo destino'
    $group = Resolve-GroupBySearch -Prompt 'Correo, nombre o alias del grupo destino'
    $GroupEmail = $group.PrimarySmtpAddress

    $OutputFolder = Split-Path -Path $InputCsv -Parent
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $resultCsv = Join-Path $OutputFolder ("add_members_result_$stamp.csv")

    Show-Section -Title 'Operación'
    Write-KeyValue -Key 'CSV entrada'  -Value $InputCsv
    Write-KeyValue -Key 'Carpeta out'  -Value $OutputFolder
    Write-KeyValue -Key 'Grupo'        -Value ("{0} <{1}>" -f $group.DisplayName, $GroupEmail)
    Write-KeyValue -Key 'Tipo'         -Value $group.GroupType
    Write-KeyValue -Key 'Resultado'    -Value $resultCsv

    Write-Log 'Leyendo CSV...' -Source 'Add-Members'
    $rows = @(Import-FlexibleCsv -Path $InputCsv -Schema EmailId)
    Write-Log ("Filas detectadas: {0}" -f $rows.Count) -Level OK -Source 'Add-Members'

    Write-Log 'Cargando miembros actuales del grupo...' -Source 'Add-Members'
    $existingMap = Get-ExistingMemberMap -GroupType $group.GroupType -Identity $group.Identity
    Write-Log ("Miembros actuales: {0}" -f $existingMap.Count) -Level OK -Source 'Add-Members'

    $results     = New-Object System.Collections.Generic.List[object]
    $noAdded     = New-Object System.Collections.Generic.List[object]
    $seenIds     = @{}
    $graphCache  = @{}

    $counters = @{
        Procesados     = 0
        Agregados      = 0
        SinExchange    = 0
        YaExistian     = 0
        NoResueltos    = 0
        IdNoResueltos  = 0
        GuidInvalidos  = 0
        Duplicados     = 0
        SinSmtp        = 0
        Vacios         = 0
        Errores        = 0
    }

    $currentIdx = 0
    foreach ($row in $rows) {
        $baseEmail = if ($row.PSObject.Properties.Name -contains 'Email') { [string]$row.Email } else { '' }
        $baseId    = if ($row.PSObject.Properties.Name -contains 'Id')    { [string]$row.Id }    else { '' }

        $expanded = @(Split-MultiEmailValue -Value $baseEmail)
        if ($expanded.Count -eq 0) { $expanded = @('') }

        foreach ($emailItem in $expanded) {
            $currentIdx++
            $counters.Procesados++

            $inputEmail = [string]$emailItem
            $inputId = [string]$baseId
            $email = $inputEmail.Trim()
            $id = $inputId.Trim()
            $source = ''

            $pct = if ($rows.Count -gt 0) { [math]::Round(($currentIdx / $rows.Count) * 100, 2) } else { 0 }
            Write-Progress -Activity 'Procesando miembros' -Status "Fila $currentIdx" -PercentComplete $pct

            if (-not $email -and -not $id) {
                $counters.Vacios++
                $r = New-ResultObject -InputEmail $inputEmail -InputId $inputId -Email $email -Id $id -Estado 'REGISTRO_VACIO' -Detalle 'Fila sin datos útiles'
                $results.Add($r); $noAdded.Add($r)
                Write-Log "Fila $currentIdx vacía — omitida" -Level WARN -Source 'Add-Members'
                continue
            }

            if (($email -and -not $id) -or ($id -and -not $email)) {
                $resolved = Resolve-GraphIdentity -Email $email -Id $id -Cache $graphCache
                $email = [string]$resolved.Email
                $id = [string]$resolved.Id
                $source = [string]$resolved.Source
            }

            if (-not $id) {
                $counters.IdNoResueltos++
                $r = New-ResultObject -InputEmail $inputEmail -InputId $inputId -Email $email -Id $id -Estado 'ID_NO_RESUELTO' -Detalle 'No se pudo resolver Id vía Graph' -ResolutionSource $source
                $results.Add($r); $noAdded.Add($r)
                Write-Log ("Id no resuelto: {0}" -f $email) -Level WARN -Source 'Add-Members'
                continue
            }

            if (-not (Test-Guid -Value $id)) {
                $counters.GuidInvalidos++
                $r = New-ResultObject -InputEmail $inputEmail -InputId $inputId -Email $email -Id $id -Estado 'ID_GUID_INVALIDO' -Detalle 'Formato GUID inválido' -ResolutionSource $source
                $results.Add($r); $noAdded.Add($r)
                Write-Log ("GUID inválido: {0}" -f $id) -Level WARN -Source 'Add-Members'
                continue
            }

            if ($seenIds.ContainsKey($id.ToLowerInvariant())) {
                $counters.Duplicados++
                $r = New-ResultObject -InputEmail $inputEmail -InputId $inputId -Email $email -Id $id -Estado 'DUPLICADO_EN_CSV' -Detalle 'Id duplicado en CSV' -ResolutionSource $source
                $results.Add($r); $noAdded.Add($r)
                Write-Log ("Duplicado CSV: {0}" -f $id) -Level WARN -Source 'Add-Members'
                continue
            }
            $seenIds[$id.ToLowerInvariant()] = $true

            try {
                $rcpt = Resolve-RecipientByIdOrEmail -Id $id -Email $email

                if (-not $rcpt) {
                    $fallback = if ($email) { $email } else { $id }
                    $key = $fallback.ToLowerInvariant()

                    if ($existingMap.ContainsKey($key)) {
                        $counters.YaExistian++
                        $r = New-ResultObject -InputEmail $inputEmail -InputId $inputId -Email $email -Id $id -Estado 'YA_EXISTE_EN_GRUPO' -Detalle 'Ya pertenece al grupo (sin recipient en Exchange)' -ResolutionSource $source
                        $results.Add($r); $noAdded.Add($r)
                        continue
                    }

                    try {
                        Add-MemberToTargetGroup -GroupType $group.GroupType -Identity $group.Identity -MemberSmtp $fallback
                        $existingMap[$key] = $true
                        $counters.Agregados++
                        $counters.SinExchange++
                        $r = New-ResultObject -InputEmail $inputEmail -InputId $inputId -Email $email -Id $id -Estado 'AGREGADO_SIN_EXCHANGE' -Detalle 'Añadido vía Graph email/id sin recipient EXO' -ResolutionSource $source
                        $results.Add($r)
                        Write-Log ("Agregado sin EXO: {0}" -f $fallback) -Level OK -Source 'Add-Members'
                    } catch {
                        $counters.NoResueltos++
                        $r = New-ResultObject -InputEmail $inputEmail -InputId $inputId -Email $email -Id $id -Estado 'NO_RESUELTO_EN_EXCHANGE' -Detalle ('No resuelto en Exchange. ' + $_.Exception.Message) -ResolutionSource $source
                        $results.Add($r); $noAdded.Add($r)
                        Write-Log ("Fallback falló: {0}" -f $fallback) -Level WARN -Source 'Add-Members'
                    }
                    continue
                }

                $rcptSmtp = ''
                if ($rcpt.PrimarySmtpAddress) { $rcptSmtp = $rcpt.PrimarySmtpAddress.ToString().Trim() }

                if (-not $rcptSmtp) {
                    $fallback = if ($email) { $email } else { $id }
                    $key = $fallback.ToLowerInvariant()

                    if ($existingMap.ContainsKey($key)) {
                        $counters.YaExistian++
                        $r = New-ResultObject -InputEmail $inputEmail -InputId $inputId -Email $email -Id $id -Estado 'YA_EXISTE_EN_GRUPO' -Detalle 'Ya pertenece (sin SMTP en EXO)' -RecipientType ([string]$rcpt.RecipientTypeDetails) -DisplayName ([string]$rcpt.DisplayName) -ResolutionSource $source
                        $results.Add($r); $noAdded.Add($r)
                        continue
                    }

                    try {
                        Add-MemberToTargetGroup -GroupType $group.GroupType -Identity $group.Identity -MemberSmtp $fallback
                        $existingMap[$key] = $true
                        $counters.Agregados++
                        $counters.SinExchange++
                        $r = New-ResultObject -InputEmail $inputEmail -InputId $inputId -Email $email -Id $id -Estado 'AGREGADO_SIN_SMTP' -Detalle 'Añadido vía Graph sin SMTP EXO' -RecipientType ([string]$rcpt.RecipientTypeDetails) -DisplayName ([string]$rcpt.DisplayName) -ResolutionSource $source
                        $results.Add($r)
                        Write-Log ("Agregado sin SMTP: {0}" -f $fallback) -Level OK -Source 'Add-Members'
                    } catch {
                        $counters.SinSmtp++
                        $r = New-ResultObject -InputEmail $inputEmail -InputId $inputId -Email $email -Id $id -Estado 'RECIPIENT_SIN_SMTP' -Detalle ('Recipient existe sin SMTP. ' + $_.Exception.Message) -RecipientType ([string]$rcpt.RecipientTypeDetails) -DisplayName ([string]$rcpt.DisplayName) -ResolutionSource $source
                        $results.Add($r); $noAdded.Add($r)
                        Write-Log ("Recipient sin SMTP: {0}" -f $fallback) -Level WARN -Source 'Add-Members'
                    }
                    continue
                }

                $rcptKey = $rcptSmtp.ToLowerInvariant()
                if ($existingMap.ContainsKey($rcptKey)) {
                    $counters.YaExistian++
                    $r = New-ResultObject -InputEmail $inputEmail -InputId $inputId -Email $email -Id $id -Estado 'YA_EXISTE_EN_GRUPO' -Detalle 'Ya pertenece al grupo' -RecipientType ([string]$rcpt.RecipientTypeDetails) -PrimarySmtpAddress $rcptSmtp -DisplayName ([string]$rcpt.DisplayName) -ResolutionSource $source
                    $results.Add($r); $noAdded.Add($r)
                    continue
                }

                Add-MemberToTargetGroup -GroupType $group.GroupType -Identity $group.Identity -MemberSmtp $rcptSmtp
                $existingMap[$rcptKey] = $true
                $counters.Agregados++
                $r = New-ResultObject -InputEmail $inputEmail -InputId $inputId -Email $email -Id $id -Estado 'AGREGADO' -Detalle 'Añadido al grupo' -RecipientType ([string]$rcpt.RecipientTypeDetails) -PrimarySmtpAddress $rcptSmtp -DisplayName ([string]$rcpt.DisplayName) -ResolutionSource $source
                $results.Add($r)
                Write-Log ("Agregado: {0}" -f $rcptSmtp) -Level OK -Source 'Add-Members'

            } catch {
                $counters.Errores++
                $r = New-ResultObject -InputEmail $inputEmail -InputId $inputId -Email $email -Id $id -Estado 'ERROR_AL_AGREGAR' -Detalle $_.Exception.Message -ResolutionSource $source
                $results.Add($r); $noAdded.Add($r)
                Write-Log ("Error agregando {0}: {1}" -f $email, $_.Exception.Message) -Level ERROR -Source 'Add-Members'
            }
        }
    }

    Write-Progress -Activity 'Procesando miembros' -Completed

    if ($results.Count -gt 0) {
        $results | Export-Csv -Path $resultCsv -NoTypeInformation -Encoding UTF8
        Write-Log ('CSV de resultados: ' + $resultCsv) -Level OK -Source 'Add-Members'
    }

    Show-Section -Title 'Resumen'
    Write-KeyValue -Key 'Grupo'           -Value ($group.DisplayName + ' <' + $GroupEmail + '>')
    Write-KeyValue -Key 'Tipo'            -Value $group.GroupType
    Write-KeyValue -Key 'Procesados'      -Value $counters.Procesados
    Write-KeyValue -Key 'Agregados'       -Value ('{0} ({1} sin EXO)' -f $counters.Agregados, $counters.SinExchange) -ValueColor Green
    Write-KeyValue -Key 'Ya existían'     -Value $counters.YaExistian -ValueColor Yellow
    Write-KeyValue -Key 'No resueltos EXO'-Value $counters.NoResueltos -ValueColor Yellow
    Write-KeyValue -Key 'Id no resueltos' -Value $counters.IdNoResueltos -ValueColor Yellow
    Write-KeyValue -Key 'GUID inválidos'  -Value $counters.GuidInvalidos -ValueColor Yellow
    Write-KeyValue -Key 'Duplicados CSV'  -Value $counters.Duplicados -ValueColor Yellow
    Write-KeyValue -Key 'Sin SMTP'        -Value $counters.SinSmtp -ValueColor Yellow
    Write-KeyValue -Key 'Registros vacíos'-Value $counters.Vacios -ValueColor Yellow
    Write-KeyValue -Key 'Errores'         -Value $counters.Errores -ValueColor Red
    Write-KeyValue -Key 'CSV resultado'   -Value $resultCsv

    if ($noAdded.Count -gt 0) {
        Show-Section -Title ('No añadidos (' + $noAdded.Count + ')')
        foreach ($item in $noAdded) {
            Write-Indent -Level 2
            Write-Host ('{0,-32} {1,-40} {2}' -f $item.InputEmail, $item.ResolvedEmail, $item.Estado) -ForegroundColor Yellow
        }
    }

    $logFile = Stop-LogSession -Persist
    if ($logFile) { Write-KeyValue -Key 'Log'             -Value $logFile }

    if ($counters.Agregados -gt 0 -and $group.Id) {
        Write-Host ''
        Show-AdminLink -Type $group.GroupType -Id $group.Id -DisplayName $group.DisplayName
    }

} catch {
    Show-ErrorBlock -Title 'Operación fallida' -Detail $_.Exception.Message
    Stop-LogSession -Persist | Out-Null
    throw
}

