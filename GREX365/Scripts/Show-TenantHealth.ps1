#requires -Version 7.4
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'
try { $Host.UI.RawUI.WindowTitle = 'GREX365 · Tenant health' } catch {}

# Read-only tenant snapshot: licenses, mailbox quotas, MFA coverage, stale users,
# privileged role holders, app secrets expiring. Pull on-demand, no caching across runs.

Assert-RequiredServicesReady

Show-Header -Title 'GREX365' -Subtitle 'Tenant Health Dashboard'

$session = Start-LogSession -Name 'TenantHealth'
$started = Get-Date

$metrics = [ordered]@{}
$items   = New-Object System.Collections.Generic.List[object]

function Add-HealthItem {
    param([string]$Category,[string]$Item,[string]$Value,[string]$Estado = 'OK',[string]$Detalle = '')
    $items.Add([PSCustomObject]@{
        Categoria = $Category
        Item      = $Item
        Valor     = $Value
        Estado    = $Estado
        Detalle   = $Detalle
    })
}

try {
    # --- Licencias ---
    # Map SKU part numbers to friendly product names (Microsoft public list).
    $skuFriendly = @{
        'ENTERPRISEPACK'              = 'Office 365 E3'
        'ENTERPRISEPREMIUM'           = 'Office 365 E5'
        'ENTERPRISEPACK_USGOV_DOD'    = 'Office 365 E3 (DOD)'
        'STANDARDPACK'                = 'Office 365 E1'
        'SPE_E3'                      = 'Microsoft 365 E3'
        'SPE_E5'                      = 'Microsoft 365 E5'
        'SPB'                         = 'Microsoft 365 Business Premium'
        'O365_BUSINESS_PREMIUM'       = 'Microsoft 365 Business Standard'
        'O365_BUSINESS_ESSENTIALS'    = 'Microsoft 365 Business Basic'
        'EXCHANGESTANDARD'            = 'Exchange Online Plan 1'
        'EXCHANGEENTERPRISE'          = 'Exchange Online Plan 2'
        'EXCHANGEDESKLESS'            = 'Exchange Online Kiosk'
        'POWER_BI_PRO'                = 'Power BI Pro'
        'POWER_BI_STANDARD'           = 'Power BI (free)'
        'PROJECTPROFESSIONAL'         = 'Project Plan 3'
        'PROJECTPREMIUM'              = 'Project Plan 5'
        'VISIOCLIENT'                 = 'Visio Plan 2'
        'VISIOONLINE_PLAN1'           = 'Visio Plan 1'
        'AAD_PREMIUM'                 = 'Entra ID P1'
        'AAD_PREMIUM_P2'              = 'Entra ID P2'
        'EMS'                         = 'Enterprise Mobility + Security E3'
        'EMSPREMIUM'                  = 'Enterprise Mobility + Security E5'
        'INTUNE_A'                    = 'Intune'
        'TEAMS_EXPLORATORY'           = 'Teams Exploratory'
        'FLOW_FREE'                   = 'Power Automate Free'
        'POWERAUTOMATE_ATTENDED_RPA'  = 'Power Automate per user'
        'MCOEV'                       = 'Teams Phone Standard'
        'MCOMEETADV'                  = 'Audio Conferencing'
        'WIN10_PRO_ENT_SUB'           = 'Windows 10/11 Enterprise E3'
        'WIN_ENT_E5'                  = 'Windows 10/11 Enterprise E5'
    }
    function Get-SkuFriendly { param([string]$Sku); if ($skuFriendly.ContainsKey($Sku)) { return $skuFriendly[$Sku] }; return $Sku }

    Show-Section -Title 'Licencias por SKU'
    $totalAssigned = 0; $totalEnabled = 0; $skusOver = 0; $skusNearFull = 0
    try {
        $skus = @(Invoke-WithRetry -OperationName 'Get-MgSubscribedSku' -ScriptBlock { Get-MgSubscribedSku -All -ErrorAction Stop })
        $skusSorted = @($skus | Sort-Object SkuPartNumber)
        foreach ($sku in $skusSorted) {
            $partNo   = [string]$sku.SkuPartNumber
            $friendly = Get-SkuFriendly -Sku $partNo
            $assigned = [int]$sku.ConsumedUnits
            $enabled  = if ($sku.PrepaidUnits) { [int]$sku.PrepaidUnits.Enabled } else { 0 }
            $warning  = if ($sku.PrepaidUnits) { [int]$sku.PrepaidUnits.Warning } else { 0 }
            $suspended = if ($sku.PrepaidUnits) { [int]$sku.PrepaidUnits.Suspended } else { 0 }
            $totalAssigned += $assigned
            $totalEnabled  += $enabled
            $free = $enabled - $assigned
            $pct  = if ($enabled -gt 0) { [math]::Round(($assigned / $enabled) * 100, 1) } else { 0 }

            $estado = 'OK'
            if ($enabled -gt 0 -and $assigned -gt $enabled) { $estado = 'ERROR'; $skusOver++ }
            elseif ($pct -ge 95)                            { $estado = 'WARN';  $skusNearFull++ }
            elseif ($warning -gt 0)                         { $estado = 'WARN' }

            $detail = "asignadas=$assigned · disponibles=$enabled · libres=$free · uso=${pct}%"
            if ($warning -gt 0)   { $detail += " · warning=$warning" }
            if ($suspended -gt 0) { $detail += " · suspended=$suspended" }
            Add-HealthItem -Category 'Licencias' -Item ("$friendly  ($partNo)") -Value "$assigned/$enabled" -Estado $estado -Detalle $detail
            Write-KeyValue -Key $friendly.PadRight(40) -Value ("{0,4} / {1,-4}  {2}%" -f $assigned, $enabled, $pct)
        }
        $metrics['SKUs activos']              = $skus.Count
        $metrics['Licencias asignadas']       = "$totalAssigned / $totalEnabled"
        $metrics['SKUs sobreasignados']       = $skusOver
        $metrics['SKUs >=95% uso']            = $skusNearFull

        # Recommendations
        if ($skusOver -gt 0)     { Add-HealthItem -Category 'Recomendación' -Item 'Sobreasignación de licencias' -Value "$skusOver SKUs" -Estado 'ERROR' -Detalle 'Comprar más licencias o reasignar' }
        if ($skusNearFull -gt 0) { Add-HealthItem -Category 'Recomendación' -Item 'SKUs cerca del límite' -Value "$skusNearFull SKUs" -Estado 'WARN' -Detalle 'Revisar consumo y compra anticipada' }
    } catch {
        Write-Log "No se pudieron leer SKUs: $($_.Exception.Message)" -Level WARN -Source 'TenantHealth'
        Add-HealthItem -Category 'Licencias' -Item 'consulta' -Value 'fallo' -Estado 'ERROR' -Detalle $_.Exception.Message
    }

    # Users without license vs disabled-with-license
    Show-Section -Title 'Asignación de licencias por usuario'
    try {
        $allUsers = @(Invoke-WithRetry -OperationName 'Get-MgUser license-check' -ScriptBlock {
            Get-MgUser -All -ConsistencyLevel eventual -Property Id,UserPrincipalName,DisplayName,AccountEnabled,AssignedLicenses,UserType -ErrorAction Stop
        })
        $noLicMembers   = @($allUsers | Where-Object { $_.AccountEnabled -and $_.UserType -ne 'Guest' -and (-not $_.AssignedLicenses -or @($_.AssignedLicenses).Count -eq 0) })
        $disabledLic    = @($allUsers | Where-Object { -not $_.AccountEnabled -and $_.AssignedLicenses -and @($_.AssignedLicenses).Count -gt 0 })

        Write-KeyValue -Key 'Usuarios activos sin licencia' -Value $noLicMembers.Count -ValueColor Yellow
        Write-KeyValue -Key 'Deshabilitados con licencia'   -Value $disabledLic.Count -ValueColor Yellow
        $metrics['Activos sin licencia']     = $noLicMembers.Count
        $metrics['Deshabilitados con lic.']  = $disabledLic.Count

        foreach ($u in $noLicMembers | Select-Object -First 20) {
            Add-HealthItem -Category 'Sin licencia' -Item $u.UserPrincipalName -Value 'activo sin licencia' -Estado 'WARN' -Detalle ([string]$u.DisplayName)
        }
        foreach ($u in $disabledLic | Select-Object -First 20) {
            $n = @($u.AssignedLicenses).Count
            Add-HealthItem -Category 'Disabled+License' -Item $u.UserPrincipalName -Value "$n licencias" -Estado 'WARN' -Detalle 'Liberar licencias bloqueadas en cuenta deshabilitada'
        }
        if ($disabledLic.Count -gt 0) {
            Add-HealthItem -Category 'Recomendación' -Item 'Liberar licencias en cuentas deshabilitadas' -Value "$($disabledLic.Count) cuentas" -Estado 'WARN' -Detalle 'Quitar licencias antes de archivar la cuenta'
        }
    } catch {
        Write-Log "Análisis de licencias fallido: $($_.Exception.Message)" -Level WARN -Source 'TenantHealth'
    }

    # --- Cuotas de buzón ---
    # Batched: pull mailbox list (lightweight) + per-mailbox stats with timeout + progress.
    # Robust parse for ByteQuantifiedSize.
    Show-Section -Title 'Buzones cerca de cuota'
    function ConvertTo-BytesSafe {
        param($Size)
        if ($null -eq $Size) { return 0 }
        try {
            if ($Size.PSObject.Properties.Name -contains 'Value' -and $Size.Value) {
                $v = $Size.Value
                if ($v.PSObject.Methods.Name -contains 'ToBytes') { return [int64]$v.ToBytes() }
            }
        } catch {}
        $s = [string]$Size
        if ($s -match '([0-9.,]+)\s*([KMGT]?B)') {
            $num = [double](($matches[1] -replace ',', ''))
            $unit = $matches[2]
            $mult = switch ($unit) { 'B' { 1 } 'KB' { 1KB } 'MB' { 1MB } 'GB' { 1GB } 'TB' { 1TB } default { 1 } }
            return [int64]($num * $mult)
        }
        return 0
    }

    try {
        $mbxList = @()
        try {
            $mbxList = @(Invoke-WithRetry -OperationName 'Get-EXOMailbox list' -ScriptBlock {
                Get-EXOMailbox -ResultSize 250 -Properties UserPrincipalName,DisplayName,ProhibitSendQuota -ErrorAction Stop
            })
        } catch {
            Write-Log "No se pudieron listar buzones: $($_.Exception.Message)" -Level WARN -Source 'TenantHealth'
        }

        $near = @()
        $i = 0
        $total = $mbxList.Count
        foreach ($mbx in $mbxList) {
            $i++
            $pct = if ($total -gt 0) { [math]::Round(($i / $total) * 100, 1) } else { 0 }
            Write-Progress -Activity 'Calculando cuotas' -Status "$i / $total" -PercentComplete $pct -Id 10
            try {
                $quotaB = if ($mbx.ProhibitSendQuota -and "$($mbx.ProhibitSendQuota)" -ne 'Unlimited') { ConvertTo-BytesSafe -Size $mbx.ProhibitSendQuota } else { 0 }
                if ($quotaB -le 0) { continue }
                $stats = $null
                try {
                    $stats = Get-EXOMailboxStatistics -Identity $mbx.UserPrincipalName -Properties TotalItemSize -ErrorAction Stop
                } catch { continue }
                $usedB = ConvertTo-BytesSafe -Size $stats.TotalItemSize
                if ($usedB -le 0) { continue }
                $ratio = $usedB / $quotaB
                if ($ratio -gt 0.85) {
                    $near += [PSCustomObject]@{ Upn=$mbx.UserPrincipalName; DisplayName=$mbx.DisplayName; Ratio=$ratio }
                }
            } catch {}
        }
        Write-Progress -Activity 'Calculando cuotas' -Completed -Id 10

        $metrics['Buzones >85% cuota'] = $near.Count
        Write-KeyValue -Key 'Buzones revisados'    -Value $total
        Write-KeyValue -Key 'Cerca de cuota (>85%)' -Value $near.Count
        foreach ($m in $near | Select-Object -First 25) {
            $pctTxt = '{0:N1}%' -f ($m.Ratio * 100)
            Add-HealthItem -Category 'Cuotas' -Item $m.Upn -Value $pctTxt -Estado 'WARN' -Detalle $m.DisplayName
        }
    } catch {
        Write-Log "Cuotas no disponibles: $($_.Exception.Message)" -Level WARN -Source 'TenantHealth'
    }

    # --- MFA gap ---
    Show-Section -Title 'MFA coverage'
    try {
        $regs = @(Invoke-WithRetry -OperationName 'Get-MgReportAuthMethodUserRegistrationDetail' -ScriptBlock {
            Get-MgReportAuthenticationMethodUserRegistrationDetail -All -ErrorAction Stop
        })
        $noMfa = @($regs | Where-Object { -not $_.IsMfaRegistered -and $_.UserType -eq 'member' })
        $metrics['Usuarios sin MFA'] = $noMfa.Count
        Write-KeyValue -Key 'Sin MFA (members)' -Value $noMfa.Count
        foreach ($u in $noMfa | Select-Object -First 25) {
            Add-HealthItem -Category 'MFA gap' -Item $u.UserPrincipalName -Value 'sin registrar' -Estado 'WARN' -Detalle ([string]$u.UserDisplayName)
        }
    } catch {
        Write-Log "Reporte MFA no disponible: $($_.Exception.Message)" -Level WARN -Source 'TenantHealth'
    }

    # --- Privileged roles ---
    Show-Section -Title 'Roles privilegiados'
    try {
        $roles = @('Global Administrator','Exchange Administrator','Privileged Role Administrator','User Administrator','Security Administrator')
        $totalPriv = 0
        foreach ($roleName in $roles) {
            try {
                $r = Invoke-WithRetry -OperationName "Get-MgDirectoryRoleByDisplayName" -Quiet -ScriptBlock {
                    Get-MgDirectoryRole -Filter "displayName eq '$roleName'" -ErrorAction Stop | Select-Object -First 1
                }
                if (-not $r) { continue }
                $members = @(Invoke-WithRetry -OperationName 'Get-MgDirectoryRoleMember' -Quiet -ScriptBlock {
                    Get-MgDirectoryRoleMember -DirectoryRoleId $r.Id -All -ErrorAction Stop
                })
                $totalPriv += $members.Count
                $estado = if ($roleName -eq 'Global Administrator' -and $members.Count -gt 4) { 'WARN' } else { 'OK' }
                Add-HealthItem -Category 'Roles' -Item $roleName -Value ("{0} miembros" -f $members.Count) -Estado $estado
            } catch {}
        }
        $metrics['Asignaciones privilegiadas'] = $totalPriv
    } catch {
        Write-Log "Roles no disponibles: $($_.Exception.Message)" -Level WARN -Source 'TenantHealth'
    }

    # --- App secrets expiring ---
    Show-Section -Title 'App secrets caducando (<30d)'
    try {
        $apps = @(Invoke-WithRetry -OperationName 'Get-MgApplication' -ScriptBlock {
            Get-MgApplication -All -Property Id,DisplayName,AppId,PasswordCredentials,KeyCredentials -ErrorAction Stop
        })
        $cutoff = (Get-Date).AddDays(30)
        $atRisk = 0
        foreach ($app in $apps) {
            $creds = @()
            if ($app.PasswordCredentials) { $creds += @($app.PasswordCredentials | ForEach-Object { [PSCustomObject]@{ Type='secret'; EndDate=$_.EndDateTime } }) }
            if ($app.KeyCredentials)      { $creds += @($app.KeyCredentials      | ForEach-Object { [PSCustomObject]@{ Type='key';    EndDate=$_.EndDateTime } }) }
            foreach ($c in $creds) {
                if (-not $c.EndDate) { continue }
                if ($c.EndDate -lt (Get-Date)) { continue }
                if ($c.EndDate -le $cutoff) {
                    $atRisk++
                    $days = [math]::Round(($c.EndDate - (Get-Date)).TotalDays, 0)
                    Add-HealthItem -Category 'App secrets' -Item $app.DisplayName -Value "$($c.Type) caduca en ${days}d" -Estado 'WARN' -Detalle ([string]$app.AppId)
                }
            }
        }
        $metrics['App secrets caducando'] = $atRisk
        Write-KeyValue -Key 'Apps revisadas' -Value $apps.Count
        Write-KeyValue -Key 'Secrets <30d' -Value $atRisk
    } catch {
        Write-Log "App secrets no disponibles: $($_.Exception.Message)" -Level WARN -Source 'TenantHealth'
    }

    # --- Stale users ---
    Show-Section -Title 'Usuarios inactivos'
    try {
        $stale = @(Invoke-WithRetry -OperationName 'Get-MgUser stale' -ScriptBlock {
            Get-MgUser -All -Property Id,UserPrincipalName,AccountEnabled,SignInActivity -ConsistencyLevel eventual -ErrorAction Stop
        })
        $cutoff = (Get-Date).AddDays(-90)
        $oldUsers = @($stale | Where-Object {
            $_.AccountEnabled -eq $true -and (
                ($_.SignInActivity -and $_.SignInActivity.LastSignInDateTime -and $_.SignInActivity.LastSignInDateTime -lt $cutoff) -or
                (-not $_.SignInActivity)
            )
        })
        $metrics['Usuarios inactivos >90d'] = $oldUsers.Count
        Write-KeyValue -Key 'Sin login >90d' -Value $oldUsers.Count
        foreach ($u in $oldUsers | Select-Object -First 20) {
            $last = if ($u.SignInActivity -and $u.SignInActivity.LastSignInDateTime) { $u.SignInActivity.LastSignInDateTime.ToString('yyyy-MM-dd') } else { 'nunca' }
            Add-HealthItem -Category 'Stale' -Item $u.UserPrincipalName -Value $last -Estado 'WARN'
        }
    } catch {
        Write-Log "Usuarios stale no disponibles: $($_.Exception.Message)" -Level WARN -Source 'TenantHealth'
    }

    # --- Service health (Graph Service announcements + last issues) ---
    Show-Section -Title 'Salud de servicios M365'
    $criticalSvc = @('Exchange Online','Microsoft Teams','SharePoint Online','OneDrive for Business','Microsoft 365 suite','Microsoft Entra','Identity Service','Office for the web')
    try {
        $health = @(Invoke-WithRetry -OperationName 'Get-MgServiceAnnouncementHealthOverview' -ScriptBlock {
            Get-MgServiceAnnouncementHealthOverview -All -ErrorAction Stop
        })
        $degraded = @($health | Where-Object { $_.Status -ne 'serviceOperational' })
        $metrics['Servicios totales']     = $health.Count
        $metrics['Servicios degradados']  = $degraded.Count
        Write-KeyValue -Key 'Servicios totales' -Value $health.Count
        Write-KeyValue -Key 'Degradados'        -Value $degraded.Count -ValueColor $(if ($degraded.Count -gt 0) { 'Red' } else { 'Green' })

        # Always surface critical workloads first
        foreach ($svc in $criticalSvc) {
            $h = $health | Where-Object { $_.Service -eq $svc } | Select-Object -First 1
            if (-not $h) { continue }
            $estado = switch -Wildcard ([string]$h.Status) {
                'serviceOperational' { 'OK' }
                'investigating'      { 'WARN' }
                'serviceRestored'    { 'OK' }
                'restoringService'   { 'WARN' }
                'verifyingService'   { 'WARN' }
                'serviceDegradation' { 'ERROR' }
                'serviceInterruption'{ 'ERROR' }
                default              { 'WARN' }
            }
            Add-HealthItem -Category 'Service health' -Item ([string]$h.Service) -Value ([string]$h.Status) -Estado $estado
        }
        # Other degraded services
        foreach ($s in $degraded) {
            if ($s.Service -in $criticalSvc) { continue }
            Add-HealthItem -Category 'Service health' -Item ([string]$s.Service) -Value ([string]$s.Status) -Estado 'WARN'
        }
    } catch {
        Write-Log "Service health no disponible: $($_.Exception.Message)" -Level WARN -Source 'TenantHealth'
    }

    # Last issues / advisories
    Show-Section -Title 'Avisos recientes (últimos 7 días)'
    try {
        $issues = @(Invoke-WithRetry -OperationName 'Get-MgServiceAnnouncementIssue' -ScriptBlock {
            Get-MgServiceAnnouncementIssue -All -ErrorAction Stop
        })
        $recent = @($issues | Where-Object { $_.LastModifiedDateTime -and $_.LastModifiedDateTime -gt (Get-Date).AddDays(-7) } | Sort-Object LastModifiedDateTime -Descending)
        $metrics['Avisos 7d'] = $recent.Count
        Write-KeyValue -Key 'Avisos últimos 7d' -Value $recent.Count
        foreach ($iss in $recent | Select-Object -First 10) {
            $sev = switch ([string]$iss.Classification) {
                'incident' { 'ERROR' }
                'advisory' { 'WARN' }
                default    { 'WARN' }
            }
            Add-HealthItem -Category 'Aviso' -Item ([string]$iss.Title) -Value ([string]$iss.Service) -Estado $sev -Detalle ("status=$($iss.Status) · " + $iss.LastModifiedDateTime.ToString('yyyy-MM-dd HH:mm'))
        }
    } catch {
        Write-Log "Issues recientes no disponibles: $($_.Exception.Message)" -Level WARN -Source 'TenantHealth'
    }

    # Directory sync errors (only relevant for hybrid)
    Show-Section -Title 'Errores de sincronización (Entra)'
    try {
        $org = Invoke-WithRetry -OperationName 'Get-MgOrganization' -ScriptBlock {
            Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
        }
        $syncEnabled = $false
        try { $syncEnabled = [bool]$org.OnPremisesSyncEnabled } catch {}
        Write-KeyValue -Key 'OnPremises Sync' -Value $syncEnabled
        if ($syncEnabled) {
            $lastSync = $null
            try { $lastSync = $org.OnPremisesLastSyncDateTime } catch {}
            Write-KeyValue -Key 'Última sync'   -Value ($lastSync ? $lastSync.ToString('yyyy-MM-dd HH:mm') : 'desconocida')
            if ($lastSync -and $lastSync -lt (Get-Date).AddHours(-2)) {
                Add-HealthItem -Category 'Sync' -Item 'Última sync >2h' -Value $lastSync.ToString('yyyy-MM-dd HH:mm') -Estado 'WARN' -Detalle 'Verificar AAD Connect'
            }
        } else {
            $metrics['On-prem sync'] = 'no'
        }
    } catch {
        Write-Log "Sync info no disponible: $($_.Exception.Message)" -Level WARN -Source 'TenantHealth'
    }

    $ended = Get-Date
    Show-Section -Title 'Snapshot'
    foreach ($k in $metrics.Keys) { Write-KeyValue -Key $k -Value $metrics[$k] }

    $logFile = Stop-LogSession -Persist
    $report = Publish-OperationReport `
        -Title 'Tenant Health · Snapshot' `
        -Operation 'Show-TenantHealth' `
        -StartTime $started -EndTime $ended `
        -Summary $metrics `
        -Items $items `
        -LogFile $logFile
    Write-KeyValue -Key 'Informe HTML' -Value $report

} catch {
    Show-ErrorBlock -Title 'Dashboard fallido' -Detail $_.Exception.Message
    Stop-LogSession -Persist | Out-Null
    throw
}

