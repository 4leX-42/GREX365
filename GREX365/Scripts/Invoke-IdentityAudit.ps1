#requires -Version 7.4
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'
try { $Host.UI.RawUI.WindowTitle = 'GREX365 · Identity audit' } catch {}

# Read-only identity audit. Flags:
#   - Stale members  (enabled, no sign-in >180d)
#   - Stale guests   (no sign-in >90d)
#   - Disabled users with mailbox / licenses still attached
#   - Groups without owner
#   - M365 groups with no members
#   - Distribution lists with no members

Assert-RequiredServicesReady

Show-Header -Title 'GREX365' -Subtitle 'Identity Audit'

$session = Start-LogSession -Name 'IdentityAudit'
$started = Get-Date
$findings = New-Object System.Collections.Generic.List[object]

function Add-Finding {
    param([string]$Category,[string]$Identity,[string]$Detail,[string]$Severity = 'WARN')
    $findings.Add([PSCustomObject]@{
        Categoria = $Category
        Identity  = $Identity
        Detalle   = $Detail
        Severidad = $Severity
    })
}

try {
    # --- All users for cross-checks ---
    $users = @()
    try {
        $users = @(Invoke-WithRetry -OperationName 'Get-MgUser audit' -ScriptBlock {
            Get-MgUser -All -ConsistencyLevel eventual `
                -Property Id,UserPrincipalName,DisplayName,AccountEnabled,UserType,AssignedLicenses,SignInActivity,Mail `
                -ErrorAction Stop
        })
    } catch {
        Write-Log "Get-MgUser fallo: $($_.Exception.Message)" -Level ERROR -Source 'IdentityAudit'
        throw
    }

    Show-Section -Title 'Análisis de usuarios'

    $cutoffMember = (Get-Date).AddDays(-180)
    $cutoffGuest  = (Get-Date).AddDays(-90)
    $staleMembers = 0; $staleGuests = 0; $disabledWithLicense = 0

    foreach ($u in $users) {
        $upn = [string]$u.UserPrincipalName
        $isGuest = ([string]$u.UserType -eq 'Guest')
        $lastSignIn = $null
        if ($u.SignInActivity -and $u.SignInActivity.LastSignInDateTime) {
            $lastSignIn = $u.SignInActivity.LastSignInDateTime
        }
        $licenseCount = if ($u.AssignedLicenses) { @($u.AssignedLicenses).Count } else { 0 }

        if (-not $u.AccountEnabled -and $licenseCount -gt 0) {
            $disabledWithLicense++
            Add-Finding -Category 'Disabled+License' -Identity $upn -Detail "Deshabilitado con $licenseCount licencias asignadas" -Severity 'WARN'
        }

        if ($u.AccountEnabled) {
            if ($isGuest) {
                if (-not $lastSignIn -or $lastSignIn -lt $cutoffGuest) {
                    $staleGuests++
                    $last = if ($lastSignIn) { $lastSignIn.ToString('yyyy-MM-dd') } else { 'nunca' }
                    Add-Finding -Category 'Stale guest' -Identity $upn -Detail "último login: $last (>90d)" -Severity 'WARN'
                }
            } else {
                if (-not $lastSignIn -or $lastSignIn -lt $cutoffMember) {
                    $staleMembers++
                    $last = if ($lastSignIn) { $lastSignIn.ToString('yyyy-MM-dd') } else { 'nunca' }
                    Add-Finding -Category 'Stale member' -Identity $upn -Detail "último login: $last (>180d)" -Severity 'WARN'
                }
            }
        }
    }

    Write-KeyValue -Key 'Usuarios revisados'            -Value $users.Count
    Write-KeyValue -Key 'Stale members (>180d)'         -Value $staleMembers -ValueColor Yellow
    Write-KeyValue -Key 'Stale guests (>90d)'           -Value $staleGuests -ValueColor Yellow
    Write-KeyValue -Key 'Deshabilitados con licencia'   -Value $disabledWithLicense -ValueColor Yellow

    # --- Groups: owner / membership health ---
    Show-Section -Title 'Análisis de grupos'

    $groups = @()
    try {
        $groups = @(Invoke-WithRetry -OperationName 'Get-MgGroup audit' -ScriptBlock {
            Get-MgGroup -All -Property Id,DisplayName,Mail,GroupTypes,MailEnabled,SecurityEnabled -ErrorAction Stop
        })
    } catch {
        Write-Log "Get-MgGroup fallo: $($_.Exception.Message)" -Level ERROR -Source 'IdentityAudit'
    }

    $noOwner = 0; $emptyM365 = 0; $emptyDls = 0
    $cap = [math]::Min(200, $groups.Count)
    $i = 0
    foreach ($g in $groups | Select-Object -First $cap) {
        $i++
        Write-Progress -Activity 'Auditando grupos' -Status "$i/$cap" -PercentComplete ([math]::Round(($i / $cap) * 100, 1))

        try {
            $owners = @(Invoke-WithRetry -OperationName 'Get-MgGroupOwner' -Quiet -ScriptBlock {
                Get-MgGroupOwner -GroupId $g.Id -All -ErrorAction Stop
            })
            if ($owners.Count -eq 0) {
                $noOwner++
                Add-Finding -Category 'Group no owner' -Identity ([string]$g.DisplayName) -Detail ('id=' + $g.Id) -Severity 'WARN'
            }
        } catch {}

        try {
            $members = @(Invoke-WithRetry -OperationName 'Get-MgGroupMember' -Quiet -ScriptBlock {
                Get-MgGroupMember -GroupId $g.Id -Top 1 -ErrorAction Stop
            })
            if ($members.Count -eq 0) {
                if ($g.GroupTypes -contains 'Unified') {
                    $emptyM365++
                    Add-Finding -Category 'M365 group vacío' -Identity ([string]$g.DisplayName) -Detail ('id=' + $g.Id) -Severity 'WARN'
                } elseif ($g.MailEnabled -and -not $g.SecurityEnabled) {
                    $emptyDls++
                    Add-Finding -Category 'DL vacía' -Identity ([string]$g.DisplayName) -Detail ('id=' + $g.Id) -Severity 'WARN'
                }
            }
        } catch {}
    }
    Write-Progress -Activity 'Auditando grupos' -Completed

    Write-KeyValue -Key 'Grupos revisados'   -Value $cap
    Write-KeyValue -Key 'Sin owner'          -Value $noOwner -ValueColor Yellow
    Write-KeyValue -Key 'M365 vacíos'        -Value $emptyM365 -ValueColor Yellow
    Write-KeyValue -Key 'DL vacías'          -Value $emptyDls -ValueColor Yellow

    $ended = Get-Date
    $summary = [ordered]@{
        'Usuarios'                       = $users.Count
        'Stale members'                  = $staleMembers
        'Stale guests'                   = $staleGuests
        'Disabled+License'               = $disabledWithLicense
        'Grupos revisados'               = $cap
        'Sin owner'                      = $noOwner
        'M365 vacíos'                    = $emptyM365
        'DL vacías'                      = $emptyDls
    }

    $logFile = Stop-LogSession -Persist
    $report = Publish-OperationReport `
        -Title 'Identity Audit' `
        -Operation 'Invoke-IdentityAudit' `
        -StartTime $started -EndTime $ended `
        -Summary $summary `
        -Items $findings `
        -WarnFields @('Stale members','Stale guests','Disabled+License','Sin owner','M365 vacíos','DL vacías') `
        -LogFile $logFile
    Write-KeyValue -Key 'Informe HTML' -Value $report

} catch {
    Show-ErrorBlock -Title 'Auditoría fallida' -Detail $_.Exception.Message
    Stop-LogSession -Persist | Out-Null
    throw
}

