#requires -Version 7.4
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
try { $Host.UI.RawUI.WindowTitle = 'GREX365 · Offboarding wizard' } catch {}

# Orchestrator for employee offboarding. Performs up to 14 steps:
#   1. Resolve user
#   2. Block sign-in
#   3. Revoke sessions
#   4. Remove licenses
#   5. Convert mailbox to Shared
#   6. Configure auto-reply (template-based)
#   7. Set forwarding to delegate
#   8. Grant FullAccess to delegate
#   9. Grant SendAs to delegate
#  10. Hide from GAL
#  11. Remove from distribution lists / M365 groups
#  12. Disable MFA methods (best effort)
#  13. Handover note (rendered template, logged)
#  14. HTML report generation
#
# Safety: dry-run forced in support mode. Real execution requires admin role,
# typed UPN confirmation, and (by default) UPN starting with "testeo" or an
# explicit -AllowProductionTarget flag in metadata.

Assert-RequiredServicesReady
Assert-Role -Required admin -Operation 'offboarding'

Show-Header -Title 'GREX365' -Subtitle 'Offboarding wizard'

$session = Start-LogSession -Name 'Offboarding'
$started = Get-Date

# --- step helpers ---

$script:OffSteps = New-Object System.Collections.Generic.List[object]

function Add-OffboardingStep {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][ValidateSet('OK','FAIL','SKIP','DRY')][string]$Result,
        [string]$Detail = ''
    )
    $script:OffSteps.Add([PSCustomObject]@{
        Step    = $Name
        Result  = $Result
        Detail  = $Detail
        At      = (Get-Date).ToString('HH:mm:ss')
    })
    $level = switch ($Result) {
        'OK'   { 'OK' }
        'FAIL' { 'ERROR' }
        'SKIP' { 'WARN' }
        'DRY'  { 'INFO' }
    }
    Write-Log ("[{0}] {1}{2}" -f $Result, $Name, $(if ($Detail) { ' — ' + $Detail } else { '' })) -Level $level -Source 'Offboarding'
    try { Write-AuditEvent -EventType 'OffboardingStep' -Properties @{ step=$Name; result=$Result; detail=$Detail } } catch {}
}

function Invoke-Step {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][scriptblock]$Action,
        [switch]$DryRun
    )
    if ($DryRun) {
        Add-OffboardingStep -Name $Name -Result 'DRY' -Detail 'no se ejecutó (dry-run)'
        return
    }
    try {
        $detail = & $Action
        Add-OffboardingStep -Name $Name -Result 'OK' -Detail ([string]$detail)
    } catch {
        Add-OffboardingStep -Name $Name -Result 'FAIL' -Detail $_.Exception.Message
    }
}

try {
    # --- inputs ---

    $rawUpn = Read-Input -Prompt 'Correo del usuario saliente'
    $upn = Normalize-Input -Value $rawUpn
    if (-not (Test-Email -Value $upn)) { throw 'INPUT_INVALID' }

    # Delegates: accept comma/semicolon-separated list. First one is the "primary"
    # (gets the mailbox forward); all of them get FullAccess + SendAs + appear in
    # the auto-reply / handover template.
    $rawDelegate = Read-Input -Prompt 'Correo(s) de delegado(s) — separados por coma'
    $delegateInput = Normalize-Input -Value $rawDelegate
    if ([string]::IsNullOrWhiteSpace($delegateInput)) { throw 'INPUT_INVALID' }
    $delegateEmails = @($delegateInput -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    foreach ($d in $delegateEmails) {
        if (-not (Test-Email -Value $d)) { throw "Email de delegado inválido: $d" }
    }

    $rawManager = Read-Input -Prompt 'Correo del manager (handover note, opcional)'
    $manager = Normalize-Input -Value $rawManager
    if ($manager -and -not (Test-Email -Value $manager)) { throw 'INPUT_INVALID' }

    # Org name: derive from connected tenant (no longer prompts).
    $orgName = 'la organización'
    try {
        $state = Get-SessionState
        if ($state.TenantDomain) {
            $orgName = ([string]$state.TenantDomain -replace '^([^.]+).*$','$1')
            if ($orgName) { $orgName = $orgName.Substring(0,1).ToUpperInvariant() + $orgName.Substring(1) }
        }
    } catch {}

    # --- resolve identities ---

    Show-Section -Title 'Validando identidades'
    $mgUser = $null
    try {
        $mgUser = Invoke-WithRetry -OperationName 'Get-MgUser' -ScriptBlock {
            Get-MgUser -UserId $upn -Property Id,DisplayName,UserPrincipalName,AccountEnabled,AssignedLicenses -ErrorAction Stop
        }
    } catch { throw "Usuario no encontrado en Graph: $upn" }
    Write-KeyValue -Key 'Usuario' -Value ($mgUser.DisplayName + ' <' + $mgUser.UserPrincipalName + '>')

    # Resolve all delegates.
    $delegateUsers = New-Object System.Collections.Generic.List[object]
    foreach ($d in $delegateEmails) {
        $du = $null
        try {
            $du = Invoke-WithRetry -OperationName ('Get-MgUser(delegate) ' + $d) -ScriptBlock {
                Get-MgUser -UserId $d -Property Id,DisplayName,UserPrincipalName -ErrorAction Stop
            }
        } catch { throw "Delegado no encontrado en Graph: $d" }
        if ($du) { $delegateUsers.Add($du) }
    }
    $delegateUser = $delegateUsers[0]   # primary
    Write-KeyValue -Key 'Delegado(s)' -Value (($delegateUsers | ForEach-Object { $_.DisplayName + ' <' + $_.UserPrincipalName + '>' }) -join ' · ')
    if ($delegateUsers.Count -gt 1) {
        Write-Log ("Detectados " + $delegateUsers.Count + " delegados. Primary (forward): " + $delegateUser.UserPrincipalName + ". Resto recibe FullAccess + SendAs.") -Level INFO -Source 'Offboarding'
    }

    $mbox = $null
    try {
        $mbox = Invoke-WithRetry -OperationName 'Get-Mailbox' -ScriptBlock { Get-Mailbox -Identity $upn -ErrorAction Stop }
        Write-KeyValue -Key 'Buzón' -Value ($mbox.PrimarySmtpAddress.ToString() + ' (' + $mbox.RecipientTypeDetails + ')')
    } catch {
        Write-Log "Usuario sin buzón en EXO — pasos de mailbox se omitirán." -Level WARN -Source 'Offboarding'
    }

    # --- template selection ---

    Show-Section -Title 'Plantilla de auto-reply'
    $template = Select-TemplateInteractive -Category 'auto-reply'

    # --- safety gates ---

    $supportMode = Test-IsSupportMode
    $isTestUser  = $upn.ToLowerInvariant().StartsWith('testeo')
    $forcedDry   = $supportMode -or -not $isTestUser

    Show-Section -Title 'Resumen de ejecución'
    Write-KeyValue -Key 'Usuario'      -Value $upn
    Write-KeyValue -Key 'Delegado(s)'  -Value (($delegateUsers | ForEach-Object { $_.UserPrincipalName }) -join ', ')
    Write-KeyValue -Key 'Manager'      -Value ($manager ? $manager : '—')
    Write-KeyValue -Key 'Plantilla'  -Value ($template ? $template.Name : '— (sin auto-reply)')
    Write-KeyValue -Key 'Modo UI'    -Value (Get-CurrentUIMode)
    Write-KeyValue -Key 'Rol'        -Value (Get-CurrentRole)
    Write-KeyValue -Key 'Dry-run forzado' -Value $forcedDry

    if ($forcedDry) {
        Show-WarningBlock -Title 'Modo DRY-RUN activado' -Detail @"
Razones posibles:
  · UIMode = support (cambia a 'advanced' en Preferencias para permitir ejecución real)
  · UPN no empieza por 'testeo' (regla de seguridad para evitar offboardings accidentales en producción)

Se mostrarán los pasos sin tocar el tenant.
"@
    } else {
        if (-not (Confirm-DestructiveAction -Operation 'OFFBOARDING COMPLETO' -Target $upn -RequireType)) {
            Add-OffboardingStep -Name 'Confirmación' -Result 'SKIP' -Detail 'Cancelado por el usuario'
            throw 'INPUT_EMPTY'
        }
        $typed = Read-Input -Prompt "Escribe el UPN exacto del usuario a offboard"
        if ((Normalize-Input -Value $typed) -ne $upn) {
            throw 'UPN tecleado no coincide. Cancelado.'
        }
    }

    Show-Section -Title 'Ejecutando pasos'

    $dr = $forcedDry

    # 1. Block sign-in
    Invoke-Step -Name 'Bloquear inicio de sesión' -DryRun:$dr -Action {
        Invoke-WithRetry -OperationName 'Update-MgUser disable' -ScriptBlock {
            Update-MgUser -UserId $mgUser.Id -AccountEnabled:$false -ErrorAction Stop
        }
        "AccountEnabled=false aplicado"
    }

    # 2. Revoke sessions
    Invoke-Step -Name 'Revocar sesiones activas' -DryRun:$dr -Action {
        Invoke-WithRetry -OperationName 'Revoke-MgUserSignInSession' -ScriptBlock {
            Revoke-MgUserSignInSession -UserId $mgUser.Id -ErrorAction Stop | Out-Null
        }
        "Sesiones revocadas"
    }

    # 3. Remove licenses
    Invoke-Step -Name 'Eliminar licencias' -DryRun:$dr -Action {
        $skus = @($mgUser.AssignedLicenses | ForEach-Object { $_.SkuId })
        if (-not $skus -or $skus.Count -eq 0) { return 'sin licencias asignadas' }
        Invoke-WithRetry -OperationName 'Set-MgUserLicense' -ScriptBlock {
            Set-MgUserLicense -UserId $mgUser.Id -AddLicenses @() -RemoveLicenses $skus -ErrorAction Stop | Out-Null
        }
        ("{0} licencias removidas: {1}" -f $skus.Count, ($skus -join ', '))
    }

    if ($mbox) {

        # 4. Convert to shared
        Invoke-Step -Name 'Convertir buzón a Shared' -DryRun:$dr -Action {
            Invoke-WithRetry -OperationName 'Set-Mailbox Shared' -ScriptBlock {
                Set-Mailbox -Identity $upn -Type Shared -ErrorAction Stop
            }
            'Mailbox -> Shared'
        }

        # 5. Auto-reply (custom body overrides template if provided via env GREX365_OFF_CUSTOMBODY)
        $customBody = [string]$env:GREX365_OFF_CUSTOMBODY
        if ($template -or $customBody) {
            Invoke-Step -Name 'Configurar auto-reply' -DryRun:$dr -Action {
                $delegatesNames = (($delegateUsers | ForEach-Object { $_.DisplayName + ' (' + $_.UserPrincipalName + ')' }) -join ', ')
                $delegatesEmails = (($delegateUsers | ForEach-Object { $_.UserPrincipalName }) -join ', ')
                $vals = @{
                    user              = $mgUser.DisplayName
                    org               = $orgName
                    date              = (Get-Date -Format 'yyyy-MM-dd')
                    replacement       = $delegateUser.DisplayName
                    replacement_email = $delegateUser.UserPrincipalName
                    delegates         = $delegatesNames
                    delegates_emails  = $delegatesEmails
                    manager           = $manager
                }

                if ($customBody) {
                    # Custom body wins. Expand {placeholders}.
                    $body = $customBody
                    foreach ($k in $vals.Keys) { $body = $body -replace ('\{' + [regex]::Escape($k) + '\}'), [string]$vals[$k] }
                    Invoke-WithRetry -OperationName 'Set-MailboxAutoReplyConfiguration(custom)' -ScriptBlock {
                        Set-MailboxAutoReplyConfiguration -Identity $upn `
                            -AutoReplyState Enabled `
                            -ExternalAudience All `
                            -InternalMessage $body `
                            -ExternalMessage $body `
                            -ErrorAction Stop
                    }
                    return "Auto-reply Enabled (mensaje personalizado)"
                }

                $rendered = Render-Template -Name $template.Name -Values $vals
                $external = if ($rendered.BodyHtml) { $rendered.BodyHtml } else { $rendered.BodyText }
                $internal = $rendered.BodyText
                Invoke-WithRetry -OperationName 'Set-MailboxAutoReplyConfiguration' -ScriptBlock {
                    Set-MailboxAutoReplyConfiguration -Identity $upn `
                        -AutoReplyState Enabled `
                        -ExternalAudience All `
                        -InternalMessage $internal `
                        -ExternalMessage $external `
                        -ErrorAction Stop
                }
                "Auto-reply Enabled (plantilla=$($template.Name))"
            }
        } else {
            Add-OffboardingStep -Name 'Configurar auto-reply' -Result 'SKIP' -Detail 'sin plantilla ni mensaje personalizado'
        }

        # 6. Forward — first delegate only (Exchange only accepts a single SMTP)
        Invoke-Step -Name 'Forward al delegado primario' -DryRun:$dr -Action {
            Invoke-WithRetry -OperationName 'Set-Mailbox forwarding' -ScriptBlock {
                Set-Mailbox -Identity $upn `
                    -ForwardingSmtpAddress $delegateUser.UserPrincipalName `
                    -DeliverToMailboxAndForward $true `
                    -ErrorAction Stop
            }
            "Forward -> $($delegateUser.UserPrincipalName) (entrega doble)"
        }

        # 7. FullAccess for every delegate
        if ($dr) {
            Add-OffboardingStep -Name 'FullAccess delegado(s)' -Result 'DRY' -Detail ('no se ejecutó · destinos: ' + (($delegateUsers | ForEach-Object { $_.UserPrincipalName }) -join ', '))
        } else {
            $okFA = 0; $failFA = @()
            foreach ($du in $delegateUsers) {
                try {
                    Invoke-WithRetry -OperationName ('Add-MailboxPermission ' + $du.UserPrincipalName) -ScriptBlock {
                        Add-MailboxPermission -Identity $upn -User $du.UserPrincipalName `
                            -AccessRights FullAccess -InheritanceType All -AutoMapping:$false `
                            -Confirm:$false -ErrorAction Stop | Out-Null
                    }
                    $okFA++
                } catch { $failFA += ($du.UserPrincipalName + ' (' + $_.Exception.Message + ')') }
            }
            if ($failFA.Count -eq 0) {
                Add-OffboardingStep -Name 'FullAccess delegado(s)' -Result 'OK' -Detail "FullAccess concedido a $okFA delegado(s)"
            } elseif ($okFA -gt 0) {
                Add-OffboardingStep -Name 'FullAccess delegado(s)' -Result 'OK' -Detail "Parcial · ok=$okFA fail=$($failFA -join '; ')"
            } else {
                Add-OffboardingStep -Name 'FullAccess delegado(s)' -Result 'FAIL' -Detail ($failFA -join '; ')
            }
        }

        # 8. SendAs for every delegate
        if ($dr) {
            Add-OffboardingStep -Name 'SendAs delegado(s)' -Result 'DRY' -Detail ('no se ejecutó · destinos: ' + (($delegateUsers | ForEach-Object { $_.UserPrincipalName }) -join ', '))
        } else {
            $okSA = 0; $failSA = @()
            foreach ($du in $delegateUsers) {
                try {
                    Invoke-WithRetry -OperationName ('Add-RecipientPermission ' + $du.UserPrincipalName) -ScriptBlock {
                        Add-RecipientPermission -Identity $upn -Trustee $du.UserPrincipalName `
                            -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                    }
                    $okSA++
                } catch { $failSA += ($du.UserPrincipalName + ' (' + $_.Exception.Message + ')') }
            }
            if ($failSA.Count -eq 0) {
                Add-OffboardingStep -Name 'SendAs delegado(s)' -Result 'OK' -Detail "SendAs concedido a $okSA delegado(s)"
            } elseif ($okSA -gt 0) {
                Add-OffboardingStep -Name 'SendAs delegado(s)' -Result 'OK' -Detail "Parcial · ok=$okSA fail=$($failSA -join '; ')"
            } else {
                Add-OffboardingStep -Name 'SendAs delegado(s)' -Result 'FAIL' -Detail ($failSA -join '; ')
            }
        }

        # 9. Hide from GAL (handle hybrid: object synced from on-prem AD can't be modified in EXO)
        if ($dr) {
            Add-OffboardingStep -Name 'Ocultar de la GAL' -Result 'DRY' -Detail 'no se ejecutó (dry-run)'
        } else {
            try {
                Invoke-WithRetry -OperationName 'Set-Mailbox HideFromGAL' -ScriptBlock {
                    Set-Mailbox -Identity $upn -HiddenFromAddressListsEnabled $true -ErrorAction Stop
                }
                Add-OffboardingStep -Name 'Ocultar de la GAL' -Result 'OK' -Detail 'HiddenFromAddressListsEnabled=true'
            } catch {
                $msg = $_.Exception.Message
                if ($msg -match 'sincroniz|on-prem|on premises|write scope|ámbito de escritura|organización local|cannot be performed.*synchron') {
                    Add-OffboardingStep -Name 'Ocultar de la GAL' -Result 'SKIP' -Detail 'Objeto híbrido sincronizado desde AD on-prem · debe aplicarse en AD local (msExchHideFromAddressLists=TRUE) y esperar a que Entra Connect propague.'
                } else {
                    Add-OffboardingStep -Name 'Ocultar de la GAL' -Result 'FAIL' -Detail $msg
                }
            }
        }
    } else {
        foreach ($n in 'Convertir buzón a Shared','Configurar auto-reply','Forward al delegado primario','FullAccess delegado(s)','SendAs delegado(s)','Ocultar de la GAL') {
            Add-OffboardingStep -Name $n -Result 'SKIP' -Detail 'no hay buzón'
        }
    }

    # 10. Remove from DLs / M365 groups (best effort via Graph member-of)
    Invoke-Step -Name 'Quitar de DLs / M365 Groups' -DryRun:$dr -Action {
        $memberOf = @()
        try {
            $memberOf = @(Invoke-WithRetry -OperationName 'Get-MgUserMemberOf' -ScriptBlock {
                Get-MgUserMemberOf -UserId $mgUser.Id -All -ErrorAction Stop
            })
        } catch {}
        $removed = 0
        foreach ($g in $memberOf) {
            $gid = [string]$g.Id
            if (-not $gid) { continue }
            try {
                Invoke-WithRetry -OperationName 'Remove-MgGroupMemberByRef' -Quiet -ScriptBlock {
                    Remove-MgGroupMemberByRef -GroupId $gid -DirectoryObjectId $mgUser.Id -ErrorAction Stop
                }
                $removed++
            } catch {}
        }
        "Eliminado de $removed grupos (de $($memberOf.Count) candidatos)"
    }

    # 11. Disable MFA methods (auth methods removal)
    Invoke-Step -Name 'Desactivar métodos MFA' -DryRun:$dr -Action {
        $removed = 0
        try {
            $methods = @(Invoke-WithRetry -OperationName 'Get-MgUserAuthenticationMethod' -ScriptBlock {
                Get-MgUserAuthenticationMethod -UserId $mgUser.Id -All -ErrorAction Stop
            })
            foreach ($m in $methods) {
                $type = [string]$m.AdditionalProperties.'@odata.type'
                $id   = [string]$m.Id
                if (-not $id) { continue }
                if ($type -match 'passwordAuthenticationMethod') { continue } # cannot remove primary password
                try {
                    switch -Wildcard ($type) {
                        '*phoneAuthenticationMethod*'           { Remove-MgUserAuthenticationPhoneMethod        -UserId $mgUser.Id -PhoneAuthenticationMethodId $id -ErrorAction Stop; $removed++ }
                        '*microsoftAuthenticatorAuthenticationMethod*' { Remove-MgUserAuthenticationMicrosoftAuthenticatorMethod -UserId $mgUser.Id -MicrosoftAuthenticatorAuthenticationMethodId $id -ErrorAction Stop; $removed++ }
                        '*fido2AuthenticationMethod*'           { Remove-MgUserAuthenticationFido2Method        -UserId $mgUser.Id -Fido2AuthenticationMethodId $id -ErrorAction Stop; $removed++ }
                        '*softwareOathAuthenticationMethod*'    { Remove-MgUserAuthenticationSoftwareOathMethod -UserId $mgUser.Id -SoftwareOathAuthenticationMethodId $id -ErrorAction Stop; $removed++ }
                        '*windowsHelloForBusinessAuthenticationMethod*' { Remove-MgUserAuthenticationWindowsHelloForBusinessMethod -UserId $mgUser.Id -WindowsHelloForBusinessAuthenticationMethodId $id -ErrorAction Stop; $removed++ }
                    }
                } catch {}
            }
        } catch {}
        "Métodos eliminados: $removed"
    }

    # 12. Handover note
    if ($manager) {
        Invoke-Step -Name 'Nota de traspaso (handover)' -DryRun:$false -Action {
            $delegatesNames  = (($delegateUsers | ForEach-Object { $_.DisplayName + ' (' + $_.UserPrincipalName + ')' }) -join ', ')
            $delegatesEmails = (($delegateUsers | ForEach-Object { $_.UserPrincipalName }) -join ', ')
            $vals = @{
                user              = $mgUser.DisplayName
                org               = $orgName
                date              = (Get-Date -Format 'yyyy-MM-dd')
                manager           = $manager
                replacement       = $delegateUser.DisplayName
                replacement_email = $delegateUser.UserPrincipalName
                delegates         = $delegatesNames
                delegates_emails  = $delegatesEmails
            }
            $hand = $null
            try { $hand = Render-Template -Name 'handover-internal-es' -Values $vals } catch {}
            if ($hand) {
                Write-Log ("Handover SUBJECT: " + $hand.Subject) -Source 'Offboarding'
                foreach ($line in $hand.BodyText -split "`n") { Write-Log ('  ' + $line) -Source 'Offboarding' }
                'plantilla renderizada y logueada'
            } else {
                'plantilla handover-internal-es no encontrada'
            }
        }
    } else {
        Add-OffboardingStep -Name 'Nota de traspaso (handover)' -Result 'SKIP' -Detail 'sin manager'
    }

    # --- summary ---

    $ended = Get-Date
    $okCount   = @($script:OffSteps | Where-Object Result -eq 'OK').Count
    $failCount = @($script:OffSteps | Where-Object Result -eq 'FAIL').Count
    $dryCount  = @($script:OffSteps | Where-Object Result -eq 'DRY').Count
    $skipCount = @($script:OffSteps | Where-Object Result -eq 'SKIP').Count

    Show-Section -Title 'Resumen'
    Write-KeyValue -Key 'Pasos OK'     -Value $okCount  -ValueColor Green
    Write-KeyValue -Key 'Pasos DRY'    -Value $dryCount -ValueColor Cyan
    Write-KeyValue -Key 'Pasos SKIP'   -Value $skipCount -ValueColor Yellow
    Write-KeyValue -Key 'Pasos FAIL'   -Value $failCount -ValueColor Red
    Write-KeyValue -Key 'Duración'     -Value (('{0:N1}s' -f ($ended - $started).TotalSeconds))

    # 13. HTML report
    $logFile = Stop-LogSession -Persist
    $reportPath = Publish-OperationReport `
        -Title  ("Offboarding · $upn") `
        -Operation ("Offboarding · {0}" -f $mgUser.DisplayName) `
        -StartTime $started `
        -EndTime   $ended `
        -Summary ([ordered]@{
            Usuario       = $mgUser.UserPrincipalName
            'Delegado(s)' = (($delegateUsers | ForEach-Object { $_.UserPrincipalName }) -join ', ')
            Manager       = ($manager ? $manager : '—')
            'Pasos OK'    = $okCount
            'Pasos DRY'   = $dryCount
            'Pasos SKIP'  = $skipCount
            'Pasos FAIL'  = $failCount
        }) `
        -Items $script:OffSteps `
        -OkFields    @('Pasos OK') `
        -ErrorFields @('Pasos FAIL') `
        -WarnFields  @('Pasos DRY','Pasos SKIP') `
        -LogFile $logFile

    Write-KeyValue -Key 'Informe HTML' -Value $reportPath
    Write-KeyValue -Key 'Log'          -Value $logFile

} catch {
    if ($_.Exception.Message -in @('INPUT_EMPTY','INPUT_INVALID','INPUT_NOTFOUND')) {
        Show-WarningBlock -Title 'Entrada inválida o cancelada'
    } else {
        Show-ErrorBlock -Title 'Offboarding fallido' -Detail $_.Exception.Message
    }
    Stop-LogSession -Persist | Out-Null
    throw
}


