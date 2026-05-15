#requires -Version 7.4
[CmdletBinding()]
param(
    [string]$TargetUpn      = 'testeo224@es.andersen.com',
    [string]$DelegateUpn    = 'testeo6@es.andersen.com',
    [string]$GroupNameSeed  = 'testeo-selftest',
    [switch]$SkipCleanup
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
try { $Host.UI.RawUI.WindowTitle = 'GREX365 · Self-test' } catch {}

# Autonomous validation against testeo* objects.
#
# What it does (CREATE-ONLY, never touches existing objects):
#   1. Resolve testeo* users via Graph (read-only).
#   2. Create a new DL whose name contains 'testeo-selftest-<timestamp>'.
#   3. Add the resolved testeo users as members via Add-DistributionGroupMember.
#   4. Verify membership.
#   5. Grant + revoke FullAccess between two testeo mailboxes (if available).
#   6. Apply a HiddenFromGAL=true and revert.
#   7. (optional) clean up: delete the temp DL.
#
# Memory rule honoured: "only new objects via scripts, never touch or delete existing".
# The temp DL is created fresh and removed at end (unless -SkipCleanup).

Assert-RequiredServicesReady

Show-Header -Title 'GREX365' -Subtitle 'Self-test sobre objetos testeo*'

$session = Start-LogSession -Name 'SelfTest'
$started = Get-Date
$report  = New-Object System.Collections.Generic.List[object]
$ok = 0; $fail = 0; $skip = 0

function Add-TestResult {
    param([string]$Step,[string]$Estado,[string]$Detalle = '')
    $report.Add([PSCustomObject]@{ Step=$Step; Estado=$Estado; Detalle=$Detalle; At=(Get-Date).ToString('HH:mm:ss') })
    switch ($Estado) {
        'OK'   { $script:ok++ }
        'FAIL' { $script:fail++ }
        'SKIP' { $script:skip++ }
    }
    $level = switch ($Estado) { 'OK'{'OK'} 'FAIL'{'ERROR'} default{'WARN'} }
    Write-Log "[$Estado] $Step$(if ($Detalle){' — '+$Detalle})" -Level $level -Source 'SelfTest'
}

function Try-Step {
    param([string]$Name,[scriptblock]$Block)
    try {
        $detail = & $Block
        Add-TestResult -Step $Name -Estado 'OK' -Detalle ([string]$detail)
        return $true
    } catch {
        Add-TestResult -Step $Name -Estado 'FAIL' -Detalle $_.Exception.Message
        return $false
    }
}

$tempDl = $null
$grantedFA = $false

try {
    # 1. Resolve testeo users
    $target = $null; $delegate = $null
    Try-Step 'Resolver usuario destino (testeo)' {
        $script:target = Invoke-WithRetry -OperationName 'Get-MgUser target' -ScriptBlock {
            Get-MgUser -UserId $TargetUpn -Property Id,DisplayName,UserPrincipalName -ErrorAction Stop
        }
        if (-not $target) { throw "No encontrado: $TargetUpn" }
        "$($target.DisplayName) <$($target.UserPrincipalName)>"
    } | Out-Null

    Try-Step 'Resolver usuario delegado (testeo)' {
        $script:delegate = Invoke-WithRetry -OperationName 'Get-MgUser delegate' -ScriptBlock {
            Get-MgUser -UserId $DelegateUpn -Property Id,DisplayName,UserPrincipalName -ErrorAction Stop
        }
        if (-not $delegate) { throw "No encontrado: $DelegateUpn" }
        "$($delegate.DisplayName) <$($delegate.UserPrincipalName)>"
    } | Out-Null

    if (-not $target -or -not $delegate) {
        Add-TestResult -Step 'Pre-requisitos' -Estado 'FAIL' -Detalle 'Usuarios testeo* no encontrados. Abortando.'
        return
    }

    # 2. Create a new DL exclusively for this test
    $stamp = Get-Date -Format 'yyyyMMddHHmmss'
    $dlAlias = ("$GroupNameSeed-$stamp").ToLowerInvariant()
    $dlName  = ("$GroupNameSeed-$stamp")
    Try-Step 'Crear DL temporal (testeo-selftest-<timestamp>)' {
        Invoke-WithRetry -OperationName 'New-DistributionGroup' -ScriptBlock {
            $script:tempDl = New-DistributionGroup -Name $dlName -Alias $dlAlias -Type Distribution -ErrorAction Stop
        }
        "DL creada: $($tempDl.PrimarySmtpAddress)"
    } | Out-Null

    if (-not $tempDl) {
        Add-TestResult -Step 'DL temporal' -Estado 'FAIL' -Detalle 'No se pudo crear DL. Abortando.'
        return
    }

    # 3. Add members
    Try-Step 'Añadir miembros (target + delegate)' {
        Invoke-WithRetry -OperationName 'Add-DistributionGroupMember target' -ScriptBlock {
            Add-DistributionGroupMember -Identity $dlAlias -Member $target.UserPrincipalName -ErrorAction Stop
        }
        Invoke-WithRetry -OperationName 'Add-DistributionGroupMember delegate' -ScriptBlock {
            Add-DistributionGroupMember -Identity $dlAlias -Member $delegate.UserPrincipalName -ErrorAction Stop
        }
        '2 miembros añadidos'
    } | Out-Null

    # 4. Verify
    Try-Step 'Verificar miembros (2 esperados)' {
        $members = @(Get-DistributionGroupMember -Identity $dlAlias -ResultSize Unlimited -ErrorAction Stop)
        if ($members.Count -lt 2) { throw "Esperados >=2, obtenidos $($members.Count)" }
        "$($members.Count) miembros confirmados"
    } | Out-Null

    # 5. HiddenFromGAL toggle
    Try-Step 'HiddenFromGAL=true' {
        Invoke-WithRetry -OperationName 'Set-DistributionGroup hide' -ScriptBlock {
            Set-DistributionGroup -Identity $dlAlias -HiddenFromAddressListsEnabled $true -ErrorAction Stop
        }
        'oculto'
    } | Out-Null
    Try-Step 'HiddenFromGAL=false (revert)' {
        Invoke-WithRetry -OperationName 'Set-DistributionGroup unhide' -ScriptBlock {
            Set-DistributionGroup -Identity $dlAlias -HiddenFromAddressListsEnabled $false -ErrorAction Stop
        }
        'visible'
    } | Out-Null

    # 6. FullAccess between testeo mailboxes (skip if either has no mailbox)
    $targetHasMbx = $false; $delegateHasMbx = $false
    try { Get-Mailbox -Identity $TargetUpn -ErrorAction Stop | Out-Null; $targetHasMbx = $true } catch {}
    try { Get-Mailbox -Identity $DelegateUpn -ErrorAction Stop | Out-Null; $delegateHasMbx = $true } catch {}
    if ($targetHasMbx -and $delegateHasMbx) {
        Try-Step 'Add-MailboxPermission FullAccess (delegate -> target)' {
            Invoke-WithRetry -OperationName 'Add-MailboxPermission FA' -ScriptBlock {
                Add-MailboxPermission -Identity $TargetUpn -User $DelegateUpn -AccessRights FullAccess -InheritanceType All -AutoMapping:$false -Confirm:$false -ErrorAction Stop | Out-Null
            }
            $script:grantedFA = $true
            'FullAccess otorgado'
        } | Out-Null
        Try-Step 'Verificar permiso FullAccess' {
            $perms = @(Get-MailboxPermission -Identity $TargetUpn -User $DelegateUpn -ErrorAction Stop)
            $hit = $perms | Where-Object { $_.AccessRights -contains 'FullAccess' } | Select-Object -First 1
            if (-not $hit) { throw 'FullAccess no encontrado en consulta posterior' }
            'verificado'
        } | Out-Null
        Try-Step 'Remove-MailboxPermission FullAccess (cleanup)' {
            Invoke-WithRetry -OperationName 'Remove-MailboxPermission FA' -ScriptBlock {
                Remove-MailboxPermission -Identity $TargetUpn -User $DelegateUpn -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop | Out-Null
            }
            $script:grantedFA = $false
            'revocado'
        } | Out-Null
    } else {
        Add-TestResult -Step 'FullAccess test' -Estado 'SKIP' -Detalle "Sin buzón en target=$targetHasMbx delegate=$delegateHasMbx"
    }

    # 7. Cleanup DL
    if (-not $SkipCleanup) {
        Try-Step 'Eliminar DL temporal (cleanup)' {
            Invoke-WithRetry -OperationName 'Remove-DistributionGroup' -ScriptBlock {
                Remove-DistributionGroup -Identity $dlAlias -Confirm:$false -ErrorAction Stop
            }
            'DL eliminada'
        } | Out-Null
    } else {
        Add-TestResult -Step 'Cleanup' -Estado 'SKIP' -Detalle "Solicitado -SkipCleanup. DL persiste: $dlAlias"
    }

} catch {
    Add-TestResult -Step 'Self-test runner' -Estado 'FAIL' -Detalle $_.Exception.Message
} finally {
    # Defensive: revoke FA if granted and not yet revoked
    if ($grantedFA) {
        try {
            Remove-MailboxPermission -Identity $TargetUpn -User $DelegateUpn -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        } catch {}
    }
}

$ended = Get-Date

Show-Section -Title 'Resumen self-test'
Write-KeyValue -Key 'Pasos OK'   -Value $ok   -ValueColor Green
Write-KeyValue -Key 'Pasos FAIL' -Value $fail -ValueColor Red
Write-KeyValue -Key 'Pasos SKIP' -Value $skip -ValueColor Yellow
Write-KeyValue -Key 'Duración'   -Value (('{0:N1}s' -f ($ended - $started).TotalSeconds))

$logFile = Stop-LogSession -Persist
$rep = Publish-OperationReport `
    -Title 'Self-test · testeo*' `
    -Operation 'Invoke-SelfTest' `
    -StartTime $started -EndTime $ended `
    -Summary ([ordered]@{ 'Pasos OK'=$ok; 'Pasos FAIL'=$fail; 'Pasos SKIP'=$skip; 'Target'=$TargetUpn; 'Delegate'=$DelegateUpn }) `
    -Items $report `
    -OkFields @('Pasos OK') -ErrorFields @('Pasos FAIL') -WarnFields @('Pasos SKIP') `
    -LogFile $logFile
Write-KeyValue -Key 'Informe HTML' -Value $rep

