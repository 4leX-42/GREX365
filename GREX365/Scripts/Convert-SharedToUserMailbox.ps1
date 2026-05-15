#requires -Version 7.4
[CmdletBinding(SupportsShouldProcess)]
param([string]$Identity)

try { $Host.UI.RawUI.WindowTitle = 'GREX365 · SharedMailbox → UserMailbox' } catch {}

function Convert-SharedToRegular {
    param([Parameter(Mandatory = $true)][string]$UserIdentity)

    Write-Log ("Buscando buzón '{0}'..." -f $UserIdentity) -Source 'ConvertMbx'
    try {
        $mbx = Get-Mailbox -Identity $UserIdentity -ErrorAction Stop
    } catch {
        Show-ErrorBlock -Title 'Buzón no encontrado' -Detail $_.Exception.Message
        return $false
    }

    Show-Section -Title 'Buzón detectado'
    Write-KeyValue -Key 'Nombre'   -Value $mbx.DisplayName
    Write-KeyValue -Key 'UPN'      -Value $mbx.UserPrincipalName
    Write-KeyValue -Key 'Tipo'     -Value $mbx.RecipientTypeDetails

    if ($mbx.RecipientTypeDetails -ne 'SharedMailbox') {
        Show-WarningBlock -Title 'Sin acción necesaria' -Detail ("El buzón ya es '{0}'." -f $mbx.RecipientTypeDetails)
        return $true
    }

    Show-WarningBlock -Title 'SharedMailbox detectado' -Detail 'Este tipo impide que el usuario aparezca en Microsoft Teams.'
    $confirm = Read-Input -Prompt '¿Convertir a UserMailbox (Regular)? (S/N)'
    if ($confirm -notmatch '^[Ss]') {
        Write-Log 'Cancelado por el usuario.' -Source 'ConvertMbx'
        return $false
    }

    if (-not $PSCmdlet.ShouldProcess($UserIdentity, 'Set-Mailbox -Type Regular')) {
        Write-Log 'WhatIf activo. Sin cambios.' -Source 'ConvertMbx'
        return $false
    }

    Write-Log 'Aplicando Set-Mailbox -Type Regular...' -Source 'ConvertMbx'
    try {
        Set-Mailbox -Identity $UserIdentity -Type Regular -ErrorAction Stop
        Write-Log 'Set-Mailbox ejecutado.' -Level OK -Source 'ConvertMbx'
    } catch {
        Show-ErrorBlock -Title 'Error en Set-Mailbox' -Detail $_.Exception.Message
        return $false
    }

    $timeout = 60; $pollEvery = 5
    $elapsed = 0; $lastPoll = 0
    $current = $null
    $confirmed = $false

    Write-Host ''
    Write-Indent
    Write-Host 'Propagación Exchange Online' -ForegroundColor White
    Write-Indent
    Write-Host 'Set-Mailbox devolvió OK. El backend de EXO tarda 30–90s en reflejar el nuevo tipo.' -ForegroundColor DarkGray
    Write-Indent
    Write-Host ('Polling Get-Mailbox cada {0}s hasta confirmar UserMailbox o timeout en {1}s.' -f $pollEvery, $timeout) -ForegroundColor DarkGray
    Write-Host ''

    $barLen = 24
    while ($elapsed -lt $timeout -and -not $confirmed) {
        $filled = [int][Math]::Floor(($elapsed / [double]$timeout) * $barLen)
        if ($filled -lt 0) { $filled = 0 } elseif ($filled -gt $barLen) { $filled = $barLen }
        $bar = ('━' * $filled) + ('─' * ($barLen - $filled))
        $line = '  [{0,2}s / {1}s]  {2}  esperando sync EXO' -f $elapsed, $timeout, $bar
        try { [Console]::Write("`r" + $line.PadRight(78)) } catch {}

        Start-Sleep -Seconds 1
        $elapsed++

        if (($elapsed - $lastPoll) -ge $pollEvery) {
            $lastPoll = $elapsed
            try { $current = Get-Mailbox -Identity $UserIdentity -ErrorAction Stop } catch { $current = $null }
            if ($current -and $current.RecipientTypeDetails -eq 'UserMailbox') { $confirmed = $true }
        }
    }
    try { [Console]::Write("`r" + (' ' * 78) + "`r") } catch {}

    if ($current -and $current.RecipientTypeDetails -eq 'UserMailbox') {
        Write-Log 'Conversión confirmada.' -Level OK -Source 'ConvertMbx'
        Show-WarningBlock -Title 'Sincronización con Teams' -Detail '15–60 minutos. Verifica que la licencia de Teams está asignada.'

        $mbxId = ''
        if ($current.ExternalDirectoryObjectId) { $mbxId = [string]$current.ExternalDirectoryObjectId }
        elseif ($current.ExchangeObjectId)      { $mbxId = [string]$current.ExchangeObjectId }
        elseif ($current.Guid)                  { $mbxId = [string]$current.Guid }
        if ($mbxId) {
            Write-Host ''
            Show-AdminLink -Type 'UserMailbox' -Id $mbxId -DisplayName $current.DisplayName
        }
        return $true
    }

    Show-ErrorBlock -Title ('No se confirmó el cambio tras ' + $timeout + 's') -Detail 'Revisa manualmente con Get-Mailbox.'
    return $false
}

# --- ENTRYPOINT ---

Assert-RequiredServicesReady
Show-Header -Title 'GREX365' -Subtitle 'Convertir SharedMailbox → UserMailbox'

$session = Start-LogSession -Name 'Convert-Mailbox'

try {
    if ($Identity) {
        [void](Convert-SharedToRegular -UserIdentity $Identity)
    } else {
        do {
            $rawTarget = Read-Input -Prompt 'Email/UPN del usuario'
            $target = Normalize-Input -Value $rawTarget
            if (-not $target) { break }
            [void](Convert-SharedToRegular -UserIdentity $target)
            $more = Read-Input -Prompt '¿Comprobar otro usuario? (S/N)'
            if ($more -notmatch '^[Ss]') { break }
        } while ($true)
    }
} catch {
    Show-ErrorBlock -Title 'Operación fallida' -Detail $_.Exception.Message
}
finally {
    Stop-LogSession -Persist | Out-Null
}

