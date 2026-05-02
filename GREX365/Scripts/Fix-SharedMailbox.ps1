#requires -Version 5.1
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$Identity
)

$Host.UI.RawUI.WindowTitle = 'Fix SharedMailbox → UserMailbox | GREX365'

# --- LOGGING ---

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')][string]$Level = 'INFO'
    )
    $time = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $color = switch ($Level) { 'OK' { 'Green' } 'WARN' { 'Yellow' } 'ERROR' { 'Red' } default { 'Cyan' } }
    Write-Host "[$time][$Level] $Message" -ForegroundColor $color
}

# --- CABECERA ---

function Get-CenterPadding {
    param([Parameter(Mandatory = $true)][string]$Text)
    $w = [Console]::WindowWidth
    $pad = [Math]::Floor(($w - $Text.Length) / 2)
    if ($pad -lt 0) { $pad = 0 }
    return (' ' * $pad)
}

function Write-Centered {
    param([Parameter(Mandatory = $true)][string]$Text, [ConsoleColor]$Color = [ConsoleColor]::Gray)
    Write-Host ((Get-CenterPadding -Text $Text) + $Text) -ForegroundColor $Color
}

function Show-Header {
    Clear-Host
    Write-Host ''
    Write-Centered '◀◀◀ ════════════  Grex365  ════════════ ▶▶▶' -Color DarkCyan
    Write-Host ''
    Write-Centered 'CORREGIR SHAREDMAILBOX → USERMAILBOX' -Color Green
    Write-Centered '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━' -Color Cyan
    Write-Centered 'Hace al usuario visible en Microsoft Teams' -Color DarkCyan
    Write-Host ''
}

# --- VALIDACIÓN ---

function Assert-ExchangeReady {
    $ok = $false
    try {
        if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
            $s = Get-ConnectionInformation -ErrorAction SilentlyContinue
            if ($s | Where-Object { $_.State -eq 'Connected' }) { $ok = $true }
        }
    } catch {}
    if (-not $ok) {
        throw "No hay sesión activa de Exchange Online. Ejecuta este script desde Main.ps1."
    }
}

# --- LÓGICA PRINCIPAL ---

function Convert-SharedToRegular {
    param([Parameter(Mandatory = $true)][string]$UserIdentity)

    Write-Log "Buscando buzón para '$UserIdentity'..."
    try {
        $mailbox = Get-Mailbox -Identity $UserIdentity -ErrorAction Stop
    }
    catch {
        Write-Log "No se encontró el buzón '$UserIdentity': $($_.Exception.Message)" 'ERROR'
        return $false
    }

    Write-Host ''
    Write-Host "  Nombre      : $($mailbox.DisplayName)"          -ForegroundColor White
    Write-Host "  UPN         : $($mailbox.UserPrincipalName)"    -ForegroundColor White
    Write-Host "  Tipo actual : $($mailbox.RecipientTypeDetails)" -ForegroundColor White
    Write-Host ''

    if ($mailbox.RecipientTypeDetails -ne 'SharedMailbox') {
        Write-Log "El buzón ya es '$($mailbox.RecipientTypeDetails)'. No se requiere acción." 'WARN'
        return $true
    }

    Write-Log "El buzón es SharedMailbox. Esto impide que aparezca en Teams." 'WARN'
    $confirm = Read-Host '¿Convertir a UserMailbox (Regular)? (S/N)'
    if ($confirm -notmatch '^[Ss]') {
        Write-Log 'Operación cancelada. Sin cambios.' 'INFO'
        return $false
    }

    if (-not $PSCmdlet.ShouldProcess($UserIdentity, "Set-Mailbox -Type Regular")) {
        Write-Log 'WhatIf activo. Sin cambios.' 'INFO'
        return $false
    }

    Write-Log 'Aplicando cambio...'
    try {
        Set-Mailbox -Identity $UserIdentity -Type Regular -ErrorAction Stop
        Write-Log "Set-Mailbox -Type Regular ejecutado." 'OK'
    }
    catch {
        Write-Log "Error al convertir el buzón: $($_.Exception.Message)" 'ERROR'
        return $false
    }

    $timeout = 60; $elapsed = 0; $interval = 5
    $current = $null
    while ($elapsed -lt $timeout) {
        Start-Sleep -Seconds $interval
        $elapsed += $interval
        try { $current = Get-Mailbox -Identity $UserIdentity -ErrorAction Stop }
        catch { $current = $null }
        if ($current -and $current.RecipientTypeDetails -eq 'UserMailbox') { break }
        Write-Log "Esperando confirmación... ($elapsed/$timeout s)"
    }

    if ($current -and $current.RecipientTypeDetails -eq 'UserMailbox') {
        Write-Log 'Cambio confirmado.' 'OK'
        Write-Log 'Sincronización con Teams: 15-60 minutos. Verifica licencia de Teams asignada.' 'WARN'
        return $true
    }

    Write-Log "No se pudo confirmar el cambio tras $timeout s. Revisa manualmente." 'ERROR'
    return $false
}

# --- ENTRADA ---

try {
    Show-Header
    Assert-ExchangeReady

    if ($Identity) {
        [void](Convert-SharedToRegular -UserIdentity $Identity)
    }
    else {
        do {
            Write-Host ''
            $target = Read-Host 'Email/UPN del usuario (vacío = salir)'
            if ([string]::IsNullOrWhiteSpace($target)) { break }

            [void](Convert-SharedToRegular -UserIdentity $target.Trim())

            Write-Host ''
            $more = Read-Host '¿Comprobar otro usuario? (S/N)'
            if ($more -notmatch '^[Ss]') { break }
        } while ($true)
    }
}
catch {
    Write-Log $_.Exception.Message 'ERROR'
}
