# Roles module
# Lightweight client-side RBAC + UI mode toggle.
#
# Roles: viewer < operator < admin
#   - viewer:  read-only operations (audits, reports, dashboards). Blocks anything destructive.
#   - operator: standard daily ops (group membership, mailbox conversion, exports).
#   - admin:    can run offboarding wizard, bulk destructive runners, cert deletion.
# Default role: operator.
#
# UI Modes: support (default) | advanced
#   - support:  forces -DryRun on destructive ops where supported, extra confirmations.
#   - advanced: shows raw cmdlets, fewer confirmations, no forced dry-run.
#
# Persistence: stored under user_preferences.json keys: Role, UIMode.
# Note: this is governance UX, not a security boundary. Backed by preferences,
# bypassable by anyone who can edit the JSON.

# Note: rank lookup is function-local. Using $script: failed because Roles.ps1 is
# consumed by scripts invoked via `&` with Set-StrictMode Latest, where $script:
# rebinds to the child script's scope. Function-local is safe and trivial.

function Get-RoleRankTable {
    return @{
        'viewer'   = 1
        'operator' = 2
        'admin'    = 3
    }
}

function Get-CurrentRole {
    $ranks = Get-RoleRankTable
    $prefs = Get-UserPreferences
    if ($prefs.PSObject.Properties.Name -contains 'Role' -and $prefs.Role) {
        $r = [string]$prefs.Role
        if ($ranks.ContainsKey($r)) { return $r }
    }
    return 'operator'
}

function Set-CurrentRole {
    param([Parameter(Mandatory = $true)][ValidateSet('viewer','operator','admin')][string]$Role)
    Set-PreferenceValue -Key 'Role' -Value $Role
    Write-Log ("Rol cambiado a: $Role") -Level OK -Source 'Roles'
}

function Test-RoleAtLeast {
    param([Parameter(Mandatory = $true)][ValidateSet('viewer','operator','admin')][string]$Required)
    $ranks = Get-RoleRankTable
    $current = Get-CurrentRole
    return ($ranks[$current] -ge $ranks[$Required])
}

function Assert-Role {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('viewer','operator','admin')][string]$Required,
        [string]$Operation = 'esta operación'
    )
    if (-not (Test-RoleAtLeast -Required $Required)) {
        $current = Get-CurrentRole
        throw "Rol insuficiente para $Operation. Requerido='$Required', actual='$current'. Cambia en Preferencias > Rol."
    }
}

function Get-CurrentUIMode {
    $prefs = Get-UserPreferences
    if ($prefs.PSObject.Properties.Name -contains 'UIMode' -and $prefs.UIMode) {
        $m = [string]$prefs.UIMode
        if ($m -in @('support','advanced')) { return $m }
    }
    return 'support'
}

function Set-CurrentUIMode {
    param([Parameter(Mandatory = $true)][ValidateSet('support','advanced')][string]$Mode)
    Set-PreferenceValue -Key 'UIMode' -Value $Mode
    Write-Log ("UI Mode cambiado a: $Mode") -Level OK -Source 'Roles'
}

function Test-IsSupportMode  { return ((Get-CurrentUIMode) -eq 'support') }
function Test-IsAdvancedMode { return ((Get-CurrentUIMode) -eq 'advanced') }

function Confirm-DestructiveAction {
    param(
        [Parameter(Mandatory = $true)][string]$Operation,
        [string]$Target,
        [switch]$RequireType
    )

    if (Test-IsAdvancedMode -and -not $RequireType) {
        Write-Host ''
        Write-Indent
        Write-Host "Modo avanzado · ejecutando $Operation sobre $Target" -ForegroundColor Yellow
        return $true
    }

    Write-Host ''
    if (Get-Command Show-WarningBlock -ErrorAction SilentlyContinue) {
        Show-WarningBlock -Title ("Acción destructiva: $Operation") -Detail ("Objetivo: $Target")
    }

    if ($RequireType) {
        $confirm = Read-Input -Prompt "Escribe CONFIRMAR para continuar"
        return ($confirm -eq 'CONFIRMAR')
    }
    $confirm = Read-Input -Prompt '¿Continuar? (S/N)'
    return ($confirm -match '^[Ss]')
}
