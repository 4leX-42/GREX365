#requires -Version 7.4
# Headless validator for Start-Gui.ps1.
# - Loads the GUI script with ShowDialog stubbed.
# - Verifies that every named control and click handler exists.
# - Programmatically clicks every button and checks no exception is raised.
# - Drains the IPC queues and reports results.
# Exit code: 0 = all good, 1 = failures detected.

[CmdletBinding()]
param([switch]$Detailed)

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$guiPath = Join-Path $here 'Start-Gui.ps1'
if (-not (Test-Path $guiPath)) { Write-Host "GUI no encontrado: $guiPath" -ForegroundColor Red; exit 1 }

# Patch source: replace ShowDialog with no-op so the window does not block.
$source = Get-Content $guiPath -Raw
$patched = $source -replace '\$null\s*=\s*\$Window\.ShowDialog\(\)', '# ShowDialog stubbed by Test-Gui'

# Inject the GUI dir via env so the patched script can resolve modules even though
# we are executing it from a scriptblock (no PSScriptRoot / $MyInvocation.Path).
$env:GREX365_GUI_DIR = $here

# Run the patched script in this process so we get $Window in scope after.
$sb = [scriptblock]::Create($patched)
$failures = 0
function Fail { param([string]$Msg); Write-Host "FAIL $Msg" -ForegroundColor Red; $script:failures++ }
function Pass { param([string]$Msg); Write-Host "OK   $Msg" -ForegroundColor Green }

try {
    . $sb
} catch {
    Fail "GUI script throws on load: $($_.Exception.Message)"
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    exit 1
}

# At this point $Window etc are in scope.
if (-not $Window) { Fail '$Window not set'; exit 1 }
Pass 'Window loaded'

# Verify expected named elements
$expected = @(
    'HeaderSubtitle','BtnConnect','BtnDisconnect','SideNav',
    'CardTenant','CardAccount','CardGraph','CardExo',
    'QuickHealth','QuickAudit','QuickOpenLogs','QuickOpenReports',
    'PanelWelcome','PanelHealth','PanelAudit','PanelGroups','PanelPermissions','PanelOffboarding','PanelPrefs',
    'BtnRunHealth','HealthStatus','BtnRunAudit','AuditStatus',
    'GroupsCsvPath','GroupsBrowse','BtnRunGroups','GroupsStatus',
    'PermsCsvPath','PermsBrowse','BtnRunPerms','PermsStatus',
    'OffUpn','OffDelegate','OffManager','OffOrg','OffTemplate','OffDryRun','BtnRunOff','OffStatus',
    'PrefMethod','PrefRole','PrefUiMode','BtnSavePrefs','PrefStatus',
    'LogList','OpStatusText','BtnClearLog',
    'SbGraph','SbExo','SbTenant','SbAccount','SbClock','DotGraph','DotExo'
)
foreach ($n in $expected) {
    $e = $Window.FindName($n)
    if ($null -eq $e) { Fail "FindName($n) returns null" } else { if ($Detailed) { Pass "FindName($n)" } }
}
if (-not $Detailed) { Pass "FindName: $($expected.Count) controls" }

# Side-nav panel switching
foreach ($tag in 'welcome','health','audit','groups','permissions','offboarding','prefs') {
    try {
        Show-Panel -Tag $tag
        $panel = $panels[$tag]
        if ($panel.Visibility -ne 'Visible') { Fail "Show-Panel($tag) did not become Visible" }
    } catch { Fail "Show-Panel($tag) threw: $($_.Exception.Message)" }
}
Pass 'Show-Panel switching'

# Click each button — defer to RaiseEvent so we go through real handlers.
# Buttons that start runspaces should not throw immediately.
$buttonsToClick = @(
    'BtnConnect','BtnDisconnect',
    'BtnRunHealth','BtnRunAudit',
    'BtnClearLog','QuickHealth','QuickAudit'
)
foreach ($bname in $buttonsToClick) {
    $btn = $Window.FindName($bname)
    if (-not $btn) { Fail "Cannot click $bname (not found)"; continue }
    try {
        $args2 = New-Object System.Windows.RoutedEventArgs ([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)
        $btn.RaiseEvent($args2)
        Pass "Click $bname"
    } catch {
        Fail "Click $bname threw: $($_.Exception.Message)"
    }
}

# BtnRunOff: requires UPN — should not throw, should show 'UPN y delegado son obligatorios'
try {
    $offBtn = $Window.FindName('BtnRunOff')
    $args2 = New-Object System.Windows.RoutedEventArgs ([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)
    $offBtn.RaiseEvent($args2)
    Pass 'Click BtnRunOff (empty form)'
} catch { Fail "BtnRunOff empty threw: $($_.Exception.Message)" }

# Preferences save with default selections
Write-Host "DBG entering BtnSavePrefs block" -ForegroundColor DarkGray
try {
    $saveBtn = $Window.FindName('BtnSavePrefs')
    if (-not $saveBtn) { Fail 'FindName BtnSavePrefs returned null' }
    else {
        $args3 = New-Object System.Windows.RoutedEventArgs ([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)
        $saveBtn.RaiseEvent($args3)
        $prefStat = $Window.FindName('PrefStatus')
        $afterText = if ($prefStat) { [string]$prefStat.Text } else { '(no PrefStatus)' }
        Pass "Click BtnSavePrefs (PrefStatus='$afterText')"
    }
} catch { Fail "BtnSavePrefs threw: $($_.Exception.Message)" }

# Wait for runspaces to settle and drain queues
Start-Sleep -Milliseconds 1500
$msgs = New-Object System.Collections.Generic.List[object]
$tmp = $null
while ($SyncHash.LogQueue.TryDequeue([ref]$tmp)) { $msgs.Add($tmp) }

# Filter for runspace-level failures (Connect/Disconnect/etc will fail in test env without Graph context — that's expected)
$jobFail = @($msgs | Where-Object { $_.msg -and $_.lvl -eq 'ERROR' -and $_.msg -match 'fallido' })
$jobOk   = @($msgs | Where-Object { $_.msg -and $_.lvl -eq 'OK'    -and $_.msg -match 'completado|GUI iniciada' })
Pass "Runspaces drained: $($msgs.Count) eventos · $($jobOk.Count) OK · $($jobFail.Count) ERROR-esperados-sin-tenant"

# Close window cleanly
try {
    $Window.Close()
    Pass 'Window.Close()'
} catch { Fail "Window.Close threw: $($_.Exception.Message)" }

Write-Host ''
if ($failures -eq 0) {
    Write-Host "All Test-Gui checks passed." -ForegroundColor Green
    exit 0
} else {
    Write-Host "Failures: $failures" -ForegroundColor Red
    exit 1
}
