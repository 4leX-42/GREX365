#requires -Version 7.4
[CmdletBinding()]
param()

# GUI bootstrap. Verifies PS 7+ and forwards to GREX365/GUI/Start-Gui.ps1.

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host ''
    Write-Host '  GREX365 GUI requiere PowerShell 7.4+.' -ForegroundColor Red
    Write-Host ('  Versión detectada: {0}' -f $PSVersionTable.PSVersion) -ForegroundColor Yellow
    Write-Host '  Instala con: winget install --id Microsoft.PowerShell' -ForegroundColor DarkGray
    exit 1
}

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$guiScript = Join-Path $here 'GREX365\GUI\Start-Gui.ps1'
if (-not (Test-Path -LiteralPath $guiScript)) {
    Write-Host "GUI no encontrada: $guiScript" -ForegroundColor Red
    exit 1
}

try { $PSStyle.Progress.View = 'Minimal' } catch {}
try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = [System.Text.Encoding]::UTF8
} catch {}

# Pre-emptively unblock GUI script + modules
Get-ChildItem -Path (Join-Path $here 'GREX365') -Recurse -Filter *.ps1 -File -ErrorAction SilentlyContinue |
    Unblock-File -ErrorAction SilentlyContinue

& $guiScript
