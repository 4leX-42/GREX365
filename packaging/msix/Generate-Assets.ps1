<#
.SYNOPSIS
    Genera los PNG requeridos por Package.appxmanifest desde un texto base.

.DESCRIPTION
    Crea Square44x44Logo, Square150x150Logo, Wide310x150Logo y StoreLogo en
    packaging/msix/assets usando System.Drawing.Common (solo Windows).

    Reemplaza estos PNG por activos definitivos de marca cuando esten listos.
    Es un placeholder funcional para que el manifest valide.

.EXAMPLE
    pwsh -File packaging/msix/Generate-Assets.ps1
#>
[CmdletBinding()]
param(
    [string]$OutDir = (Join-Path $PSScriptRoot 'assets'),
    [string]$Letter = 'G',
    [string]$BackgroundHex = '#2A6CFF',
    [string]$ForegroundHex = '#FFFFFF'
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

if (-not (Test-Path $OutDir)) {
    New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
}

function New-Tile {
    param(
        [string]$Path,
        [int]$Width,
        [int]$Height
    )

    $bmp = New-Object System.Drawing.Bitmap $Width, $Height
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAlias

    $bgColor = [System.Drawing.ColorTranslator]::FromHtml($BackgroundHex)
    $fgColor = [System.Drawing.ColorTranslator]::FromHtml($ForegroundHex)
    $g.Clear($bgColor)

    $fontSize = [Math]::Max(8, [Math]::Floor([Math]::Min($Width, $Height) * 0.55))
    $font = New-Object System.Drawing.Font 'Segoe UI', $fontSize, ([System.Drawing.FontStyle]::Bold), ([System.Drawing.GraphicsUnit]::Pixel)
    $brush = New-Object System.Drawing.SolidBrush $fgColor

    $format = New-Object System.Drawing.StringFormat
    $format.Alignment = [System.Drawing.StringAlignment]::Center
    $format.LineAlignment = [System.Drawing.StringAlignment]::Center

    $rect = New-Object System.Drawing.RectangleF 0, 0, $Width, $Height
    $g.DrawString($Letter, $font, $brush, $rect, $format)

    $g.Dispose()
    $bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    Write-Host "Wrote $Path"
}

New-Tile -Path (Join-Path $OutDir 'Square44x44Logo.png')   -Width 44  -Height 44
New-Tile -Path (Join-Path $OutDir 'Square150x150Logo.png') -Width 150 -Height 150
New-Tile -Path (Join-Path $OutDir 'Wide310x150Logo.png')   -Width 310 -Height 150
New-Tile -Path (Join-Path $OutDir 'StoreLogo.png')         -Width 50  -Height 50
