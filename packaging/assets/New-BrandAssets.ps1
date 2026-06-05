<#
.SYNOPSIS
    Genera los assets de marca de GREX365 desde el monograma GX (mismo trazado que GxLogoMark).

.DESCRIPTION
    Dibuja el monograma GX (G en D-shape + X) sobre pill con el gradiente de marca
    (#1E40AF -> #3B82F6) usando System.Drawing y emite:
      - src/Grex365.App/Assets/Grex365.ico   (16/24/32/48/64/128/256, PNG-compressed ICO)
      - src/Grex365.App/Assets/Splash.png    (460x260, splash nativo WPF)
      - packaging/msix/assets/*.png          (tiles MSIX brand-consistent)

    Re-ejecutar tras cambiar el monograma en App.xaml para mantener todo sincronizado.

.EXAMPLE
    pwsh -File packaging/assets/New-BrandAssets.ps1
#>
[CmdletBinding()]
param(
    [string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..' '..'))
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$gradTop = [System.Drawing.Color]::FromArgb(0xFF, 0x1E, 0x40, 0xAF)
$gradBottom = [System.Drawing.Color]::FromArgb(0xFF, 0x3B, 0x82, 0xF6)
$splashBg = [System.Drawing.Color]::FromArgb(0xFF, 0x0B, 0x12, 0x20)

# Dibuja el monograma GX (espacio de diseño 42x28, idéntico a GxLogoMark de App.xaml)
# centrado en un cuadrado de lado $size con pill de gradiente de fondo.
function Draw-GxBadge {
    param([System.Drawing.Graphics]$g, [float]$x, [float]$y, [float]$size, [bool]$pill = $true)

    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias

    if ($pill) {
        $radius = [float]($size * 0.24)
        $rect = New-Object System.Drawing.RectangleF($x, $y, $size, $size)
        $path = New-Object System.Drawing.Drawing2D.GraphicsPath
        $d = $radius * 2
        $path.AddArc($rect.X, $rect.Y, $d, $d, 180, 90)
        $path.AddArc($rect.Right - $d, $rect.Y, $d, $d, 270, 90)
        $path.AddArc($rect.Right - $d, $rect.Bottom - $d, $d, $d, 0, 90)
        $path.AddArc($rect.X, $rect.Bottom - $d, $d, $d, 90, 90)
        $path.CloseFigure()
        $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
            $rect, $gradTop, $gradBottom, [System.Drawing.Drawing2D.LinearGradientMode]::ForwardDiagonal)
        $g.FillPath($brush, $path)
        $brush.Dispose(); $path.Dispose()
    }

    # Monograma a ~62% del ancho, centrado
    $designW = 42.0; $designH = 28.0
    $scale = [float](($size * 0.62) / $designW)
    $offX = [float]($x + ($size - $designW * $scale) / 2)
    $offY = [float]($y + ($size - $designH * $scale) / 2)

    $state = $g.Save()
    $g.TranslateTransform($offX, $offY)
    $g.ScaleTransform($scale, $scale)

    $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 3.0)
    $pen.StartCap = 'Round'; $pen.EndCap = 'Round'
    $pen.LineJoin = 'Round'

    # G: M 17,3 L 6,3 A11,11(ccw) 6,25 L 17,25 A5,5(ccw) 22,20 L 22,14 L 13,14
    $gPath = New-Object System.Drawing.Drawing2D.GraphicsPath
    $gPath.AddLine(17.0, 3.0, 6.0, 3.0)
    $gPath.AddArc(-5.0, 3.0, 22.0, 22.0, 270.0, -180.0)   # semicírculo izquierdo, ccw
    $gPath.AddLine(6.0, 25.0, 17.0, 25.0)
    $gPath.AddArc(12.0, 15.0, 10.0, 10.0, 90.0, -90.0)    # esquina inferior-dcha de la G, ccw
    $gPath.AddLine(22.0, 20.0, 22.0, 14.0)
    $gPath.AddLine(22.0, 14.0, 13.0, 14.0)
    $g.DrawPath($pen, $gPath)
    $gPath.Dispose()

    # X
    $g.DrawLine($pen, 27.0, 3.0, 39.0, 25.0)
    $g.DrawLine($pen, 39.0, 3.0, 27.0, 25.0)

    $pen.Dispose()
    $g.Restore($state)
}

function New-BadgePng {
    param([int]$Size)
    $bmp = New-Object System.Drawing.Bitmap($Size, $Size)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    Draw-GxBadge -g $g -x 0 -y 0 -size $Size
    $g.Dispose()
    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    , $ms.ToArray()
}

# ---- 1. ICO multirresolución (entradas PNG) ----
$assetsDir = Join-Path $RepoRoot 'src/Grex365.App/Assets'
New-Item -ItemType Directory -Force -Path $assetsDir | Out-Null

$sizes = 16, 24, 32, 48, 64, 128, 256
$pngs = foreach ($s in $sizes) { , (New-BadgePng -Size $s) }

$icoPath = Join-Path $assetsDir 'Grex365.ico'
$fs = [System.IO.File]::Create($icoPath)
$bw = New-Object System.IO.BinaryWriter($fs)
$bw.Write([uint16]0); $bw.Write([uint16]1); $bw.Write([uint16]$sizes.Count)
$offset = 6 + 16 * $sizes.Count
for ($i = 0; $i -lt $sizes.Count; $i++) {
    $s = $sizes[$i]; $data = $pngs[$i]
    $bw.Write([byte]($(if ($s -ge 256) { 0 } else { $s })))   # width (0 = 256)
    $bw.Write([byte]($(if ($s -ge 256) { 0 } else { $s })))   # height
    $bw.Write([byte]0); $bw.Write([byte]0)                    # colors, reserved
    $bw.Write([uint16]1); $bw.Write([uint16]32)               # planes, bpp
    $bw.Write([uint32]$data.Length); $bw.Write([uint32]$offset)
    $offset += $data.Length
}
foreach ($data in $pngs) { $bw.Write($data) }
$bw.Dispose(); $fs.Dispose()
Write-Host "OK $icoPath" -ForegroundColor Green

# ---- 2. Splash 460x260 ----
$w = 460; $h = 260
$bmp = New-Object System.Drawing.Bitmap($w, $h)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = 'AntiAlias'
$g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
$g.Clear($splashBg)
# Línea de acento inferior con el gradiente de marca
$accentRect = New-Object System.Drawing.RectangleF(0, ($h - 4), $w, 4)
$accentBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
    $accentRect, $gradTop, $gradBottom, [System.Drawing.Drawing2D.LinearGradientMode]::Horizontal)
$g.FillRectangle($accentBrush, $accentRect); $accentBrush.Dispose()

Draw-GxBadge -g $g -x 186 -y 52 -size 88
$fontTitle = New-Object System.Drawing.Font('Segoe UI Semibold', 26, [System.Drawing.FontStyle]::Regular, 'Pixel')
$fontSub = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Regular, 'Pixel')
$white = [System.Drawing.Brushes]::White
$gray = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0xFF, 0x94, 0xA3, 0xB8))
$fmt = New-Object System.Drawing.StringFormat
$fmt.Alignment = 'Center'
$g.DrawString('GREX365', $fontTitle, $white, (New-Object System.Drawing.RectangleF(0, 158, $w, 50)), $fmt)
$g.DrawString('Microsoft 365 toolkit', $fontSub, $gray, (New-Object System.Drawing.RectangleF(0, 205, $w, 30)), $fmt)
$fontTitle.Dispose(); $fontSub.Dispose(); $gray.Dispose(); $g.Dispose()
$splashPath = Join-Path $assetsDir 'Splash.png'
$bmp.Save($splashPath, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
Write-Host "OK $splashPath" -ForegroundColor Green

# ---- 3. Tiles MSIX (mismo look que el resto de assets) ----
$msixDir = Join-Path $RepoRoot 'packaging/msix/assets'
New-Item -ItemType Directory -Force -Path $msixDir | Out-Null

function New-Tile {
    param([string]$Name, [int]$W, [int]$H)
    $bmp = New-Object System.Drawing.Bitmap($W, $H)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    # Fondo transparente; badge centrado al 80% del lado menor
    $side = [Math]::Min($W, $H) * 0.8
    Draw-GxBadge -g $g -x (($W - $side) / 2) -y (($H - $side) / 2) -size $side
    $g.Dispose()
    $out = Join-Path $msixDir $Name
    $bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    Write-Host "OK $out" -ForegroundColor Green
}

New-Tile 'Square44x44Logo.png' 44 44
New-Tile 'Square150x150Logo.png' 150 150
New-Tile 'Wide310x150Logo.png' 310 150
New-Tile 'StoreLogo.png' 50 50
