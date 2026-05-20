<#
.SYNOPSIS
    Empaqueta Grex365 publicado en un .msix usando makeappx (Windows SDK).

.DESCRIPTION
    Orquesta:
      1. Publish single-file de Grex365.App (-> bin\Release\...\publish)
      2. Copia el output + Package.appxmanifest + assets en un staging
      3. Renombra Grex365.App.exe -> Grex365.exe (manifest lo declara asi)
      4. Genera assets con Generate-Assets.ps1 si no existen
      5. makeappx pack /d $staging /p $output

    No firma. La firma con signtool va en un step CI o local separado.

.PARAMETER Configuration
    Configuracion de build (default: Release).

.PARAMETER Version
    Numero de version 4 octetos para sobrescribir el de Package.appxmanifest.
    Si no se pasa, se toma del manifest sin tocar.

.PARAMETER OutFile
    Ruta del .msix a generar. Por defecto packaging/msix/out/Grex365.msix.

.EXAMPLE
    pwsh -File packaging/msix/Build-Msix.ps1 -Version 2.0.0.0
#>
[CmdletBinding()]
param(
    [string]$Configuration = 'Release',
    [string]$Version,
    [string]$OutFile = (Join-Path $PSScriptRoot 'out/Grex365.msix')
)

$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..\..')
$appProj  = Join-Path $repoRoot 'src/Grex365.App/Grex365.App.csproj'
$manifest = Join-Path $PSScriptRoot 'Package.appxmanifest'
$assets   = Join-Path $PSScriptRoot 'assets'

# 1) Publish single-file
Write-Host '[1/5] dotnet publish (single-file portable)'
& dotnet publish $appProj -c $Configuration -p:PublishProfile=win-x64-portable
if ($LASTEXITCODE -ne 0) { throw "dotnet publish exited with $LASTEXITCODE" }
$publishDir = Join-Path $repoRoot 'src/Grex365.App/bin/Release/net10.0-windows/win-x64/publish'
if (-not (Test-Path $publishDir)) { throw "Publish dir not found: $publishDir" }

# 2) Stage
$staging = Join-Path $env:TEMP ("Grex365-msix-staging-" + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $staging -Force | Out-Null
Write-Host "[2/5] Staging in $staging"

Copy-Item (Join-Path $publishDir '*') -Destination $staging -Recurse -Force

# Rename .exe to match manifest's Executable attribute
$exeSrc = Join-Path $staging 'Grex365.App.exe'
$exeDst = Join-Path $staging 'Grex365.exe'
if (Test-Path $exeSrc) {
    Move-Item $exeSrc $exeDst -Force
} elseif (-not (Test-Path $exeDst)) {
    throw "Neither Grex365.App.exe nor Grex365.exe present in publish output"
}

# 3) Assets
if (-not (Test-Path (Join-Path $assets 'Square44x44Logo.png'))) {
    Write-Host '[3/5] Generating placeholder assets'
    & (Join-Path $PSScriptRoot 'Generate-Assets.ps1')
} else {
    Write-Host '[3/5] Reusing existing assets'
}
Copy-Item -Path (Join-Path $assets '*.png') -Destination (Join-Path $staging 'assets') -Force -Recurse:$false -Container:$false -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Path (Join-Path $staging 'assets') -Force | Out-Null
Copy-Item -Path (Join-Path $assets '*.png') -Destination (Join-Path $staging 'assets') -Force

# 4) Manifest (with optional version override)
$manifestDst = Join-Path $staging 'AppxManifest.xml'
if ($Version) {
    [xml]$xml = Get-Content $manifest
    $xml.Package.Identity.Version = $Version
    $xml.Save($manifestDst)
} else {
    Copy-Item $manifest $manifestDst -Force
}

# 5) makeappx pack
$makeAppx = (Get-Command makeappx.exe -ErrorAction SilentlyContinue)?.Source
if (-not $makeAppx) {
    $sdkRoot = 'C:\Program Files (x86)\Windows Kits\10\bin'
    if (Test-Path $sdkRoot) {
        $makeAppx = Get-ChildItem -Path $sdkRoot -Recurse -Filter 'makeappx.exe' -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -like '*\x64\*' } |
            Sort-Object FullName -Descending |
            Select-Object -First 1 -ExpandProperty FullName
    }
}
if (-not $makeAppx) {
    throw 'makeappx.exe no encontrado. Instala Windows SDK (>=10.0.17763) o anyade el bin\x64 al PATH.'
}

$outDir = Split-Path $OutFile -Parent
New-Item -ItemType Directory -Path $outDir -Force | Out-Null
Write-Host "[5/5] makeappx pack -> $OutFile"
& $makeAppx pack /d $staging /p $OutFile /o
if ($LASTEXITCODE -ne 0) { throw "makeappx pack exited with $LASTEXITCODE" }

Write-Host ''
Write-Host "MSIX generado: $OutFile"
Write-Host "Para firmar:   signtool sign /fd SHA256 /a /f cert.pfx /p <pwd> `"$OutFile`""
