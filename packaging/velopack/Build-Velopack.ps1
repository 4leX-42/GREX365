<#
.SYNOPSIS
    Empaqueta GREX365 con Velopack (instalador + paquetes de update full/delta).

.DESCRIPTION
    1. dotnet publish self-contained (SIN single-file: Velopack gestiona el layout y los deltas
       funcionan mejor con el árbol de ficheros expandido).
    2. vpk pack -> packaging/velopack/out/ (Grex365-win-Setup.exe + .nupkg full/delta + RELEASES).

    El feed de updates es la carpeta `out/` subida a GitHub Releases (repo binarios-only) o a un
    share HTTP interno. La app lee el feed de Ajustes > Actualizaciones (UpdateFeedUrl).

.PARAMETER Version
    Versión semver del paquete (ej. 2.0.1). Obligatoria.

.PARAMETER SignParams
    Opcional: parámetros para signtool (ej. '/fd SHA256 /sha1 <thumbprint> /t <tsa>') si hay
    certificado de la CA interna. Velopack firma exe+dlls con ellos. Vacío = sin firma.

.PARAMETER NoObfuscate
    Salta el paso de ofuscación (Obfuscar sobre Grex365.Core.dll + Grex365.PowerShell.dll).
    Por defecto se ofusca. La capa WPF (Grex365.App.dll) NUNCA se ofusca: XAML resuelve
    tipos y propiedades por nombre (bindings, x:Class, DataTemplates) y el renombrado la rompe.

.EXAMPLE
    pwsh -File packaging/velopack/Build-Velopack.ps1 -Version 2.0.1
#>
param(
    [Parameter(Mandatory = $true)][string]$Version,
    [string]$SignParams = '',
    [string]$OutDir = '',
    [switch]$NoObfuscate
)

$ErrorActionPreference = 'Stop'
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..' '..')
$publishDir = Join-Path $PSScriptRoot 'publish'
$outDir = if ($OutDir) { $OutDir } else { Join-Path $PSScriptRoot 'out' }

# vpk CLI presente?
if (-not (Get-Command vpk -ErrorAction SilentlyContinue)) {
    Write-Host 'vpk no encontrado. Instalando dotnet tool global...' -ForegroundColor Yellow
    dotnet tool install -g vpk
    if ($LASTEXITCODE -ne 0) { throw 'No se pudo instalar vpk (dotnet tool install -g vpk).' }
}

if (Test-Path $publishDir) { Remove-Item -Recurse -Force $publishDir }

Write-Host "1/3 dotnet publish (self-contained, R2R, sin single-file)..." -ForegroundColor Cyan
dotnet publish (Join-Path $repoRoot 'src/Grex365.App/Grex365.App.csproj') `
    -c Release -r win-x64 --self-contained true `
    -p:PublishSingleFile=false `
    -p:PublishReadyToRun=true `
    -p:Version=$Version `
    -o $publishDir
if ($LASTEXITCODE -ne 0) { throw 'dotnet publish falló.' }

# ---- 2/3 Obfuscar (Core + PowerShell; la capa WPF queda intacta) ----
# KeepPublicApi=true preserva todo lo público: los contratos que App.dll consume, los records
# que System.Text.Json (de)serializa por nombre y los tipos registrados en DI. Se renombra lo
# privado/interno (métodos, campos, tipos internos) — suficiente para que un ILSpy casual no
# lea la lógica de negocio con nombres significativos.
if (-not $NoObfuscate) {
    Write-Host "2/3 Obfuscar (Grex365.Core.dll + Grex365.PowerShell.dll)..." -ForegroundColor Cyan
    if (-not (Get-Command obfuscar.console -ErrorAction SilentlyContinue)) {
        Write-Host 'obfuscar.console no encontrado. Instalando dotnet tool global...' -ForegroundColor Yellow
        dotnet tool install -g Obfuscar.GlobalTool
        if ($LASTEXITCODE -ne 0) { throw 'No se pudo instalar Obfuscar.GlobalTool.' }
    }

    $obfOut = Join-Path $PSScriptRoot 'obfuscated'
    if (Test-Path $obfOut) { Remove-Item -Recurse -Force $obfOut }
    New-Item -ItemType Directory -Force -Path $obfOut | Out-Null

    $obfConfig = Join-Path $PSScriptRoot 'obfuscar.generated.xml'
    @"
<?xml version='1.0'?>
<Obfuscator>
  <Var name="InPath" value="$publishDir" />
  <Var name="OutPath" value="$obfOut" />
  <Var name="KeepPublicApi" value="true" />
  <Var name="HidePrivateApi" value="true" />
  <Var name="RenameProperties" value="true" />
  <Var name="RenameEvents" value="true" />
  <Var name="RenameFields" value="true" />
  <Var name="HideStrings" value="true" />
  <Var name="RegenerateDebugInfo" value="false" />
  <Module file="`$(InPath)\Grex365.Core.dll" />
  <Module file="`$(InPath)\Grex365.PowerShell.dll" />
</Obfuscator>
"@ | Set-Content $obfConfig -Encoding UTF8

    & obfuscar.console $obfConfig
    if ($LASTEXITCODE -ne 0) { throw 'Obfuscar falló.' }

    foreach ($dll in 'Grex365.Core.dll', 'Grex365.PowerShell.dll') {
        $src = Join-Path $obfOut $dll
        if (-not (Test-Path $src)) { throw "Obfuscar no produjo $dll." }
        Copy-Item $src (Join-Path $publishDir $dll) -Force
    }
    Remove-Item $obfConfig -Force
    Write-Host 'Ofuscadas y reemplazadas en publish/.' -ForegroundColor Green
} else {
    Write-Host '2/3 Obfuscar OMITIDO (-NoObfuscate).' -ForegroundColor Yellow
}

Write-Host "3/3 vpk pack v$Version..." -ForegroundColor Cyan
# packId DISTINTO del dir de datos: Velopack instala en %LocalAppData%\{packId} y el
# uninstall BORRA esa carpeta. Con packId 'Grex365' arrasaría %LocalAppData%\Grex365\{config,logs}.
$vpkArgs = @(
    'pack',
    '--packId', 'Grex365.App',
    '--packVersion', $Version,
    '--packDir', $publishDir,
    '--mainExe', 'Grex365.exe',
    '--packTitle', 'GREX365',
    '--packAuthors', 'Andersen',
    '--outputDir', $outDir
)
if ($SignParams) { $vpkArgs += @('--signParams', $SignParams) }
& vpk @vpkArgs
if ($LASTEXITCODE -ne 0) { throw 'vpk pack falló.' }

Write-Host "OK -> $outDir" -ForegroundColor Green
Get-ChildItem $outDir | Select-Object Name, @{n='MB';e={[math]::Round($_.Length/1MB,1)}} | Format-Table
