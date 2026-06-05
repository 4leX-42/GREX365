# Packaging y despliegue · GREX365 v2.0

Este documento describe cómo empaquetar y distribuir GREX365 v2.0 dentro de la organización. Mapea contra la **Fase 5** del Plantamiento.

## 1. Build local desde código fuente

```powershell
dotnet build src/Grex365.App/Grex365.App.csproj -c Release
dotnet run --project src/Grex365.App/Grex365.App.csproj
```

Requisitos:
- .NET SDK 10 (preview o GA)
- Windows 10 1809+ o Windows 11
- PowerShell 7 instalado (para módulos PSExchange/MgGraph) — la app embebe runspaces, no usa `pwsh.exe`

## 2. EXE portable (self-contained, sin instalador)

Se entrega como **single-file executable** firmado, copiable a cualquier ruta.

```powershell
dotnet publish src/Grex365.App/Grex365.App.csproj `
    -c Release -p:PublishProfile=win-x64-portable
```

Salida: `src/Grex365.App/bin/Release/net10.0-windows/win-x64/publish/Grex365.App.exe`

- Tamaño esperado: 150-200 MB (incluye runtime .NET, wpf-ui, Graph SDK, EXO modules helpers)
- No requiere instalación de .NET en la estación de destino
- Compresión activada (`EnableCompressionInSingleFile=true`); el primer arranque descomprime a `%TEMP%` (delay ~3-5s)

## 3. MSIX (recomendado para Intune/AppLocker)

El scaffold vive en `packaging/msix/`:

```
packaging/msix/
├── Package.appxmanifest      # manifest con Identity es.andersen.Grex365
├── Grex365.appinstaller      # plantilla auto-update (placeholders {{FEED_BASE_URI}}, {{VERSION}})
├── Generate-Assets.ps1       # crea PNGs (44, 150, 310x150, 50) — placeholders hasta tener arte definitivo
├── Build-Msix.ps1            # publish single-file + makeappx pack
└── assets/                   # tiles PNG generados
```

### Build local

Requisitos: Windows 10/11 + Windows SDK >= 10.0.17763 (incluye `makeappx.exe` y `signtool.exe`).

```powershell
pwsh -File packaging/msix/Build-Msix.ps1 -Version 2.0.0.0
# Salida: packaging/msix/out/Grex365.msix (sin firmar)
```

El script:
1. Lanza `dotnet publish` con el perfil portable.
2. Renombra `Grex365.App.exe` → `Grex365.exe` (lo declara así el manifest).
3. Copia el output + manifest + assets a un staging temporal.
4. Ejecuta `makeappx pack`.

### Firma

```powershell
signtool sign /fd SHA256 /a /f cert.pfx /p <pwd> `
  /tr http://timestamp.digicert.com /td SHA256 `
  packaging/msix/out/Grex365.msix
```

El `Publisher` del `Package.appxmanifest` (`CN=Andersen ES, OU=IT, O=Andersen Tax LLP, C=ES`) **debe coincidir** con el `Subject` del certificado. Si firmas con otro Publisher, edita el manifest antes de empaquetar.

### CI (release on tag `v*`)

`.github/workflows/ci.yml` define el job `msix`:

- Se dispara al pushear un tag con prefijo `v` (ej. `v2.0.0`).
- Resuelve la versión del tag y la inyecta en el manifest.
- Empaqueta sin firma; si los secrets `SIGN_CERT_PFX_B64` + `SIGN_CERT_PASSWORD` están configurados, firma con `signtool`.
- Renderiza `Grex365.appinstaller` sustituyendo `{{FEED_BASE_URI}}` con la variable `MSIX_FEED_BASE_URI` del repo (Settings → Variables).
- Sube `Grex365.msix` y `Grex365.appinstaller` como artifact.

### Distribución

Subir `.msix` + `.appinstaller` firmados a:

- Azure Blob Storage / share UNC HTTPS / Artifactory / GitHub Releases.
- Distribuir el `.appinstaller` (NO el `.msix` directamente); el cliente Windows comprueba updates al arrancar.

Intune: **Apps → Windows app (Win32) → MSIX** apuntando al `.appinstaller`.

## 3b. Velopack (instalador + auto-update) — canal elegido para distribución (ver docs/DISTRIBUTION.md)

Integrado en la app (paquete `Velopack` 1.2.0): `Program.Main` ejecuta `VelopackApp.Build().Run()`
antes de WPF (hooks de install/update; no-op en builds dev/portable). El feed se configura en
**Ajustes → Actualizaciones** (`UpdateFeedUrl`: repo GitHub de releases o URL HTTP); botones
"Buscar actualizaciones" / "Actualizar y reiniciar".

### Build local

```powershell
pwsh -File packaging/velopack/Build-Velopack.ps1 -Version 2.0.1
# Salida: packaging/velopack/out/  (Grex365-win-Setup.exe + full/delta .nupkg + RELEASES)
```

El script publica **sin** single-file (Velopack gestiona el layout; los deltas necesitan el árbol
expandido), con R2R, e invoca `vpk pack` (instala el tool global si falta).

### Firma (cert CA interna Andersen — coste cero)

```powershell
pwsh -File packaging/velopack/Build-Velopack.ps1 -Version 2.0.1 `
  -SignParams '/fd SHA256 /sha1 <thumbprint> /tr http://timestamp.digicert.com /td SHA256'
```

`vpk` firma exe+dlls+instalador con esos parámetros de signtool. En máquinas del dominio con la
CA interna en Trusted Publishers → cero avisos. Fuera del dominio el binario muestra SmartScreen
(asumido — decisión 2026-06-05, ver docs/DISTRIBUTION.md).

### Publicar release

Subir el contenido de `packaging/velopack/out/` como assets de una GitHub Release del repo de
binarios (p.ej. `grex365-releases`, repo público SIN código). Los clientes con
`UpdateFeedUrl=https://github.com/<org>/grex365-releases` reciben el update (delta si es posible)
desde Ajustes.

## 4. Plugins externos

`%LOCALAPPDATA%\Grex365\plugins\*.dll` se cargan al iniciar. Cada DLL puede contener una o varias clases que implementan `Grex365.Core.Plugins.IModule`:

```csharp
public sealed class MyPluginModule : IModule
{
    public string Title => "Mi módulo";
    public string Glyph => "";
    public Type ViewModelType => typeof(MyViewModel);
    public Type ViewType => typeof(MyView);
    public void RegisterServices(IServiceCollection services)
    {
        services.AddSingleton<IMyService, MyService>();
    }
}
```

Detalles:
- Cada DLL se carga en su propio `AssemblyLoadContext`
- Si el load falla (dependencia rota, mismatch de versión .NET), se loguea WARN y se ignora — la app sigue arrancando
- Las dependencias del plugin deben copiarse junto al DLL (no se hace probing en el GAC)

### POC de referencia: `samples/Grex365.SamplePlugin`

POC completo y compilable. Ver `samples/Grex365.SamplePlugin/README.md` para instrucciones de build e instalación. Resumen:

```powershell
dotnet build samples/Grex365.SamplePlugin/Grex365.SamplePlugin.csproj -c Release
$dst = Join-Path $env:LOCALAPPDATA 'Grex365\plugins'
New-Item -ItemType Directory -Force -Path $dst | Out-Null
Copy-Item samples/Grex365.SamplePlugin/bin/Release/net10.0-windows/Grex365.SamplePlugin.dll $dst -Force
```

Tras reiniciar la app aparece la entrada **Sample Hello** al final de la navegación lateral. El csproj demuestra el patrón correcto para empaquetar plugins sin duplicar el grafo transitivo del host (`CopyLocalLockFileAssemblies=false` + `ExcludeAssets=runtime` en `PackageReference`/`ProjectReference`).

## 5. CI/CD

`.github/workflows/ci.yml` corre en cada push a `main` o `grex365-2.0`:

- Restore + build de Core, PowerShell, App y Tests
- `dotnet test` con TRX logger
- Sube resultados como artefacto

Pendiente: paso de **publish + sign + release** para producir el MSIX en cada tag `v*`.

## 6. Datos en estación

```
%LOCALAPPDATA%\Grex365\
├── config\
│   ├── preferences.json        # tenant lock, theme, last nav
│   └── exo-app-params.json     # cert config (AppId, TenantId, Org, Thumbprint)
├── logs\
│   └── grex365-YYYY-MM-DD.log  # Serilog rotativo (30 días)
└── plugins\
    └── *.dll                   # plugins externos (Fase 4)
```

Para desinstalar limpio: borra esa carpeta + desinstala MSIX (o borra el `.exe` portable).
