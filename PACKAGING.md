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

> ⚠️ Pendiente de completar. Tracking en `PROGRESS.md` → Fase 5.

Pasos previstos (no implementados aún):

1. Crear proyecto `Grex365.App.Package` (Windows Application Packaging Project, `.wapproj`) en Visual Studio
2. `Package.appxmanifest` con:
   - `Identity Name="es.andersen.Grex365"` `Publisher="CN=Andersen, ..."`
   - `Capabilities` mínimas (no se necesita `runFullTrust` salvo para RunspacePool en local)
   - `Application` apuntando a `Grex365.App.exe`
3. Firmar con certificado de código (EV preferible)
4. Generar `.msix` y `.appinstaller` con feed URL interno
5. Subir `.msix` y `.appinstaller` a un share interno o a Azure Blob Storage / GitHub Releases
6. Distribuir vía **Intune** (App > Windows app (Win32) → MSIX) o **SCCM**

Auto-update: el cliente App Installer comprueba el `.appinstaller` al iniciar y aplica updates de forma transparente. Configuración recomendada:

```xml
<UpdateSettings>
  <OnLaunch HoursBetweenUpdateChecks="24" UpdateBlocksActivation="false" ShowPrompt="true" />
  <AutomaticBackgroundTask />
</UpdateSettings>
```

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
