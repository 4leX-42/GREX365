# Grex365.SamplePlugin

POC plugin externo para Grex365 v2.0. Demuestra el contrato `IModule` (Fase 4 del Plantamiento).

## Que hace

Anyade una entrada de navegacion **Sample Hello** que muestra un mensaje + boton refrescar. La logica vive en un servicio inyectado por DI (`HelloService`) que el modulo registra en el contenedor del host al cargarse.

## Build

```powershell
dotnet build samples/Grex365.SamplePlugin/Grex365.SamplePlugin.csproj -c Release
```

Salida: `samples/Grex365.SamplePlugin/bin/Release/net10.0-windows/Grex365.SamplePlugin.dll` (solo el DLL del plugin, sin copias de `Grex365.Core.dll` ni `CommunityToolkit.Mvvm.dll`).

## Instalar

Copia el DLL al directorio de plugins del usuario:

```powershell
$dst = Join-Path $env:LOCALAPPDATA 'Grex365\plugins'
New-Item -ItemType Directory -Force -Path $dst | Out-Null
Copy-Item samples/Grex365.SamplePlugin/bin/Release/net10.0-windows/Grex365.SamplePlugin.dll $dst -Force
```

Reinicia Grex365. La entrada **Sample Hello** aparecera al final de la navegacion lateral.

## Como esta hecho

- `HelloModule` implementa `Grex365.Core.Plugins.IModule` y declara `Title`, `Glyph`, `ViewModelType`, `ViewType`.
- `RegisterServices` registra `HelloService` como singleton del scope DI del host.
- `HelloViewModel` recibe `HelloService` por constructor (instanciado por `IServiceProvider.GetRequiredService`).
- `HelloView` es un `UserControl` WPF estandar; nada de wpf-ui para no obligar a redistribuir dependencias del host.

## Decisiones de empaquetado

El csproj fuerza que el DLL no arrastre copias del runtime ni de las dependencias compartidas con el host:

- `CopyLocalLockFileAssemblies=false` impide que la carpeta de output incluya el grafo transitivo.
- Los `PackageReference` usan `ExcludeAssets=runtime` + `PrivateAssets=all`.
- El `ProjectReference` a `Grex365.Core` usa `Private=false` + `ExcludeAssets=runtime`.

En runtime el `AssemblyLoadContext` del plugin resuelve `Grex365.Core`, `CommunityToolkit.Mvvm` y `Microsoft.Extensions.DependencyInjection.Abstractions` via fallback al ALC por defecto, donde el host ya las cargo. Esto evita duplicar tipos y mismatches de version.

Si un plugin necesita una dependencia que el host **no** carga, debe distribuir ese DLL adicional junto al suyo en `%LOCALAPPDATA%\Grex365\plugins\`.
