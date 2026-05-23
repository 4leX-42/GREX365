# GREX365 v2.0

Toolkit de administración Microsoft 365 para sysadmin. Operaciones sobre Exchange Online, Microsoft Graph y Entra ID. Autenticación por certificado (app-only) o device code (delegado). Auditorías de seguridad. Bulk CSV. RBAC opcional.

> **Estado del proyecto (2026-05)**: reescritura **C# .NET 10 + WPF + WPF-UI (Fluent)** activa en branch `grex365-2.0`. La versión PowerShell (`GREX365/`) sigue funcional para fallback durante la migración. La nueva app reemplaza al toolkit legacy en producción.
>
> Documentación operativa:
> - [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) — stack, capas, patrones, decisiones técnicas.
> - [`docs/RUNBOOK.md`](docs/RUNBOOK.md) — manual operativo: instalación, primer arranque, troubleshooting, plugins, telemetría.
> - [`docs/ROADMAP.md`](docs/ROADMAP.md) — punch list completo H0-H6, hitos + status.
> - [`docs/MIGRATION.md`](docs/MIGRATION.md) — mapeo legacy PS → nuevo .NET, feature-by-feature.
> - [`PACKAGING.md`](PACKAGING.md) — empaquetado MSIX, AppInstaller, Intune.
> - [`PROGRESS.md`](PROGRESS.md) — bitácora autoritativa de sesiones y commits.

---

## Capacidades (v2.0)

15 módulos de navegación + plugins:

| Módulo | Qué hace | Conexión requerida |
|---|---|---|
| Dashboard | Hero + acciones rápidas + summary findings | — |
| Conexión | Manual cert / device-code + autoinstall EXO | — |
| Licencias | Tarjetas por SKU con utilization + filtro + deep-link a Usuarios | Graph |
| Usuarios | Search debounced + bulk CSV (`assign:<SKU>`) + RBAC | Graph |
| Grupos | Bulk M365/DL con override + Graph replica retry | Graph |
| Onboarding | Wizard creación usuario + UPN/password/usage validation | Graph |
| Offboarding | Disable + revoke sessions + convert to shared + remove licenses + OOO + GAL hide | Graph + EXO |
| Buzones compartidos | Apply / convert / FullAccess / SendAs / SendOnBehalf | Graph + EXO |
| Reglas de buzón | OOO + forwarding + calendar permissions | Graph + EXO |
| Flujo de correo | `Get-TransportRule` viewer con filtro | EXO |
| Auditoría | 12 analizadores de seguridad (MFA / CA policies / privileged / OAuth / etc.) + export HTML/JSON/CSV + baseline diff | Graph + EXO |
| Registro de auditoría | JSONL append-only de ops privilegiadas + métricas + filtros | — |
| Consola PS | REPL embebido (mismo runspace que app) | — |
| Asistente cert | Self-signed RSA 2048 + auto-create App Registration (Graph + EXO + admin-consent URL) + Export PFX | — |
| Comprobación DNS | MX / SPF / DKIM / DMARC | — |

---

## Quick start

### Build local desde código fuente

Requisitos: Windows 10 1809+ o Windows 11, .NET SDK 10.

```powershell
git clone https://github.com/4leX-42/GREX365.git
cd GREX365
git checkout grex365-2.0

# Restore + build
dotnet build src/Grex365.slnx -c Release

# Run tests
dotnet test src/Grex365.slnx -c Release

# Run app
dotnet run --project src/Grex365.App -c Debug
```

### Publish single-file `.exe` (portable)

```powershell
dotnet publish src/Grex365.App/Grex365.App.csproj `
  -c Release -r win-x64 --self-contained `
  -p:PublishSingleFile=true `
  -p:IncludeAllContentForSelfExtract=true `
  -p:EnableCompressionInSingleFile=true
```

Output: `bin/Release/net10.0-windows/win-x64/publish/Grex365.exe` (~70-90 MB, sin .NET preinstalado).

### Distribución MSIX (Intune / SCCM)

Ver [`PACKAGING.md`](PACKAGING.md).

---

## Datos en disco

```
%LOCALAPPDATA%\Grex365\
├── config\
│   ├── user_preferences.json     # tenant lock, theme, last nav, window pos, plugins, log level
│   └── exo-app-params.json       # AppId, TenantId, Org, cert thumbprint (sin clave privada)
├── logs\
│   └── grex365-YYYY-MM-DD.log    # Serilog rotativo (30 días)
├── audit\
│   └── grex365-YYYY-MM.jsonl     # privileged-action log append-only
└── plugins\
    └── *.dll                     # plugins externos (IModule)
```

Operación detallada + troubleshooting en [`docs/RUNBOOK.md`](docs/RUNBOOK.md).

---

## Tests

**532 verdes** (401 Core + 131 App) — xUnit + FluentAssertions + Moq.

```powershell
dotnet test src/Grex365.slnx --logger trx
```

CI corre en GitHub Actions (`.github/workflows/ci.yml`) en cada push y PR.

---

## Modelo de seguridad

- **Cert local**: clave privada en `CurrentUser\My` con ACL restringida al usuario actual + SYSTEM.
- **JSON metadata**: `exo-app-params.json` contiene `TenantId`, `AppId`, `Thumbprint`, `Organization`. Sin secretos.
- **TenantLock**: enforced post-conexión (cert + device-code). Aborta si tenant resultante no matchea.
- **RBAC opcional**: `RbacGuard` vía membership Entra group. Gating en VMs destructivas (Users / Groups / SharedMailbox / MailboxRules).
- **Auditoría JSONL**: cada op privilegiada escrita append-only en `%LOCALAPPDATA%/Grex365/audit/`.
- **Telemetría opt-in**: Application Insights connection string en preferencias. `NullTelemetry` por default.

---

## Permisos otorgados al ServicePrincipal (Asistente Cert)

**Graph (Application AppRoles)**:
- `User.ReadWrite.All` · `Group.ReadWrite.All` · `GroupMember.ReadWrite.All`
- `Directory.ReadWrite.All` · `Organization.Read.All`
- `RoleManagement.Read.Directory` · `UserAuthenticationMethod.ReadWrite.All`
- `Policy.Read.All` · `AuditLog.Read.All` · `Reports.Read.All` · `Application.Read.All`

**Exchange Online**:
- `Exchange.ManageAsApp` (AppRole sobre Office 365 Exchange Online)

**Roles directorio Entra**:
- Exchange Administrator · User Administrator · Groups Administrator

El admin debe abrir el URL de consent devuelto por el asistente y aprobar. La clave privada nunca sale del equipo.

---

## Stack

- **.NET 10** (LTS Nov 2025) · **C# 13** · **WPF + WPF-UI 4.3.0** (Fluent / Mica)
- **CommunityToolkit.Mvvm 8.4.2** (ObservableProperty / RelayCommand source-generated)
- **Microsoft.Extensions.Hosting 10.0.8** + DI 10.0.8
- **Serilog 10.0.0** (Extensions.Logging) + 7.0.0 (File sink) — rolling + ObservableLogSink UI
- **Microsoft.Graph SDK 5.x** (cert / device-code via Azure.Identity)
- **ExchangeOnlineManagement** PS module via `System.Management.Automation` runspace pool
- **Microsoft.ApplicationInsights 2.23.0** (opt-in telemetry)
- **xUnit + FluentAssertions + Moq + coverlet** (tests)

Decisiones rechazadas (WinUI 3 / Avalonia / Blazor / Electron / Prism / Velopack) documentadas en [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) §2.

---

## Plugins (extensibilidad)

Drop DLL en `%LOCALAPPDATA%\Grex365\plugins\` y reinicia. Cada DLL implementa `IModule` (en `Grex365.Core.Plugins`):

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

Cada DLL se carga en su propio `AssemblyLoadContext`. Failures se reportan WARN, no bloquean arranque. Habilitar/deshabilitar por DLL desde Settings.

POC compilable en [`samples/Grex365.SamplePlugin/`](samples/Grex365.SamplePlugin/).

---

## Troubleshooting rápido

Ver [`docs/RUNBOOK.md`](docs/RUNBOOK.md) §7 para escenarios completos. Quick:

| Síntoma | Solución |
|---|---|
| Status bar "Exchange: desconectado" | Connect view → Comprobar módulo EXO → Instalar |
| Cert thumbprint not found | Asistente cert → Crear App Registration + cert |
| Tenant Lock mismatch | Settings → actualizar Expected Tenant ID/Domain, o uncheck Enforce (no recomendado en prod) |
| Audit "Permiso faltante: AuditLog.Read.All" | Azure Portal → App registration → API permissions → grant admin consent |
| Tema no sigue Windows | Settings → Tema → "Auto (seguir sistema)" |
| RBAC bloquea op | Añadir UPN al grupo Entra configurado, o usar Cert auth (app-only bypassea RBAC by design) |

---

## Estructura del repo

```
GREX365-main_2/
├── src/
│   ├── Grex365.slnx              # solution file
│   ├── Grex365.Core/             # business logic, no UI ref
│   ├── Grex365.PowerShell/       # embedded PS runspace pool
│   └── Grex365.App/              # WPF UI (20 VMs, 16 Views, 9 converters)
├── tests/
│   ├── Grex365.Core.Tests/       # 401 tests
│   └── Grex365.App.Tests/        # 131 tests
├── samples/
│   └── Grex365.SamplePlugin/     # plugin POC
├── packaging/
│   └── msix/                     # Package.appxmanifest + Build-Msix.ps1 + appinstaller
├── docs/                         # ARCHITECTURE / RUNBOOK / ROADMAP / MIGRATION
├── .github/workflows/ci.yml      # GitHub Actions
├── GREX365/                      # legacy PS toolkit (durante migración)
├── PROGRESS.md                   # bitácora sesiones autoritativa
└── README.md
```

---

## Legacy (PowerShell toolkit modo previo)

La versión PS sigue funcional en `GREX365/`. Para usarla:

```powershell
git clone https://github.com/4leX-42/GREX365.git
cd GREX365
pwsh .\Main.ps1     # PS 7+ requerido (no Windows PowerShell 5.1)
```

Requisitos legacy:
- PowerShell 7.4 LTS o superior (`winget install --id Microsoft.PowerShell`)
- Módulos auto-instalados primera ejecución: `ExchangeOnlineManagement`, `Microsoft.Graph.*`

Estructura legacy y formato CSV detallados en [`docs/CSV-Schemas.html`](docs/CSV-Schemas.html). La versión PS será deprecada tras un release estable del nuevo .NET app.

---

## Contribución

- **Branch activa**: `grex365-2.0`. PRs contra esta branch.
- **Conventional commits**: `feat(scope): …`, `fix(scope): …`, `docs:`, `refactor:`, `test:`, `build:`, `ci:`, `ux:`.
- **Tests obligatorios** para lógica pura en `Grex365.Core` (xUnit + FluentAssertions).
- **VMs UI-agnostic**: nunca referencia WPF types directos — usa `IDialogService` / `IClipboardService` abstractions.

---

## Licencia

Uso interno Andersen. Sin licencia pública.
