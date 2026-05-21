# GREX365 v2.0 — Progress log

> **Documento maestro de seguimiento.** Mapea el estado del proyecto contra `Plantamiento_arquitectura_de_la_herramienta.md` (roadmap arquitectónico) y `deep-research-report.md` (research técnico). Toda feature shipped y todo pendiente vive aquí.

- Branch: `grex365-2.0` · Pushed up to `origin/grex365-2.0`
- Stack actual: **C# · .NET 10 · WPF + wpf-ui (Fluent) · MVVM (CommunityToolkit.Mvvm) · Serilog · Microsoft.Extensions.Hosting · Microsoft.ApplicationInsights**
- Tests: **182 passing** (xUnit + FluentAssertions)
- Última actualización: 2026-05-21

## Bitácora sesiones

### 2026-05-21 — Sesión "cert auto + audit + telemetry" (8 commits)

Tests **161 → 174** (+13). Fase 6 ítem "Application Insights" cerrado. Auto-cert end-to-end.

Commits:
- `fix(ui)` IconFontFamily resource (String → FontFamily) — bloqueaba ya glyphs en nav
- `fix(powershell)` todos los scripts PS declaran `param()` para que `AddParameter` enlace correctamente
- `fix(ui)` feedback visual claro para nav items deshabilitados (Opacity 0.35 + tooltip "Conecta Graph/Exchange")
- `feat(connect)` device-code auth para Microsoft Graph (cliente de Azure CLI, no requiere App Registration previa)
- `feat(connect+certwizard)` auto-install EXO module (Start-Process `pwsh.exe` externo para evitar ACL WindowsApps) + auto-create App Registration con todos los permisos Graph/EXO + admin consent URL
- `fix(powershell)` runspace pool sanity: drop probe cross-runspace de EXO (no compartida), unwrap PSObject solo para tipos primitivos, sanitizar `PSModulePath` quitando WindowsApps
- `feat(audit)` grupos sin actividad reciente via `/reports/getOffice365GroupsActivityDetail` (period D7/D30/D90/D180); CSV parse puro + analyzer; UI con NumberBox umbral. Suma permiso `Reports.Read.All` al App Reg auto-create.
- `feat(telemetry)` Application Insights opt-in: `ITelemetry`/`NullTelemetry` en Core, `ApplicationInsightsTelemetry` en App; conn string en Settings (vacío = NullTelemetry); UiLogSink envía cada Ok/Warn/Error como TrackEvent/TrackException.
- `feat(rbac)` permisos por rol vía `Me.CheckMemberGroups`. `IRbacGuard` + `IMembershipChecker` (pure-testable); Settings textbox `AuthorizationGroupId`; Offboarding gateado con audit WARN; cache invalidado en Disconnect.

Plantamiento status: Fase 6 ahora **6/8 hechos** (App Insights, métricas, audit JSONL, viewer UI, log level configurable, **RBAC vía membership Entra**). Falta: docs internas, QA escenarios.

---

### 2026-05-20 — Sesión 9 commits (Fase 4 cerrada, Fase 5 MSIX scaffold, Fase 6 avanza)

Tests **149 → 161** (+12). Módulos UI **13 → 14** (Mail flow añadido).

Commits:
- `feat(plugins)` SamplePlugin POC (`samples/Grex365.SamplePlugin`) + CI build artifact
- `build(fase5)` MSIX manifest + Build-Msix.ps1 + Generate-Assets.ps1 + appinstaller template + release CI job (`tag v*`) con firma opcional vía secrets
- `feat(plugins)` Settings UI enable/disable plugins (`UserPreferences.DisabledPluginAssemblies` + PluginLoader honra lista + UI con CheckBox por plugin)
- `feat(logging)` LogLevel configurable vía Settings con `Serilog.LoggingLevelSwitch` (aplica al instante)
- `feat(audit)` MetricsAggregator (pura) + summary cards en AuditLogView (totales, error rate, last 24h, top sources, errores recientes)
- `feat(ui)` Nav gating Graph/Exchange (NavigationItem.RequiresGraph/RequiresExchange + reactivo al ConnectionStateMonitor)
- `feat(ui)` Botón Tema en sidebar (toggle Dark/Light al instante)
- `feat(cert)` Export PFX con password desde CertWizard (PasswordBox + ICertificateGenerator.ExportPfx)
- `feat(mailflow)` Nuevo módulo "Mail flow" — viewer Get-TransportRule de EXO

Plantamiento status: Fase 4 **DONE** · Fase 5 **MSIX SCAFFOLD DONE** · Fase 6 **IN PROGRESS** (3/8 hechos)

---

### 2026-05-16 / 2026-05-17 — Sesión grande (17 commits)
Tests 70 → 145 (+75). Módulos UI 10 → 13. Plantamiento Fases 1-3 cerradas; Fase 4 foundation; Fase 5-6 in-progress.

Features:
- License assignment UI con SKU picker (single + bulk via `assign:<SkuPartNumber>`)
- Bulk M365 groups + DL via EXO desde CSV (forward-fill GroupName)
- Onboarding wizard (crear user + SKUs + grupos)
- Reglas buzón — Out-of-Office + Forwarding + Permisos calendario
- Audit extendido — guests en grupos M365 privados
- Toasts wpf-ui Snackbar (Ok/Warn/Error)
- TenantHealth — barras de progreso por SKU + agregado
- Plugin system foundation — `IModule` + `PluginLoader` (AssemblyLoadContext)
- CI workflow GitHub Actions + `PublishSingleFile` profile + `PACKAGING.md`
- Audit trail JSONL persistente (`FileAuditLog`) + viewer UI

Docs: Plantamiento añadido como North Star del repo. `deep-research-report.md` restaurado.

> Nota stack: el plantamiento sugiere WinUI 3 como preferente y WPF como fallback aceptable. Se eligió **WPF + wpf-ui** por madurez, ecosistema y compatibilidad con Win10/11. Migración a WinUI 3 queda como posible Fase 7 si surge necesidad.

---

## Estado por fase (Plantamiento §7)

### Fase 1 — Refactor backend + scaffolding plataforma — **DONE**
- [x] Solución .NET 10 con 4 proyectos: `Grex365.Core` (lib), `Grex365.App` (WPF), `Grex365.PowerShell` (helpers), `Grex365.Core.Tests`
- [x] Inyección de dependencias con `Microsoft.Extensions.Hosting`
- [x] MVVM esqueleto con CommunityToolkit.Mvvm (ObservableProperty, RelayCommand)
- [x] Legacy config importer (`UserPreferences`, `CertConfig` desde `GREX365/config/*.json`)
- [x] Modelos de dominio (`UserSummary`, `GroupSummary`, `TenantHealth`, etc.)

### Fase 2 — Motor PowerShell + ejecución asincrónica — **DONE**
- [x] `IPowerShellRunner` con runspace pool (1..5)
- [x] Streams Output/Error/Warning/Verbose redirigidos a `IProgress<LogEntry>`
- [x] Cancellation tokens en cada operación larga
- [x] Serilog → archivo rotativo `%LOCALAPPDATA%\Grex365\logs\` (30 días)
- [x] Global exception handlers (UI dispatcher + AppDomain + TaskScheduler)
- [x] Conexión Graph cert-based (`IGraphConnection`)
- [x] Conexión EXO cert-based (`IExchangeConnection`)
- [x] `IConnectionStateMonitor` con polling 2s + INotifyPropertyChanged
- [x] `ICertValidator` (existencia + validez del thumbprint en CurrentUser\My)
- [x] `ITenantLock` (bloquea conexión si tenant ID no coincide)
- [x] Disconnect Graph + EXO + botón global "Desconectar todo"

### Fase 3 — UI moderna (WPF + Fluent) — **DONE**
Navegación lateral con 14 módulos:
- [x] **Dashboard** — status cards (Graph/EXO/Tenant/Cuenta) + quick actions
- [x] **Conexion** — cert auth Graph + EXO con feedback en vivo
- [x] **Salud tenant** — org + counts usuarios/grupos + SKUs consumidos con **barras de progreso por SKU + total agregado**
- [x] **Usuarios** — buscar, perfil, membresías, enable/disable, quitar licencias, **asignar licencia (SKU picker)**, bulk CSV (`enable`/`disable`/`remove-licenses`/`assign:<SkuPartNumber>`)
- [x] **Grupos** — buscar, miembros, añadir (texto/CSV), eliminar, exportar CSV, **bulk create M365 o DL desde CSV (forward-fill GroupName, toggle M365/DL)**
- [x] **Buzones** — lookup + permisos actuales, Regular↔Shared, FullAccess/SendAs/SendOnBehalf, CSV import/export
- [x] **Reglas buzón** — Out-of-Office (Disabled/Enabled/Scheduled + mensajes interno/externo + rango fechas) · Forwarding (SMTP destino + DeliverToMailboxAndForward) · **Permisos calendario** (Add/Update/Remove via *-MailboxFolderPermission)
- [x] **Auditoria** — identidades (stale members/guests + disabled+licensed) + grupos (sin owner / vacíos / **guests en grupos M365 privados**), paralelizado 8x
- [x] **Onboarding** — wizard compuesto (crear user + UsageLocation + asignar SKUs múltiples + añadir a grupos)
- [x] **Offboarding** — wizard compuesto (deshabilitar + quitar licencias + convertir a shared)
- [x] **Cert Wizard** — generar self-signed RSA 2048, instalar CurrentUser\My, exportar .cer
- [x] **DNS check** — MX/TXT/SPF/DMARC (no requiere auth)
- [x] **Settings (modal)** — tenant lock, cert picker, tema persistido

UX/QoL fase 3:
- [x] Tema Dark/Light persistido en `UserPreferences.Theme`
- [x] Sidebar nav con persistencia del último seleccionado
- [x] Status bar global (Graph/EXO/Tenant/Cuenta + Desconectar)
- [x] Log panel con filtros por severidad + Limpiar
- [x] ProgressRing en operaciones largas
- [x] MessageBox confirm en destructivas (disable, remove licenses, remove member, offboarding, bulk create)
- [x] Cert picker dialog (lista certs CurrentUser\My)
- [x] Toast notifications (wpf-ui `SnackbarPresenter`) en Ok/Warn/Error desde `UiLogSink`

### Fase 4 — Arquitectura modular / plugins — **DONE**
- [x] Contrato `IModule` (Title, Glyph, ViewModelType, ViewType, RegisterServices)
- [x] `PluginLoader` con `AssemblyLoadContext` por DLL desde `%LOCALAPPDATA%\Grex365\plugins\*.dll`
- [x] Discovery con tolerancia a fallos (corruptos/ReflectionTypeLoadException → log warn, no aborta)
- [x] App.xaml.cs: plugins inyectan servicios en DI + registran ViewModels + DataTemplate dinámico
- [x] MainViewModel: append nav entries por cada `IModule` descubierto
- [x] **Sample plugin externo** (`samples/Grex365.SamplePlugin`) — POC compilable y desplegable, con README de empaquetado correcto (no duplica deps del host)
- [x] **Settings UI enable/disable** — `UserPreferences.DisabledPluginAssemblies`, `PluginLoader` honra la lista, panel Plugins en `SettingsWindow` muestra todos los DLLs (cargados / con error / deshabilitados) con CheckBox. Los cambios requieren reinicio.

### Fase 5 — Packaging y despliegue — **MSIX SCAFFOLD DONE**
- [x] PublishSingleFile self-contained para `.exe` portable (`PublishProfiles/win-x64-portable.pubxml`)
- [x] Pipeline CI (GitHub Actions): build + test multiplataforma (.github/workflows/ci.yml)
- [x] Documentación de despliegue (`PACKAGING.md`) con Intune/SCCM/AppInstaller
- [x] **`Package.appxmanifest`** + scaffold completo en `packaging/msix/` (Build-Msix.ps1, Generate-Assets.ps1, assets PNG placeholders)
- [x] **`.appinstaller` plantilla** con `{{FEED_BASE_URI}}` / `{{VERSION}}` para auto-update
- [x] **Job CI `msix`** triggered on tag `v*` (resuelve versión del tag, empaqueta, sube artifact)
- [x] **Hook de firma opcional** en CI (secrets `SIGN_CERT_PFX_B64` + `SIGN_CERT_PASSWORD`; el step se salta si no están configurados)
- [ ] Reemplazar PNG placeholders por arte definitivo de marca
- [ ] Configurar variable `MSIX_FEED_BASE_URI` + secrets de firma en el repo
- [ ] Smoke test de instalación end-to-end con un cert real

### Fase 6 — Telemetría + features enterprise — **IN PROGRESS**
- [x] Audit trail JSONL persistente (`FileAuditLog`) en `%LOCALAPPDATA%\Grex365\audit\audit-YYYY-MM.jsonl`
- [x] `UiLogSink` escribe Ok/Warn/Error a audit con `Environment.UserName` como actor
- [x] Thread-safe via `SemaphoreSlim` y fire-and-forget desde sink
- [x] Audit viewer UI (módulo "Audit log" en la navegación)
- [x] **Niveles de logging DEBUG/INFO/WARN/ERROR configurables vía Settings** (Serilog `LoggingLevelSwitch`, aplica al instante sin reinicio)
- [x] **Métricas agregadas** — `MetricsAggregator` puro (totales por outcome, error rate, last 24h, top sources, errores recientes); AuditLogView muestra panel resumen al cargar mes
- [x] **Application Insights wired** — `ITelemetry`/`NullTelemetry` en Core + `ApplicationInsightsTelemetry` opt-in vía conn string en Settings; `UiLogSink` reenvía Ok/Warn/Error como TrackEvent/TrackException
- [x] **Permisos por rol (RBAC)** — `IRbacGuard` + `IMembershipChecker` (`GraphMembershipChecker` envuelve `/me/checkMemberGroups`). Settings textbox `AuthorizationGroupId`. Offboarding gateado: si no autorizado → status + audit WARN, no ejecuta. Cache invalidado en Disconnect.
- [ ] Documentación técnica interna (arquitectura, manual operación)
- [ ] QA escenarios reales (100+ ops simultáneas)

---

## Backlog funcional (no asociado a una fase concreta)

### Features útiles pendientes
- [x] **Auth interactivo Graph sin cert preexistente** — device-code via cliente público de Azure CLI (`ConnectByDeviceCodeAsync`); valida acceso real con `Me` + `Organization` antes de marcar conectado
- [x] **Mail flow rules viewer** — nuevo modulo de navegacion ("Mail flow") que lista `Get-TransportRule` de EXO (Name/State/Priority/Mode/Description) con filtro libre; gated por RequiresExchange
- [x] **Auditoría: grupos sin actividad reciente** — `RunGroupActivityAuditAsync` consume `/reports/getOffice365GroupsActivityDetail` (period D7/D30/D90/D180), UI con NumberBox umbral, exporta junto al resto de findings
- [x] **Cert export PFX con password** — `ICertificateGenerator.ExportPfx`, panel "Exportar PFX" en CertWizardView con PasswordBox + tests de validacion (no encontrado, password vacio, etc.)
- [x] **Auto-create App Registration vía Graph** — `GraphAppRegistrationService.CreateAndConfigureAsync` aplica todos los AppRoles (User/Group/GroupMember/Organization/AuditLog/Directory + Exchange.ManageAsApp + Reports.Read.All), sube cert como `KeyCredential`, crea ServicePrincipal y devuelve admin-consent URL clickable. Reemplaza los 29 pasos manuales del legacy.
- [x] **Auto-install módulo EXO** — `ExchangeConnection.InstallModuleAsync` lanza `pwsh.exe` externo (Start-Process) para esquivar el ACL de WindowsApps que niega `Microsoft.PackageManagement.dll` en runspaces embebidos. UI muestra estado del módulo + botones Comprobar/Instalar.

### Polish UI
- [ ] Terminal PowerShell embebido (`EasyWindowsTerminalControl`)
- [x] **Theme toggle desde sidebar** — botón "Tema" junto a "Ajustes" persiste y aplica al instante
- [x] **Disable nav items cuando Graph/Exchange desconectado** — `NavigationItem.RequiresGraph/RequiresExchange`, `MainViewModel.UpdateNavEnabledStates` reactivo al `ConnectionStateMonitor`

---

## Tests (174 passing)

| Suite | Tests | Cubre |
|-------|-------|-------|
| LogEntry | 4 | Niveles + factory methods |
| PreferencesStore | 5 | Load/save/defaults |
| PowerShellRunner | 4 | Streams + cancellation |
| CertValidator | 4 | Thumbprint existencia/validez |
| LegacyPreferencesImporter | 3 | Import legacy JSON |
| TenantLock | 5 | Match/mismatch/unset |
| SharedMailboxService | 12 | Apply/convert/permisos/errores |
| FlexibleCsvReader | 8 | Delimitadores, quoted, BOM, edge cases |
| ConnectionStateMonitor | 4 | Estado inicial, polling, fallos, dispose |
| IdentityAuditAnalyzer | 9 | Stale, disabled+lic, totales |
| GroupActivityAnalyzer | 10 | CSV parse (quoted/missing date) + analyze (cutoff strict-less, no-activity, deleted, guard) |
| OffboardingService | 6 | Empty UPN, missing user, per-flag, errores |
| SkuInfo | 6 | Math available, ordering, display, fallback guid |
| BulkGroupRowPreprocessor | 13 | Forward-fill, skip orphans, trim, IsEmail theory |
| OnboardingValidator | 16 | UPN/password/usage/mail-nickname validation + derive |
| MailboxRulesValidator | 15 | OOO state transitions, date ranges, forwarding SMTP shape |
| BulkUserActionParser | 17 | enable/disable/remove-licenses + assign:&lt;SKU&gt; parse + lookup |
| PluginLoader | 4 | empty dir / corrupt dll / whitespace path |
| FileAuditLog | 4 | roundtrip / append-jsonl / missing-month / concurrent-writes |
| MetricsAggregator | 6 | Totales, error rate, last 24h, top sources, recientes |
| CertificateGenerator | 4 | Self-signed + ExportPfx |
| NullTelemetry | 3 | IsEnabled=false + no-throw para TrackEvent/Exception/Flush |
| RbacGuard | 8 | sin-grupo short-circuit, whitespace, miembro/no, checker-throws, cache, Invalidate, trim |

---

## Cómo lanzar

```powershell
dotnet run --project src/Grex365.App/Grex365.App.csproj
```

Primer arranque: 5-15s para JIT. Después abre `MainWindow` (FluentWindow Mica, 1280x800).

Datos persistidos en `%LOCALAPPDATA%\Grex365\`:
- `config/preferences.json` — tenant lock, theme, last nav
- `config/exo-app-params.json` — cert config (AppId, TenantId, Org, Thumbprint)
- `logs/grex365-YYYY-MM-DD.log` — Serilog rotativo (30 días)

---

## Por validar (no automatizable)
- Render visual real en sesión gráfica (no he podido lanzar la UI desde sesión headless)
- Conexión M365 real (necesita tenant + cert reales)
- Comportamiento offline de cada vista (debe mostrar "Graph no está conectado.")
- Comportamiento bulk con CSVs grandes (1k+ filas)

---

## Próximo bloque planificado

**Orden propuesto (mayor utilidad / menor riesgo primero):**
1. **Extender RBAC** al resto de acciones destructivas (Users Disable/RemoveLicense, Groups Delete, SharedMailbox Convert) — patrón ya validado en Offboarding
2. **Documentación técnica interna** — arquitectura + manual operación (ARCHITECTURE.md + RUNBOOK.md)
3. **Asset definitivo MSIX** — reemplazar PNG placeholders por branding + smoke test instalación end-to-end con cert real
4. **QA escenarios reales** — 100+ ops simultáneas bajo carga + scripted bulk CSV grande
5. **Terminal PowerShell embebido** (`EasyWindowsTerminalControl`) — útil para troubleshooting in-app
