# GREX365 v2.0 — Progress log

> **Documento maestro de seguimiento.** Mapea el estado del proyecto contra `Plantamiento_arquitectura_de_la_herramienta.md` (roadmap arquitectónico) y `deep-research-report.md` (research técnico). Toda feature shipped y todo pendiente vive aquí.

- Branch: `grex365-2.0` · Pushed up to `origin/grex365-2.0`
- Stack actual: **C# · .NET 10 · WPF + wpf-ui (Fluent) · MVVM (CommunityToolkit.Mvvm) · Serilog · Microsoft.Extensions.Hosting**
- Tests: **141 passing** (xUnit + FluentAssertions)
- Última actualización: 2026-05-16

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
Navegación lateral con 12 módulos:
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

### Fase 4 — Arquitectura modular / plugins — **FOUNDATION DONE**
- [x] Contrato `IModule` (Title, Glyph, ViewModelType, ViewType, RegisterServices)
- [x] `PluginLoader` con `AssemblyLoadContext` por DLL desde `%LOCALAPPDATA%\Grex365\plugins\*.dll`
- [x] Discovery con tolerancia a fallos (corruptos/ReflectionTypeLoadException → log warn, no aborta)
- [x] App.xaml.cs: plugins inyectan servicios en DI + registran ViewModels + DataTemplate dinámico
- [x] MainViewModel: append nav entries por cada `IModule` descubierto
- [ ] Sample plugin externo de referencia (POC desplegable)
- [ ] Settings UI: enable/disable + reload

### Fase 5 — Packaging y despliegue — **IN PROGRESS**
- [x] PublishSingleFile self-contained para `.exe` portable (`PublishProfiles/win-x64-portable.pubxml`)
- [x] Pipeline CI (GitHub Actions): build + test multiplataforma (.github/workflows/ci.yml)
- [x] Documentación de despliegue (`PACKAGING.md`) con Intune/SCCM/AppInstaller
- [ ] Generar MSIX (single-project) con `Package.appxmanifest`
- [ ] Firma de código (cert EV) en pipeline
- [ ] `.appinstaller` con auto-update apuntando a feed interno
- [ ] Job CI de release (`tag v*`) que firma y publica MSIX

### Fase 6 — Telemetría + features enterprise — **PENDIENTE**
- [ ] Application Insights wired (`Microsoft.ApplicationInsights.WorkerService`)
- [ ] Audit DB local (SQLite con EF Core) para audit trail de acciones admin
- [ ] Métricas: ejecuciones/día, tiempos por operación, errores frecuentes
- [ ] Niveles de logging DEBUG/INFO/WARN/ERROR configurables vía Settings
- [ ] Permisos por rol (validar grupo AD/Entra del usuario actual)
- [ ] Documentación técnica interna (arquitectura, manual operación)
- [ ] QA escenarios reales (100+ ops simultáneas)

---

## Backlog funcional (no asociado a una fase concreta)

### Features útiles pendientes
- [ ] Auth tradicional/UPN interactivo (MSAL) — alternativa al cert-based actual
- [ ] Mail flow rules viewer
- [ ] Auditoría: grupos sin actividad reciente (necesita /reports/getMicrosoft365GroupsActivity)
- [ ] Cert export PFX con password
- [ ] Auto-update App Registration permisos vía Graph (legacy CertWizard hace 29 pasos)

### Polish UI
- [ ] Terminal PowerShell embebido (`EasyWindowsTerminalControl`)
- [ ] Theme toggle accesible desde título / dashboard (ahora solo en Settings)
- [ ] Disable nav items cuando Graph desconectado (gating UX)

---

## Tests (89 passing)

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
| OffboardingService | 6 | Empty UPN, missing user, per-flag, errores |
| SkuInfo | 6 | Math available, ordering, display, fallback guid |
| BulkGroupRowPreprocessor | 13 | Forward-fill, skip orphans, trim, IsEmail theory |
| OnboardingValidator | 16 | UPN/password/usage/mail-nickname validation + derive |
| MailboxRulesValidator | 15 | OOO state transitions, date ranges, forwarding SMTP shape |
| BulkUserActionParser | 17 | enable/disable/remove-licenses + assign:&lt;SKU&gt; parse + lookup |
| PluginLoader | 4 | empty dir / corrupt dll / whitespace path |

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
1. Fase 6 (Application Insights + audit DB SQLite/EF Core)
2. MSIX `Package.appxmanifest` + release CI job (cierra Fase 5)
3. Sample plugin externo (cierra Fase 4)
4. MSAL interactive auth (alternativa a cert-based)
5. Auto-update App Registration permisos via Graph (legacy CertWizard hace 29 pasos)
