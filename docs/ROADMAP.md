# GREX365 — Roadmap & status

> Comprehensive punch list. Updated on every meaningful change.
> Date created: 2026-05-15. Last update: 2026-05-26.
>
> **Live source of truth: [`PROGRESS.md`](../PROGRESS.md)** (sessions bitácora + 901 tests detail).
> This file tracks H0–H6 hitos. PROGRESS.md tracks per-session deltas.
>
> Legend: ✅ done · 🟡 in progress · 🔴 not started · ⚪ deferred · ❌ rejected
>
> Effort scale (solo autodidacta, part-time):
> - S = small (< 1 session)
> - M = medium (1-3 sessions)
> - L = large (3-10 sessions)
> - XL = extra-large (10+ sessions)

---

## 📍 Status snapshot (2026-05-26)

- **Build**: 7 projects (Core + PS + App + Tests Core + Tests App + SamplePlugin + Tests transitive), 0 errors, 0 warnings (WFO0003 suppressed — see PROGRESS Sprint AA)
- **Tests**: 1224 passing (xUnit + FluentAssertions + Moq) — 480 Core + 744 App
- **Nav modules en app**: 15 (Dashboard · Conexión · Licencias · Usuarios · Grupos · Onboarding · Offboarding · Buzones · Reglas de buzón · Flujo de correo · Auditoría · Registro de auditoría · Consola PS · Asistente cert · Comprobación DNS) + Plugins category dinámico
- **Fase 1** (refactor backend + scaffolding) — ✅ DONE
- **Fase 2** (PS engine + async) — ✅ DONE
- **Fase 3** (UI moderna WPF + Fluent) — ✅ DONE
- **Fase 4** (arquitectura modular / plugins) — ✅ DONE
- **Fase 5** (packaging) — 🟡 MSIX scaffold + CI release DONE; real assets (arte) + smoke test (cert) pending external input
- **Fase 6** (telemetría + enterprise) — 🟡 7/8: JSONL audit + viewer + level switch + metrics + AppInsights + RBAC + docs internas (ARCHITECTURE refresh + RUNBOOK nuevo) DONE; falta solo QA escenarios reales (necesita tenant)

## Realistic total effort

Initial estimate "3-5 weeks" assumed full-time dev. Solo + part-time autodidacta reality: **3-4 months** to reach a usable v1.0 that fully replaces the PS toolkit. Below is the full breakdown.

Approximate effort summary:
- H0 (cimientos): S ✅
- H1 (backend core): L ✅
- H2 (Connect feature complete): M ✅
- H3 (port 8-10 features): XL ✅
- H4 (UX polish): M 🟡 mostly done
- H5 (release v1.0): M 🟡 MSIX scaffold done
- H6 (iteration): ongoing

---

## H0 — Cimientos ✅

| # | Item | Status | Effort | Notes |
|---|------|--------|--------|-------|
| 0.1 | `docs/ARCHITECTURE.md` written | ✅ | S | Full stack + decisions |
| 0.2 | `docs/MIGRATION.md` written | ✅ | S | Per-feature status table |
| 0.3 | `docs/ROADMAP.md` (this file) | ✅ | S | Punch list |
| 0.4 | .NET 10 solution scaffolded (4 projects) | ✅ | S | `src/Grex365.slnx` |
| 0.5 | Project references wired | ✅ | S | App → Core + PS; Tests → Core + PS |
| 0.6 | NuGet packages installed | ✅ | S | Graph, Azure.Identity, Serilog, WPF-UI, CommunityToolkit.Mvvm, xUnit |
| 0.7 | `.gitignore` extended (.NET, signing) | ✅ | S | |
| 0.8 | GitHub Actions CI workflow | ✅ | S | `.github/workflows/ci.yml` — build + test + publish on tag |
| 0.9 | First green build (Release + Debug) | ✅ | S | 0 errors, 0 warnings |
| 0.10 | First green tests | ✅ | S | 5/5 passing |
| 0.11 | README updated with migration status | ✅ | S | Links to ARCHITECTURE + MIGRATION |
| 0.12 | Initial commit + branch strategy decided | 🔴 | S | Need to decide: trunk-based vs feature branches |

---

## H1 — Backend core 🟡

Goal: rock-solid services usable from any UI, fully tested.

### 1.1 PowerShellRunner

| # | Item | Status | Effort | Notes |
|---|------|--------|--------|-------|
| 1.1.1 | `IPowerShellRunner` interface | ✅ | S | |
| 1.1.2 | `RunspacePoolHost` (MTA, InitialSessionState) | ✅ | S | |
| 1.1.3 | `PowerShellRunner.RunAsync` happy path | ✅ | S | BeginInvoke + Task.Factory.FromAsync |
| 1.1.4 | Stream forwarding (Info/Warn/Error/Verbose/Debug) | ✅ | S | |
| 1.1.5 | CancellationToken → ps.Stop() | ✅ | S | |
| 1.1.6 | Unit tests for happy path | ✅ | S | PowerShellRunnerTests in test project |
| 1.1.7 | Unit tests for cancellation | ✅ | S | Start-Sleep 30s + 200ms cancel |
| 1.1.8 | Unit tests for stream forwarding | ✅ | S | Info + Warning streams covered |
| 1.1.9 | Integration test: real `Get-Date` script | ✅ | S | Plus concurrent calls test |
| 1.1.10 | Handle PSGallery first-time install (timeout, fallback) | ⚪ | M | Sidestepped: `InstallModuleAsync` launches external `pwsh.exe` (Start-Process) to avoid WindowsApps ACL on embedded runspace |
| 1.1.11 | Progress events for module install | ⚪ | M | Same — external process surfaces progress through stdout |
| 1.1.12 | Reset runspace state on error (avoid contaminated reuse) | 🟡 | M | Mitigated via `DISABLE_REST_API_USE_BY_DEFAULT=true` + cross-runspace EXO probe dropped. Full reset still pending. |

### 1.2 GraphConnection (native SDK)

| # | Item | Status | Effort | Notes |
|---|------|--------|--------|-------|
| 1.2.1 | `IGraphConnection` interface | ✅ | S | |
| 1.2.2 | `ClientCertificateCredential` flow | ✅ | S | |
| 1.2.3 | Smoke test via `Organization.GetAsync` | ✅ | S | |
| 1.2.4 | Cert loading from CurrentUser\My | ✅ | S | |
| 1.2.5 | Real connection state (token validity, not just IsConnected flag) | ✅ | M | `CheckLiveAsync` probes Graph with 10s cache |
| 1.2.6 | Tenant lock enforcement (compare actual TenantId vs expected) | ✅ | S | `ITenantLock.EnforceAsync` in `ConnectViewModel` post-Graph-connect (cert + device-code) |
| 1.2.7 | Scope handling (currently hardcoded `.default`) | ✅ | S | App-only `.default` for cert; explicit scopes for device-code |
| 1.2.8 | Device-code / traditional flow | ✅ | M | Azure CLI public client, no AppReg required |
| 1.2.9 | Connection state cache (avoid re-auth per call) | ✅ | S | `CheckLiveAsync` 10s cache TTL |
| 1.2.10 | Unit tests with Moq | 🔴 | S | Pending — `GraphConnection` integration paths |

### 1.3 ExchangeConnection (runspace)

| # | Item | Status | Effort | Notes |
|---|------|--------|--------|-------|
| 1.3.1 | `IExchangeConnection` interface | ✅ | S | |
| 1.3.2 | `Connect-ExchangeOnline` cert flow | ✅ | S | |
| 1.3.3 | Module ensure (install + import) | ✅ | S | |
| 1.3.4 | Disconnect | ✅ | S | |
| 1.3.5 | Persistent session across runspace pool | 🟡 | M | Connection persists per-runspace via `DISABLE_REST_API_USE_BY_DEFAULT=true`; pool-wide handoff not exhaustively tested |
| 1.3.6 | Real `Test-ExchangeOnlineConnected` via runspace | ✅ | S | `CheckLiveAsync` calls `Get-ConnectionInformation` |
| 1.3.7 | Tenant lock enforcement | ✅ | S | Enforced in `ConnectViewModel` (Graph-side tenant matches; EXO inherits same tenant) |
| 1.3.8 | Device-code flow | 🔴 | M | EXO cert-only; device-code for EXO not added (no real ops driver yet) |
| 1.3.9 | Integration test (mock or real tenant) | 🔴 | M | Manual only |

### 1.4 ConnectionStateMonitor

| # | Item | Status | Effort | Notes |
|---|------|--------|--------|-------|
| 1.4.1 | `IConnectionStateMonitor` interface | ✅ | S | |
| 1.4.2 | 1s poll loop with cancellation | ✅ | S | |
| 1.4.3 | `INotifyPropertyChanged` plumbing | ✅ | S | |
| 1.4.4 | Real check vs `IGraphConnection.IsConnected` | ✅ | S | Calls `CheckLiveAsync` per tick |
| 1.4.5 | Real check vs `Get-ConnectionInformation` runspace | ✅ | M | 10s cache TTL prevents stampede |
| 1.4.6 | Surface tenant + account info in state | ✅ | S | TenantId, Organization, Account flow to UI |
| 1.4.7 | Unit tests | ✅ | S | `ConnectionStateMonitorTests` (4 tests) |

### 1.5 Preferences + cert config

| # | Item | Status | Effort | Notes |
|---|------|--------|--------|-------|
| 1.5.1 | `IPreferencesStore` + JSON impl | ✅ | S | |
| 1.5.2 | `ICertConfigStore` + JSON impl | ✅ | S | |
| 1.5.3 | Unit tests roundtrip | ✅ | S | |
| 1.5.4 | Schema version + migration logic | 🔴 | S | Future-proof if shape changes |
| 1.5.5 | Validation on load (corrupted file → default + warn) | ✅ | S | `JsonPreferencesStore` + `JsonCertConfigStore` catch `JsonException` + quarantine corrupt file to `*.corrupted-yyyyMMddHHmmss.bak` + return defaults (Sprint AH) |
| 1.5.6 | Read legacy paths if found (`GREX365/config/*.json`) | ✅ | S | `LegacyPreferencesImporter` invoked on App startup |

### 1.6 Logging

| # | Item | Status | Effort | Notes |
|---|------|--------|--------|-------|
| 1.6.1 | Serilog rolling file config | ✅ | S | |
| 1.6.2 | `UiLogSink` ObservableCollection | ✅ | S | |
| 1.6.3 | Audit log separate file (who did what when) | ✅ | M | `FileAuditLog` JSONL en `%LOCALAPPDATA%\Grex365\audit\audit-YYYY-MM.jsonl` (thread-safe + AuditLogView viewer) |
| 1.6.4 | Log severity filter in UI | ✅ | S | Checkboxes Info/Ok/Warn/Err/Dbg en log panel (Dbg oculto por defecto) |
| 1.6.5 | Log export (copy/save to file) | ✅ | S | `MainViewModel.ExportLogAsync` + sidebar "Exportar" button — .txt / .csv via `SaveFileDialog`, filtered entries only |

---

## H2 — Connect feature complete 🟡

Goal: the bug that started this conversation is fully fixed in the new app.

| # | Item | Status | Effort | Notes |
|---|------|--------|--------|-------|
| 2.1 | WPF shell with Fluent theme | ✅ | S | WPF-UI Mica window |
| 2.2 | `ConnectViewModel` with Connect/Cancel/Disconnect commands | ✅ | S | |
| 2.3 | Live status dots bound via `INotifyPropertyChanged` | ✅ | S | |
| 2.4 | UI log panel virtualized | ✅ | S | |
| 2.5 | DI bootstrap in `App.xaml.cs` | ✅ | S | |
| 2.6 | Global exception handler | ✅ | S | DispatcherUnhandledException only — need 3 more |
| 2.7 | App.UnhandledException + TaskScheduler.UnobservedTaskException | ✅ | S | Wired in `WireGlobalExceptionHandlers` |
| 2.8 | Settings view (cert path, tenant id, connection method) | ✅ | M | `SettingsWindow` con tabs (cert picker, tenant lock, theme, telemetry, plugins, log level, RBAC group) |
| 2.9 | First-run wizard (no config exists) | 🟡 | M | Auto-connect on startup if cert válido (`TryAutoConnectAsync`). Sin first-run wizard explícito — usuario va a Conexión/Asistente cert manualmente |
| 2.10 | Smoke test against real tenant | 🔴 | M | Manual; documents the flow |
| 2.11 | Replace fake `IsConnected` with real state | ✅ | M | Wired to H1.2.5 + H1.3.6 + dispatcher marshalling |
| 2.12 | Cert config validation UI (warn if cert expired, missing in store) | ✅ | S | `ICertValidator` runs before Connect, blocks if invalid |
| 2.13 | Theme toggle (light/dark) | ✅ | S | Botón Tema en sidebar + persistencia `UserPreferences.Theme` |
| 2.14 | Window restore (size, position) on relaunch | ✅ | S | `MainWindow.RestoreWindowState`/`SaveWindowState` + `WindowPlacementGuard` (off-screen guard for RDP/multi-monitor disconnects) — persisted in `UserPreferences` |
| 2.15 | App icon + branding | 🟡 | S | MSIX assets placeholder; in-app sin branding final |
| 2.16 | About dialog (version, repo link) | ✅ | S | `AboutWindow.xaml` (F1 shortcut) — version + .NET runtime + data dir + open-folder button |

---

## H3 — Migrate features ✅

All major feature ports complete. See [`MIGRATION.md`](MIGRATION.md) for per-feature legacy→new table.

| # | Feature | Status | Service | View |
|---|---|---|---|---|
| 3.1 | Tenant health | ✅ | `TenantHealthService` + `SkuCatalog` | `TenantHealthView` (M365-portal-style cards) |
| 3.2 | Identity audit | ✅ | `GraphAuditService.RunIdentityAuditAsync` + `IdentityAuditAnalyzer` | `AuditView` (compartido) |
| 3.3 | Groups workflow | ✅ | `GraphGroupsService` + `DistributionListsService` + `BulkGroupRowPreprocessor` | `GroupsView` (search + bulk CSV auto-detect type) |
| 3.4 | Mailbox permissions | ✅ | `SharedMailboxService` | `SharedMailboxView` |
| 3.5 | Offboarding wizard | ✅ | `OffboardingService` (composes Users + SharedMailbox + RBAC gate) | `OffboardingView` |
| 3.6 | Cert wizard | ✅ | `CertificateGenerator` + `GraphAppRegistrationService` (auto AppReg + cert upload + consent URL) | `CertWizardView` |
| 3.7 | Roles + UI modes | ⚪ | Superseded by RBAC guard (membership-based) | — |
| 3.8 | Templates | ⚪ | Open — no demand yet | — |
| 3.9 | Reports | 🔴 | Open decision (D3): CSV/XLSX/HTML/in-app? | — |

### Extras (no legacy, ported new)

| Feature | Status | Notes |
|---|---|---|
| Onboarding wizard | ✅ | Create user + assign SKUs + add to groups |
| Mailbox rules (OOO/forwarding/calendar perms) | ✅ | EXO via `MailboxRulesService` |
| Mail flow rules viewer | ✅ | `Get-TransportRule` lister |
| DNS check (MX/TXT/SPF/DMARC) | ✅ | `DomainChecker` |
| 13 security audits (Identity / MFA / CA policy / Privileged roles / App creds / Tenant defaults / OAuth grants / Forwarding ext / Inbox rules / Transport rules / Shared mailbox sign-in / Group activity / Groups hygiene) | ✅ | Cubre baseline M365 security |
| User details drawer (mini-portal lateral M365 admin) | ✅ | Identity card + memberships + license assign/remove + reset password + revoke sessions + slide-in animation |
| Audit log JSONL + viewer with metrics | ✅ | `FileAuditLog` + `MetricsAggregator` + `AuditLogView` |

---

## H4 — UX polish 🟡

| # | Item | Status | Effort | Notes |
|---|------|--------|--------|-------|
| 4.1 | Dashboard home screen | ✅ | M | Cards: connection state + last audit + quick actions |
| 4.2 | Theme toggle (light/dark/auto from system) | ✅ | S | Botón Tema sidebar + persistencia + Auto follows OS theme via WindowsRegistryThemeProvider + SystemEvents.UserPreferenceChanged live update (Sprint H) |
| 4.3 | Sidebar navigation (NavigationView with Frame) | ✅ | M | CIPP-style grouping (Tenant/Identidad/Mail/Seguridad/Herramientas/Plugins) + accent rail PowerToys-style |
| 4.4 | Keyboard shortcuts (Ctrl+, settings, Ctrl+L logs, etc.) | ✅ | S | Audit Ctrl+R/Esc/Ctrl+E, Enter en search Users/Groups/SharedMailbox/MailboxRules/DNS, Esc cierra drawer, Ctrl+, → Settings, F1 → About, Ctrl+L toggle log panel, Ctrl+Enter run / Esc cancel en Consola PS |
| 4.5 | Empty states for every list / panel | ✅ | S | Users/Groups/Audit con placeholder + glyph + texto guía |
| 4.6 | Notification toast for completion + errors | ✅ | M | `SnackbarPresenter` wpf-ui en Ok/Warn/Error |
| 4.7 | Confirmation dialogs for destructive ops | ✅ | S | Disable user, remove licenses, remove member, offboarding, bulk create — via `IDialogService` abstraction |
| 4.8 | Per-monitor DPI testing | 🔴 | S | Pending — manifest declara `<dpiAware>True/PM</dpiAware>` pero no testeado real |
| 4.9 | Accessibility pass (keyboard nav, screen reader) | 🔴 | M | Pending — glyphs en columnas refuerzan color (audit) pero no auditado |
| 4.10 | App icon + splash | 🟡 | S | MSIX assets placeholder shipped; arte definitivo pendiente |
| 4.11 | Spanish/English locale toggle | ✅ | M | D6 cerrado (Sprints AA–AF). L10n.cs static dicts (es+en) + `L10nExtension` MarkupExtension. **20/20 superficies** migradas. 3 ConverterParameter strings hardcoded ES deferred (BoolToOnOffConverter — requiere rewrite) |
| 4.12 | Focus-ring accent en inputs (TextBox/PasswordBox/ComboBox) | ✅ | S | App.xaml global Styles con `IsKeyboardFocused`/`IsKeyboardFocusWithin` → `BrandAccentSolid` border + `AccentGlowSoftEffect` |
| 4.13 | Page transition animations | ✅ | S | ContentControl ControlTemplate fade-in (Sprint F) + Sprint M restraint pass (120ms fade only) |
| 4.14 | Card hover + elevation | ✅ | S | CardHover + MetricCard + HeroCard con storyboards (Sprint L) |
| 4.15 | Sprint L visual overhaul futurista + Sprint M Sally restraint | ✅ | L | Brand palette `#4F8CFF → #6A6CFF → #9B6CFF` + cyan `#7AB7FF` + glass borders + glow effects, 14 views con hero badges → eyebrow+título Light pattern post-restraint |

---

## H5 — Release v1.0 🟡

| # | Item | Status | Effort | Notes |
|---|------|--------|--------|-------|
| 5.1 | `dotnet publish` single-file `.exe` works | ✅ | S | `PublishProfiles/win-x64-portable.pubxml` + `PACKAGING.md` |
| 5.2 | Self-signed code-signing cert | 🔴 | S | Decision pending (D7) |
| 5.3 | `signtool` integrated into publish step | 🟡 | M | Hook opcional en CI (`SIGN_CERT_PFX_B64` + `SIGN_CERT_PASSWORD` secrets) — sin cert real configurado |
| 5.4 | Velopack auto-update | 🔴 | M | Sustituido por MSIX `.appinstaller` template con auto-update |
| 5.5 | GitHub Releases automated on tag | ✅ | S | CI workflow job `msix` triggered on `v*` tag |
| 5.6 | Release notes template | 🔴 | S | Pending |
| 5.7 | Versioning scheme (SemVer + Directory.Build.props) | ✅ | S | `Directory.Build.props` raíz con `<Version>0.2.0-alpha</Version>` single-source. `App.AppVersion` static lee `AssemblyInformationalVersionAttribute`. Sidebar header + status bar bind via `{x:Static local:App.AppVersion}` |
| 5.8 | Install / uninstall docs | ✅ | S | `PACKAGING.md` (Intune/SCCM/AppInstaller) |
| 5.9 | Real (purchased) code-signing cert decision | 🔴 | S | Open (D7) |
| 5.10 | Smoke test on clean Win10/Win11 VMs | 🔴 | M | Pending |
| 5.11 | MSIX `Package.appxmanifest` + scaffold | ✅ | M | `packaging/msix/` + `Build-Msix.ps1` + `Generate-Assets.ps1` |
| 5.12 | Real branded MSIX assets (replace PNG placeholders) | 🔴 | S | Pending |

---

## H6 — Post v1.0 / iteration ✅ (Fase 6 mapping)

| # | Item | Status | Effort | Notes |
|---|------|--------|--------|-------|
| 6.1 | Plugin system (`IModule` + PluginLoader + AssemblyLoadContext) | ✅ | M | Fase 4 shipped. Sample plugin POC + Settings UI enable/disable per DLL |
| 6.2 | MSIX packaging + AppInstaller auto-update | ✅ | L | Fase 5 scaffold shipped. `packaging/msix/` + CI release job `v*` tag. Assets + smoke test pending |
| 6.3 | Application Insights opt-in telemetry | ✅ | M | `ITelemetry` + `NullTelemetry` + `ApplicationInsightsTelemetry`. Empty conn string = NullTelemetry |
| 6.4 | Logging level configurable runtime | ✅ | S | `LoggingLevelSwitch` Serilog desde Settings (Debug/Info/Warn/Error) |
| 6.5 | Audit JSONL log + viewer + métricas | ✅ | M | `FileAuditLog` append-only + `MetricsAggregator` + AuditLogView con filtros + recent errors |
| 6.6 | RBAC vía Entra group membership | ✅ | M | `RbacGuard` cachea decisión, gating en VMs destructivos (Users/Groups/SharedMailbox/MailboxRules). App-only auth bypassea by design |
| 6.7 | Docs internas (ARCHITECTURE + RUNBOOK + ROADMAP + MIGRATION) | ✅ | M | Sprint P 2026-05-23. ARCHITECTURE refresh contra estado real + RUNBOOK manual operacional nuevo |
| 6.8 | QA escenarios reales (100+ ops simultáneas + bulk CSV grande) | 🔴 | M | Blocked: necesita tenant real con datos representativos |
| 6.9 | Multi-tenant support | ❌ | XL | Out of scope v1 (D-multitenant). Multi-profile Windows como workaround |
| 6.10 | Cross-platform (Avalonia) | ❌ | XL | Out of scope, Windows-only target |

---

## Open decisions tracker

| # | Decision | Blocks | Status |
|---|----------|--------|--------|
| D1 | Branching strategy (trunk vs feature branches) | H0.12 | 🟡 De-facto trunk-based en `grex365-2.0`; sin doc explícito |
| D2 | Offboarding: transactional or checkpoint? | H3.5 | 🟡 De-facto step-by-step fail-soft (each step try/catch) |
| D3 | Report format: CSV / XLSX / HTML / DataGrid only? | H3.9 | ✅ Closed — CSV + HTML + JSON shipped (Sprint J-K). XLSX descartado (sin Excel dependency); DataGrid in-app ya existe via AuditView |
| D4 | Roles + UI modes: keep or drop for v1? | H3.7 | ✅ Dropped — RBAC guard cubre |
| D5 | Templates: keep or drop for v1? | H3.8 | 🔴 Open — no demand observado, deferred a iteración post-v1.0 |
| D6 | i18n: Spanish only, or Spanish + English? | H4.11 | ✅ Closed — Spanish + English shipped. Switcher en Settings (restart-required hot-swap). 20/20 surfaces migradas. 3 ConverterParameter strings deferred a sprint dedicado |
| D7 | Code-signing cert: self-signed forever or buy real? | H5.9 | 🔴 Open — decision por user (cost OV/EV cert $100-300/yr vs SmartScreen friction interno) |
| D8 | Min target OS: Win10 1809+ or Win11 only? | many | ✅ Closed — Win10 1809+ de-facto via `.NET 10` runtime + app.manifest supportedOS GUIDs. Win11 inherita Win10 GUID en manifest |
| D9 | Tenant lock: keep legacy preference? | H1.2.6 | ✅ Kept + enforced post-auth (cert + device-code) |

---

## What to work on next (priorizado por valor / riesgo)

**Backlog activo** (per `PROGRESS.md` "Próximo bloque planificado"):

1. **MSIX assets reales** (H5.12 + H4.10) — reemplazar PNG placeholders por branding (necesita arte definitivo de Andersen) + smoke test instalación end-to-end con cert real (H5.10) — blocked external
2. **QA escenarios reales** (H6.8) — 100+ ops simultáneas bajo carga + scripted bulk CSV grande — blocked: necesita tenant real
3. **Decisiones abiertas pendientes** (D3/D5/D6/D7/D8) — Report format, Templates, i18n, Code-signing cert real, Min target OS
4. **A11y + DPI** (H4.8/H4.9) — auditoría accesibilidad screen reader + DPI per-monitor real
5. **i18n** (H4.11) — switcher ES/EN si la decisión D6 se resuelve a favor de bilingüe

**Items shipped en Sprint P 2026-05-23 (sesión noche autónoma)** — ver PROGRESS bitácora:
- ARCHITECTURE.md refresh + RUNBOOK.md nuevo (cierra Fase 6 docs internas)
- ComboBox focus-ring accent (H4.12)
- `chore(health)`: `.Result` purge + nav rename map ampliado
- 5 nuevos test suites: NavTitleMigrator (24) + LicenseFilterMatcher (11) + LicenseCard (13) + InMemoryAuditFindingsStore (5) + UserDetailsHost (10) = +63 tests, total 532
