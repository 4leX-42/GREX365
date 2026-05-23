# GREX365 v2.0 — Architecture

> Living document. Update on every architectural decision.
> Last refresh: 2026-05-23.

---

## 1. Goals

GREX365 is a Microsoft 365 administration toolkit for a single sysadmin operator, built as a desktop Windows application. Long-term goals:

- **Stability**: never freeze; cancellable operations; resilient to transient Graph/EXO failures (Graph eventual-consistency retry already in groups path).
- **Scalability of features**: add new admin workflows (group ops, offboarding, audits, reports) without touching unrelated code.
- **Maintainability**: clear separation of UI / business logic / external API; testable; logs trace every action.
- **UX**: Fluent-style modern Windows look, dark/light/auto theme, live status indicators, real progress bars, cancel buttons.
- **Security**: RBAC gating on destructive commands, tenant lock, signed binaries, opt-in telemetry, audit log of every privileged action.
- **Distribution**: signed single-file `.exe` (portable) + MSIX (Intune/SCCM) with auto-update via AppInstaller.

---

## 2. Stack

| Layer | Choice | Version | Why |
|---|---|---|---|
| Language | C# | 13 (with .NET 10) | First-class M365 SDK support, async/await, source generators |
| Runtime | .NET | 10 | LTS Nov 2025; latest perf; AOT-ready |
| UI framework | WPF | Built-in | Mature tooling, .NET 10 ships Fluent themes via WPF-UI; lower learning curve than WinUI 3 |
| Fluent styling | WPF-UI | 4.3.0 | Mica, NavigationView, modern controls, Fluent theme dictionaries |
| MVVM | CommunityToolkit.Mvvm | 8.4.2 | Microsoft-official, source-generated `[ObservableProperty]` / `[RelayCommand]` |
| DI | Microsoft.Extensions.DependencyInjection | 10.0.8 | Hosting bootstrap |
| Hosting | Microsoft.Extensions.Hosting | 10.0.8 | App lifetime + DI container |
| Logging | Serilog | 10.0.0 (Extensions.Logging) + 7.0.0 (File sink) | Rolling files (30 days), structured, ObservableLogSink for UI panel |
| Graph API | Microsoft.Graph SDK | 5.x | Native .NET, async |
| Auth | `ClientCertificateCredential` / `DeviceCodeCredential` (Azure.Identity) | 1.x | Token caching, MSAL under the hood |
| Exchange Online | `ExchangeOnlineManagement` PS module via runspace | latest | No native .NET SDK; runspace is canonical |
| PowerShell host | `System.Management.Automation` SDK (PowerShell 7.5) | 7.5.x | Embedded PS 7 runspaces, no `pwsh.exe` spawning |
| Telemetry | Microsoft.ApplicationInsights | 2.23.0 | Opt-in via connection string |
| Tests | xUnit 2.9.3 + FluentAssertions 8.10.0 + Moq 4.20.72 + coverlet 6.0.4 | — | Standard .NET stack |
| Build | GitHub Actions on `windows-latest` | — | Free, integrates with repo |
| Packaging | `dotnet publish` self-contained single-file **+** MSIX | — | Portable `.exe` for ad-hoc; MSIX for Intune/SCCM with AppInstaller auto-update |
| Code signing | self-signed initially, EV cert later | — | Avoid SmartScreen warnings on real deploy |

### Stack decisions NOT taken

| Considered | Rejected because |
|---|---|
| WinUI 3 / Windows App SDK | Tooling immature, frequent breaking changes; WPF + WPF-UI delivers same visual result. |
| Prism | Heavyweight; CommunityToolkit.Mvvm covers our needs. Custom plugin loader (`IModule` + `AssemblyLoadContext`) covers modular loading. |
| Avalonia | No cross-platform need; M365 admins are Windows-only. |
| .NET MAUI | Mobile-first; weak desktop story. |
| Blazor Hybrid / Electron | UI-in-web layer adds complexity, perf cost, breaks native feel. |
| Background Windows Service + gRPC IPC | Over-engineered for single-user desktop. |

### Stack decisions later reversed

| Originally rejected | Re-added because |
|---|---|
| Plugin system (MEF-style) | Phase 4 shipped: `IModule` + `PluginLoader` + `AssemblyLoadContext` + sample plugin POC + Settings UI enable/disable per DLL. |
| Application Insights / telemetry | Phase 6 shipped: opt-in via `ITelemetry`/`NullTelemetry`/`ApplicationInsightsTelemetry`. `UiLogSink` forwards Ok/Warn/Error as TrackEvent/TrackException. Empty conn string = NullTelemetry. |
| MSIX / Intune deployment | Phase 5 shipped scaffold: `packaging/msix/` (Package.appxmanifest + Build-Msix.ps1 + Generate-Assets.ps1 + appinstaller template + CI release job on `v*` tag with optional sign secrets). Real branded assets + end-to-end smoke test still pending. |

---

## 3. Solution layout

```
GREX365-main_2/
├─ src/
│   ├─ Grex365.slnx                          ← solution file (no .sln)
│   ├─ Directory.Build.props                 ← single-source <Version>
│   ├─ Grex365.Core/                         ← business logic, no UI ref
│   │   ├─ Abstractions/                     ← 29 interfaces (I*)
│   │   ├─ Audit/                            ← 12 analyzers + builders + stores
│   │   ├─ Certificates/                     ← cert generator + helpers
│   │   ├─ Connections/                      ← Graph + EXO + AppReg + TenantLock
│   │   ├─ Csv/                              ← FlexibleCsvReader
│   │   ├─ DomainChecks/                     ← DNS/MX/SPF/DKIM
│   │   ├─ Groups/                           ← bulk-group parser + service
│   │   ├─ Health/                           ← SKU catalog + license cards
│   │   ├─ Logging/                          ← LogEntry + LoggingLevelSwitch
│   │   ├─ Mailboxes/                        ← shared mbx + rules + forwarding
│   │   ├─ Models/                           ← 18 POCOs / DTOs
│   │   ├─ Offboarding/                      ← OffboardingService
│   │   ├─ Onboarding/                       ← validator + service
│   │   ├─ Plugins/                          ← IModule + PluginLoader
│   │   ├─ Preferences/                      ← JsonPreferencesStore + legacy importer
│   │   ├─ Security/                         ← RbacGuard
│   │   ├─ Users/                            ← bulk-user parser + service
│   │   └─ Grex365.Core.csproj
│   ├─ Grex365.PowerShell/                   ← embedded PS runspace pool
│   │   ├─ PowerShellRunner.cs
│   │   ├─ RunspacePoolHost.cs
│   │   └─ Grex365.PowerShell.csproj
│   └─ Grex365.App/                          ← WPF UI (depends on Core + PS)
│       ├─ App.xaml + App.xaml.cs            ← DI bootstrap + startup chain
│       ├─ MainWindow.xaml + .xaml.cs
│       ├─ SettingsWindow.xaml + .xaml.cs
│       ├─ FirstRunWizardWindow.xaml + .xaml.cs
│       ├─ AboutWindow.xaml + .xaml.cs
│       ├─ ViewModels/                       ← 20 VMs
│       ├─ Views/                            ← 16 module Views
│       ├─ Converters/                       ← 9 value converters
│       ├─ Services/                         ← WPF-only impls (Dialog, Clipboard, AppInsights, ThemeProvider, UiLogSink, Notifier)
│       └─ Grex365.App.csproj
├─ tests/
│   ├─ Grex365.Core.Tests/                   ← 386 tests, 30+ suites
│   └─ Grex365.App.Tests/                    ← 83 tests, 10 suites
├─ samples/
│   └─ Grex365.SamplePlugin/                 ← POC plugin (Fase 4 reference)
├─ packaging/
│   └─ msix/                                 ← Package.appxmanifest + Build-Msix.ps1 + Generate-Assets.ps1 + Grex365.appinstaller + assets/
├─ docs/
│   ├─ ARCHITECTURE.md                       ← this file
│   ├─ ROADMAP.md                            ← punch list H0-H6
│   ├─ MIGRATION.md                          ← legacy→new per-feature log
│   └─ RUNBOOK.md                            ← operations manual (install, troubleshoot)
├─ Plantamiento_arquitectura_de_la_herramienta.md   ← original 6-phase roadmap
├─ deep-research-report.md                   ← initial tech research
├─ PROGRESS.md                               ← authoritative session log
├─ PACKAGING.md                              ← Fase 5 docs
├─ GREX365/                                  ← legacy PS toolkit (source-of-truth porting)
├─ .github/workflows/ci.yml
└─ README.md
```

### Project dependency graph

```
Grex365.App  →  Grex365.Core
             →  Grex365.PowerShell  →  Grex365.Core

Grex365.Core.Tests  →  Grex365.Core
Grex365.App.Tests   →  Grex365.App   →  Grex365.Core + Grex365.PowerShell
Grex365.SamplePlugin → Grex365.Core (ExcludeAssets=runtime — avoid Core dupe in plugin bin)
```

**Rule**: `Grex365.Core` never references WPF/WinUI assemblies. UI-agnostic. This enables headless testing in `Grex365.Core.Tests` (net10.0).

**Rule**: ViewModels never reference WPF types directly. Use `IDialogService` / `IClipboardService` abstractions in `Grex365.Core/Abstractions`. WPF impls (`WpfDialogService`, `WpfClipboardService`) live in `Grex365.App/Services/`.

---

## 4. Layers and patterns

### 4.1 MVVM

- Views: XAML + minimal code-behind (only WPF-specific wiring like KeyBinding handlers, Loaded hooks for view-driven triggers).
- ViewModels in `Grex365.App/ViewModels/`. Inherit `ObservableObject` + `[ObservableProperty]` / `[RelayCommand]` source generators.
- Models in `Grex365.Core/Models/`. POCOs.
- All page ViewModels are **Singleton** in DI (state persists across tab switches — see Sprint state persistence 2026-05-23). Settings + FirstRun VMs stay Transient (modal one-shot).

### 4.2 Dependency injection

`App.xaml.cs` builds a `Microsoft.Extensions.Hosting.Host` with ~65 services registered. Order:

1. Logger/hosting stack (`SerilogLoggerFactory`, `ILogger<>`, `LoggingLevelSwitch`).
2. PowerShell (`RunspacePoolHost`, `IPowerShellRunner`).
3. Connections (`IGraphConnection` → `GraphConnection`, `IExchangeConnection` → `ExchangeConnection`, etc.).
4. Domain services (`IGroupsService` → `GraphGroupsService`, `IAuditService` → `GraphAuditService`, etc.).
5. UI services (`IDialogService` → `WpfDialogService`, `IClipboardService` → `WpfClipboardService`, `IUiLogSink` → `UiLogSink`, `ITelemetry` → `ApplicationInsightsTelemetry` or `NullTelemetry`).
6. ViewModels (Singleton for page VMs; Transient for `FirstRunWizardViewModel` + `SettingsViewModel`).
7. Windows (`MainWindow`, `SettingsWindow`, `FirstRunWizardWindow`, `AboutWindow`).

### 4.3 Startup chain (`App.OnStartup`)

```
1. Serilog config + LoggingLevelSwitch from preferences
2. PluginLoader.LoadAsync(plugins/ dir) → registers IModule services
3. Build IHost + IServiceProvider
4. Apply theme (resolves Auto via WindowsRegistryThemeProvider)
5. Subscribe SystemEvents.UserPreferenceChanged for live theme follow
6. ShowFirstRunWizardIfNeededAsync()       ← only if FirstRunCompleted=false
7. TryAutoConnectAsync()                    ← Graph cert + EXO if cert + tenant present
8. MainWindow.Show() with restored window pos/size from prefs
```

### 4.4 Async + cancellation

Every long-running method exposes:
- `CancellationToken` parameter
- `IProgress<LogEntry>` for streaming progress
- Returns `Task<TResult>` or `Task`

Cancel buttons in UI bind to `CancellationTokenSource` owned by the ViewModel. On Cancel: `cts.Cancel()` propagates to runspaces (`PowerShell.Stop()`) and Graph (HTTP cancellation).

Search inputs (Users, Groups) use **debounced typeahead** (250 ms) with `CancellationTokenSource` swap to cancel stale Graph calls. Min 2 chars before hitting Graph; snapshot guard discards out-of-order callbacks.

### 4.5 State observation

`IConnectionStateMonitor` polls Graph + EXO every second on a background timer. Raises `PropertyChanged` events when state flips. UI bindings update automatically. `MainViewModel.UpdateNavEnabledStates()` listens and toggles `NavigationItem.IsEnabled` based on `RequiresGraph` / `RequiresExchange` flags.

### 4.6 Logging

```csharp
Log.Logger = new LoggerConfiguration()
    .MinimumLevel.ControlledBy(logLevelSwitch)
    .WriteTo.File("%LOCALAPPDATA%/Grex365/logs/grex365-.log",
                  rollingInterval: RollingInterval.Day,
                  retainedFileCountLimit: 30)
    .WriteTo.Sink(observableLogSink)   // ObservableCollection<LogEntry> bound to UI log panel
    .CreateLogger();
```

`UiLogSink` (in `Grex365.App/Services`) wraps `IUiLogSink` and forwards `Ok` / `Warn` / `Error` log entries into `ITelemetry.TrackEvent` / `TrackException`.

`LoggingLevelSwitch` lets the user change level live from Settings (Debug / Information / Warning / Error).

### 4.7 Error handling

- Methods: try/catch only where adding context. Otherwise let exceptions bubble.
- ViewModel command handlers wrap calls in try/catch, log with Serilog, show toast / message box via `IDialogService`.
- `Application.DispatcherUnhandledException` + `AppDomain.CurrentDomain.UnhandledException` + `TaskScheduler.UnobservedTaskException` all wired to a single global handler that logs full stack trace and prompts.

---

## 5. PowerShell integration

### 5.1 Why PowerShell stays

Some Microsoft 365 surface has **no .NET SDK equivalent**:
- Exchange Online cmdlets (`Get-Mailbox`, `Set-MailboxPermission`, `Get-TransportRule`, etc.) — only available in `ExchangeOnlineManagement` module.
- Some Teams + SharePoint admin cmdlets.

For Microsoft Graph we **prefer the native .NET SDK**, not the PowerShell module. Same API surface, far less overhead, real async.

### 5.2 RunspacePool design

`Grex365.PowerShell.RunspacePoolHost`:
- Single shared `RunspacePool` (min 1, max 4) created on app startup.
- `InitialSessionState`: PSGallery trusted, `ConfirmPreference=None`, `ProgressPreference=SilentlyContinue`, `ErrorActionPreference=Continue`.
- Apartment state **MTA** (avoid MSAL deadlocks).
- Pre-imports `ExchangeOnlineManagement` on first runspace allocation.

`PowerShellRunner.RunScriptAsync(string script, IDictionary<string,object>? args, IProgress<LogEntry>? progress, CancellationToken ct)`:
1. Acquires a `PowerShell` instance from the pool.
2. Subscribes `Streams.Information / Warning / Error / Verbose / Debug / Progress` → `progress`.
3. `BeginInvoke()` async.
4. Awaits completion or cancellation. On cancel: `ps.Stop()`.
5. Returns `Collection<PSObject>` or throws.

### 5.3 EXO connection lifecycle

Cert flow (preferred for unattended sysadmin use):
```powershell
Connect-ExchangeOnline `
  -AppId $AppId `
  -CertificateThumbprint $Thumb `
  -Organization $Org `
  -ShowBanner:$false
```
Connection persists across subsequent runspace allocations (same pool, runspaces reused). Disconnect on app shutdown: `Disconnect-ExchangeOnline -Confirm:$false`.

**Module auto-install** (`ExchangeConnection.InstallModuleAsync`) launches `pwsh.exe` externally via `Start-Process` to bypass the WindowsApps ACL that denies `Microsoft.PackageManagement.dll` inside embedded runspaces. UI shows module state + Comprobar/Instalar buttons.

### 5.4 PS Console module

`PsConsoleViewModel` exposes a multi-line REPL inside the app (Herramientas → Consola PS). Reuses the same `IPowerShellRunner` — scripts execute in app context (Graph/EXO in scope). History navigable Up/Down (max 50, dedupe consecutive). Ctrl+Enter → Run, Esc → Cancel. Closes Plantamiento §6 backlog "Terminal PowerShell embebido" without adding `EasyWindowsTerminalControl` integration.

---

## 6. Connections + Auth

### 6.1 Graph (native .NET SDK)

Two auth paths, both gated by `TenantLock`:

**Cert (unattended)**: `ClientCertificateCredential(tenantId, clientId, X509Certificate2)` — token cached by MSAL.
**Device code (interactive bootstrap)**: `DeviceCodeCredential` with `organizations` tenant — used when no cert exists yet. After login, `_graph.TenantId` checked non-null before marking connected (closes bypass identified in 2026-05-22 audit: `ConnectViewModel.ConnectByDeviceCodeAsync` aborts if null).

### 6.2 App Registration auto-create

`GraphAppRegistrationService.CreateAndConfigureAsync` (refactored via `AppRegistrationSpec` pure helpers, 17 tests):
- Creates app with 9 Graph AppRoles + Exchange.ManageAsApp + Reports.Read.All.
- Uploads cert as `KeyCredential`.
- Creates ServicePrincipal.
- Returns clickable admin-consent URL.

Replaces 29 manual steps from legacy `Certificate-Setup-Steps.csv`.

### 6.3 TenantLock

`ITenantLock` enforces expected TenantId/Domain against actual `_graph.TenantId` after login. Enforced in both auto-connect and manual connect (cert + device-code). Can be skipped via Settings checkbox (`EnforceTenantLock=false`) for dev/lab.

### 6.4 RBAC

`IRbacGuard` checks user membership in a configured Entra group before allowing destructive commands (Users / Groups / SharedMailbox / MailboxRules). Cached per session, `Invalidate()` flushes on disconnect. App-only auth (cert flow) **bypasses RBAC by design** — no `me.CheckMemberGroups` context available.

### 6.5 Graph eventual consistency

`WithGraphReplicaRetryAsync` wraps `POST /groups` followed by `PATCH/POST /groups/{id}/members` ops. Exponential backoff (1.5 → 3 → 6 → 12 → 15 s cap). 8 retries on `justCreated=true`, 2 on existing groups. Catches `Request_ResourceNotFound`, `ResourceNotFound`, message "does not exist".

---

## 7. Audit subsystem

Two parallel audit surfaces — both write to the unified `FileAuditLog` (JSONL).

### 7.1 Privileged-action audit log

Every privileged command (group create, license assign, user disable, mailbox convert, etc.) writes a row to `%LOCALAPPDATA%/Grex365/audit/grex365-YYYY-MM.jsonl`. Format:

```json
{"ts":"2026-05-23T10:31:00Z","actor":"alex@tenant","source":"GroupsService","action":"CreateM365Group","target":"sales","result":"Ok","detail":"…"}
```

`MetricsAggregator` (pure) computes totals, error rate, last-24h count, top sources, recent errors — surfaced in **AuditLog view**.

### 7.2 Security audit reports (read-only analyzers)

Triggered manually from the **Auditoría** view. 12 pure analyzers + `GraphAuditService` (orchestrator):

| # | Analyzer | Detects | Required Graph permission |
|---|---|---|---|
| 1 | `IdentityAuditAnalyzer` | Stale users + disabled-with-license | `User.Read.All` |
| 2 | `GroupActivityAnalyzer` | Groups inactive ≥N days via `/reports/getOffice365GroupsActivityDetail` | `Reports.Read.All` |
| 3 | `MailboxForwardingAnalyzer` (EXO) | External forwarding (vector exfil) | EXO `Get-Mailbox` |
| 4 | `InboxRuleAnalyzer` (EXO) | BEC indicators: delete / hide / external forward + ES/EN keywords | EXO `Get-InboxRule` |
| 5 | `MfaCoverageAnalyzer` | Admin/member/guest without MFA, via `/reports/authenticationMethods/userRegistrationDetails` | `Reports.Read.All` |
| 6 | `CaPolicyAnalyzer` | Conditional Access weak/disabled/report-only-stale policies | `Policy.Read.All` |
| 7 | `PrivilegedRoleAuditAnalyzer` | Guest admins, disabled admins, 0/1 GAs, GA sprawl | `Directory.ReadWrite.All` |
| 8 | `AppCredentialAuditAnalyzer` | Expired / expiring-soon / long-lived secrets + keys | `Application.Read.All` |
| 9 | `TenantDefaultsAnalyzer` | users-can-create-apps, invitesFrom=everyone, SSPR disabled, etc. | `Policy.Read.All` |
| 10 | `OAuthGrantAnalyzer` | AllPrincipals high-risk OAuth grants (admin consent) + user-consented phish-OAuth | `Directory.ReadWrite.All` |
| 11 | `TransportRuleAuditAnalyzer` (EXO) | Forwarding/BCC/redirect to external recipients, custom outbound connectors, broad-scope deletes, disabled-security-keyword rules | EXO `Get-TransportRule` |
| 12 | `SharedMailboxSignInAnalyzer` (EXO) | Shared mailboxes with sign-in enabled (vector password attack) | EXO `Get-Mailbox` + `Get-User` |

All analyzers are **pure** (no Graph/EXO calls) — input is a typed model, output is `IEnumerable<AuditFinding>`. Tested with seed data, no mocks needed.

### 7.3 Report outputs

- **CSV** export (default)
- **HTML** report via `AuditReportHtmlBuilder` — standalone, embedded CSS, light + `prefers-color-scheme: dark`, severity pills, escape HTML entities.
- **JSON** export via `AuditReportJsonBuilder` — schema `grex365.audit.v1`, parseable for automation.
- **Baseline diff** via `AuditBaselineComparer` — load previous JSON, compute New / Resolved / Persistent. Identity = `(Category, Identity, Detail, Severity-case-insensitive)`.

---

## 8. Plugin system

`Grex365.Core/Plugins/`:
- **`IModule`** contract: `string Name`, `Version Version`, `void Register(IServiceCollection services)`, `void RegisterNav(INavRegistrar nav)`.
- **`PluginLoader`** scans `%LOCALAPPDATA%/Grex365/plugins/`, loads each DLL in its own `AssemblyLoadContext`, finds `IModule` types, instantiates, calls `Register`. Failures reported as warnings (corrupt DLL, missing dep, throw in Register) — never block app startup.

`UserPreferences.DisabledPluginAssemblies` (HashSet) — Settings UI toggles per DLL. Loader skips disabled assemblies.

Sample plugin at `samples/Grex365.SamplePlugin/`:
- `<Project Sdk="Microsoft.NET.Sdk">` with `CopyLocalLockFileAssemblies=false`.
- References `Grex365.Core` with `ExcludeAssets=runtime` + MVVM Toolkit with `PrivateAssets=all` → avoids duplicating Core in plugin bin (loaded from host).

CI builds the sample plugin and uploads as artifact (`plugin-sample.dll`).

---

## 9. Telemetry

Opt-in via `appsettings.json` or env var (Application Insights connection string). Empty / null → `NullTelemetry` (no-op).

**`ITelemetry`** contract (`Grex365.Core/Abstractions`):
- `bool IsEnabled`
- `void TrackEvent(string name, IDictionary<string,string>? properties = null)`
- `void TrackException(Exception ex, IDictionary<string,string>? properties = null)`
- `void Flush()`

Impls:
- `NullTelemetry` (`Grex365.Core`) — always returns IsEnabled=false, no-throw on every method.
- `ApplicationInsightsTelemetry` (`Grex365.App.Services`) — wraps `TelemetryClient`. Auto-collects unhandled exceptions via global handler.

`UiLogSink` forwards `Ok` → `TrackEvent("UiLog.Ok")`, `Warn` → `TrackEvent("UiLog.Warn")`, `Error` → `TrackException`. Decouples logging UI panel from telemetry pipeline.

---

## 10. Theme system

Three modes: **Dark** (default), **Light**, **Auto** (follow Windows).

`ISystemThemeProvider` abstraction (`Grex365.Core/Abstractions`) → `WindowsRegistryThemeProvider` (`Grex365.App/Services`) reads:
```
HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme
```
DWORD 0=dark / 1=light. Defaults to dark if registry inaccessible.

`SettingsViewModel.ResolveActualTheme()` maps Auto → provider.IsDarkTheme(). 

`Microsoft.Win32.SystemEvents.UserPreferenceChanged` subscribed in `App.OnStartup`: when user flips Windows theme, re-applies via dispatcher only if pref="Auto" (otherwise respect explicit choice). Unsubscribed in `OnExit`.

Palette unified in `App.xaml`:
- Brand: `BrandAccentStart #4F8CFF` → `BrandAccentMid #6A6CFF` → `BrandAccentEnd #9B6CFF` + `BrandCyan #22D3EE`.
- Semantic: `Color + Brush` resources `SemanticError/Warn/Info/Ok/Neutral/Debug` + Soft variants (alpha 0x55) + MutedText variants. 16 Color + 14 Brush keys.
- 4 converters (`SeverityToBrush`, `AuditSeverityToBrush`, `UtilizationToBrush`, `BoolToBrush`) → `Application.Current.TryFindResource(key)` lookup (dynamic). Theme toggle affects derived colors.
- Hardcoded XAML colors swept to zero (verified `grep "Foreground=\"#\|Background=\"#"` post-sweep).

---

## 11. Modules / navigation (14)

Order in sidebar nav:

| # | Nav title | View | RequiresGraph | RequiresExchange | Notes |
|---|---|---|---|---|---|
| 1 | Dashboard | DashboardView | — | — | Hero badge + quick actions |
| 2 | Conexión | ConnectView | — | — | Manual + auto-connect entry |
| 3 | Licencias | TenantHealthView | ✓ | — | (renamed 2026-05-23 from "Salud tenant" — auto-load + search filter + "Gestionar" deep-link to Usuarios) |
| 4 | Usuarios | UsersView | ✓ | — | Bulk CSV `assign:<SkuPartNumber>` + debounced typeahead |
| 5 | Grupos | GroupsView | ✓ | — | Bulk M365/DL + RadioButtons type choice + Graph replica retry |
| 6 | Buzones compartidos | SharedMailboxView | ✓ | ✓ | Apply/convert/permissions |
| 7 | Reglas de buzón | MailboxRulesView | ✓ | ✓ | OOO + forwarding + calendar permissions |
| 8 | Flujo de correo | MailFlowRulesView | — | ✓ | Get-TransportRule viewer with filter |
| 9 | Auditoría | AuditView | ✓ | ✓ | 12 security analyzers + CSV/HTML/JSON export + baseline diff |
| 10 | Registro de auditoría | AuditLogView | — | — | JSONL viewer + metrics |
| 11 | Onboarding | OnboardingView | ✓ | — | UPN/password/usage validation |
| 12 | Offboarding | OffboardingView | ✓ | ✓ | Disable + license remove + sign-out + convert shared |
| 13 | Asistente cert | CertWizardView | — | — | Self-signed + PFX export with password |
| 14 | Comprobación DNS | DomainCheckView | — | — | MX/SPF/DKIM/DMARC |
| 15 | Consola PS | PsConsoleView | — | — | Multi-line REPL (Herramientas) |

Nav items grey out when their connection is missing (`MainViewModel.UpdateNavEnabledStates` reactive to `IConnectionStateMonitor`).

---

## 12. Build, test, release

### 12.1 Local build + test

```powershell
dotnet build src/Grex365.slnx -c Release
dotnet test  src/Grex365.slnx -c Release --no-build
```

No `.sln` — `Grex365.slnx` is the slim solution file.

### 12.2 CI (`.github/workflows/ci.yml`)

Triggers: push to `main` / `grex365-2.0`, PR to `main`.

Steps:
1. Checkout
2. Setup .NET 10 SDK
3. Restore (Core, PowerShell, App, tests, sample plugin)
4. Build App (transitive cascade)
5. Build SamplePlugin
6. Test Core.Tests + App.Tests with `--logger trx`
7. Upload artifacts: sample plugin DLL + `test-results.trx`

### 12.3 Portable publish

```powershell
dotnet publish src/Grex365.App/Grex365.App.csproj `
  -c Release -r win-x64 --self-contained `
  -p:PublishSingleFile=true `
  -p:IncludeAllContentForSelfExtract=true `
  -p:EnableCompressionInSingleFile=true
```

Output: `bin/Release/net10.0-windows/win-x64/publish/Grex365.exe` (~70–90 MB, no .NET install required).

### 12.4 MSIX

```powershell
.\packaging\msix\Build-Msix.ps1 -Version 0.2.0
```

Generates `Grex365_0.2.0_x64.msix` from `Package.appxmanifest`. Sign with `signtool` (cert from secrets) — CI release job (`v*` tag) automates this when secrets present.

AppInstaller (`Grex365.appinstaller`) points to a feed URI (env var `MSIX_FEED_BASE_URI`) for auto-update.

### 12.5 Code signing

Phase 1 (dev): self-signed cert via `New-SelfSignedCertificate -Type CodeSigningCert`. Test users add to trusted publishers.
Phase 2 (release): OV/EV cert (Sectigo). Sign with `signtool sign /f cert.pfx /p pass /tr http://timestamp.sectigo.com /td sha256 /fd sha256 Grex365.exe`.

---

## 13. Data on disk

```
%LOCALAPPDATA%/Grex365/
  config/
    user_preferences.json        ← tenant lock, theme, last nav, window pos/size, disabled plugins, log level
    exo-app-params.json          ← AppId, TenantId, Org, Cert thumbprint
  logs/
    grex365-YYYY-MM-DD.log       ← Serilog daily rolling (30-day retention)
  audit/
    grex365-YYYY-MM.jsonl        ← privileged-action audit log (append-only)
  plugins/                       ← drop DLLs here for runtime discovery
```

---

## 14. Conventions

### 14.1 Naming
- C# Microsoft conventions: PascalCase types/methods, camelCase locals, `_camelCase` private fields, `I`-prefixed interfaces.
- File per type (one public type per `.cs` file).
- Namespaces match folder layout: `Grex365.Core.Connections`, etc.

### 14.2 Style
- File-scoped namespaces.
- `nullable enable` everywhere.
- Sparing `var` (only when type is obvious from `new()`).
- No regions.
- No comments restating code. Comments only for non-obvious *why*.

### 14.3 Async
- Suffix `Async` on every async method.
- `ConfigureAwait(false)` in `Grex365.Core` and `Grex365.PowerShell` (library code).
- App layer (`Grex365.App`): default context flows back to UI thread.
- Never `.Result` / `.Wait()`. Use `await`.

### 14.4 Logging
- Structured logging: `Log.Information("Connecting to {Service} as {Account}", service, account);`
- Levels: `Verbose` (noise), `Debug` (dev), `Information` (user-relevant), `Warning` (recoverable), `Error` (handled), `Fatal` (crashes).

### 14.5 Testing
- xUnit. One test class per production class.
- FluentAssertions for readable asserts.
- Moq for interface mocking (App.Tests only — Core uses pure analyzers, no mocks).
- App.Tests uses harness pattern: `TestDialogService` / `TestClipboardService` / `TestUiLogSink` / `TestUserDetailsHost` (see `tests/Grex365.App.Tests/TestFakes.cs`).

### 14.6 Commits
Conventional Commits with scope: `feat(scope):`, `fix(scope):`, `docs:`, `refactor:`, `test:`, `build:`, `ci:`, `ux:`. HEREDOC for multi-line bodies. `Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>` footer. Spanish acceptable in messages when adds context.

---

## 15. Open questions / future decisions

- [ ] EV code-signing cert: purchase for general distribution or stay self-signed for internal-only?
- [ ] AppInsights connection string distribution: per-tenant config endpoint vs hardcoded dev/prod?
- [ ] Multi-tenant scenarios: stay single-tenant per profile (current) or add tenant switcher?
- [ ] Plugin sandboxing: current `AssemblyLoadContext` isolates assemblies but plugins still run in-proc with full app perms. Sandbox needed for third-party plugins?
- [ ] Velopack vs MSIX AppInstaller: pick one auto-update mechanism. Currently MSIX path is primary; Velopack documented as alternative but not wired.

---

## 16. References

- Phase punch list + session log: [`../PROGRESS.md`](../PROGRESS.md)
- North-star roadmap: [`../Plantamiento_arquitectura_de_la_herramienta.md`](../Plantamiento_arquitectura_de_la_herramienta.md)
- Original tech research: [`../deep-research-report.md`](../deep-research-report.md)
- Per-feature migration log: [`MIGRATION.md`](MIGRATION.md)
- H0-H6 sub-item tracker: [`ROADMAP.md`](ROADMAP.md)
- MSIX/AppInstaller/Intune docs: [`../PACKAGING.md`](../PACKAGING.md)
- Operations manual: [`RUNBOOK.md`](RUNBOOK.md)
- WPF-UI (Fluent for WPF): https://wpfui.lepo.co/
- CommunityToolkit.Mvvm: https://learn.microsoft.com/dotnet/communitytoolkit/mvvm/
- Microsoft.Graph SDK: https://learn.microsoft.com/graph/sdks/sdks-overview
