# GREX365 — Architecture

> Living document. Update on every architectural decision.
> Author: alexa + Claude (Opus 4.7). Date created: 2026-05-15.

---

## 1. Goals

GREX365 is a Microsoft 365 administration toolkit for a single operator (sysadmin), built as a desktop Windows application. Long-term goals:

- **Stability**: never freeze; cancellable operations; resilient to transient Graph/EXO failures.
- **Scalability of features**: add new admin workflows (group ops, offboarding, audits, reports) without touching unrelated code.
- **Maintainability**: clear separation of UI / business logic / external API; testable; logs trace every action.
- **UX**: Fluent-style modern Windows look, dark/light theme, live status indicators, real progress bars, cancel buttons.
- **Distribution**: signed single-file `.exe` with auto-update.

---

## 2. Stack

| Layer | Choice | Version | Why |
|---|---|---|---|
| Language | C# | 13 (with .NET 10) | First-class M365 SDK support, async/await, source generators |
| Runtime | .NET | 10 (LTS, GA Nov 2025) | LTS until Nov 2028, latest perf, AOT-ready |
| UI framework | WPF | Built-in | Mature tooling (XAML designer, snoop, hot reload), huge community, .NET 10 ships Fluent themes, easier learning curve than WinUI 3 |
| Fluent styling | WPF-UI | latest | Fluent Design ported to WPF (Mica, navigation view, modern controls) |
| MVVM | CommunityToolkit.Mvvm | 8.x | Microsoft-official, source-generated `[ObservableProperty]` / `[RelayCommand]`, minimal boilerplate |
| DI | Microsoft.Extensions.DependencyInjection | 10.x | Standard, integrates with HostBuilder |
| Logging | Serilog | 4.x | Rolling files, structured logs, sinks for console + file |
| Graph API | Microsoft.Graph SDK | 5.x | Native .NET, async, no PS overhead |
| Auth | Azure.Identity (`ClientCertificateCredential`) | 1.x | Token caching, MSAL under the hood |
| Exchange Online | PowerShell `ExchangeOnlineManagement` via runspace | latest | No native .NET SDK for EXO cmdlets exists; runspace is canonical |
| PowerShell host | `System.Management.Automation` (PowerShell SDK NuGet) | 7.5.x | Embedded PS 7 runspaces, no `pwsh.exe` spawning |
| Tests | xUnit + FluentAssertions + Moq | latest | Standard .NET stack |
| Build | GitHub Actions on `windows-latest` | — | Free, integrates with repo |
| Packaging | `dotnet publish` self-contained single-file | — | Simple `.exe`, no MSIX/Intune until needed |
| Auto-update | Velopack | latest | Squirrel successor; simple GitHub Releases as update feed |
| Code signing | self-signed initially, EV cert later | — | Avoid SmartScreen warnings once mature |

### Stack decisions explicitly NOT taken

| Considered | Rejected because |
|---|---|
| WinUI 3 / Windows App SDK | Tooling immature, frequent breaking changes, smaller community, harder for self-taught dev. WPF achieves visually identical results with WPF-UI. |
| Prism | Heavyweight framework; CommunityToolkit.Mvvm covers our needs. |
| Avalonia | We don't need cross-platform; M365 admins are Windows-only. |
| .NET MAUI | Mobile-first; weak desktop story. |
| Blazor Hybrid / Electron | UI-in-web layer adds complexity, performance cost, breaks native feel. |
| Plugin system (Prism modules, MEF) | YAGNI for solo project; revisit at v2. |
| Background Windows Service + gRPC IPC | Over-engineered for single-user desktop app. |
| Application Insights / telemetry | Single operator = no need for remote telemetry. |
| MSIX / Intune deployment | Premature; start with single-file `.exe`. |

---

## 3. Solution layout

```
GREX365-main_2/
├─ src/
│   ├─ Grex365.sln
│   ├─ Grex365.Core/              ← business logic, no UI ref
│   │   ├─ Abstractions/          ← interfaces
│   │   ├─ Connections/           ← Graph + EXO orchestration
│   │   ├─ Models/                ← DTOs, domain types
│   │   ├─ Preferences/           ← user_preferences.json IO
│   │   ├─ Logging/               ← Serilog config
│   │   └─ Grex365.Core.csproj
│   ├─ Grex365.PowerShell/        ← embedded PS runspace pool
│   │   ├─ PowerShellRunner.cs
│   │   ├─ RunspacePoolHost.cs
│   │   └─ Grex365.PowerShell.csproj
│   └─ Grex365.App/                ← WPF UI (depends on Core + PS)
│       ├─ App.xaml
│       ├─ App.xaml.cs            ← DI bootstrap
│       ├─ MainWindow.xaml
│       ├─ ViewModels/
│       ├─ Views/
│       ├─ Converters/
│       ├─ Themes/
│       └─ Grex365.App.csproj
├─ tests/
│   └─ Grex365.Core.Tests/
│       └─ Grex365.Core.Tests.csproj
├─ docs/
│   ├─ ARCHITECTURE.md            ← this file (stack + patterns)
│   ├─ ROADMAP.md                 ← punch list H0–H6, every sub-item
│   └─ MIGRATION.md               ← legacy→new migration log
├─ deep-research-report.md         ← initial research (repo root)
├─ GREX365/                       ← legacy PS toolkit (untouched until ported)
├─ .github/workflows/ci.yml
└─ README.md
```

### Project dependency graph

```
Grex365.App  →  Grex365.Core
             →  Grex365.PowerShell  →  Grex365.Core

Grex365.Core.Tests  →  Grex365.Core
                    →  Grex365.PowerShell  (integration tests)
```

**Rule**: `Grex365.Core` never references WPF/WinUI assemblies. UI-agnostic. This is what enables headless testing.

---

## 4. Layers and patterns

### 4.1 MVVM

- Views are XAML + minimal code-behind (only WPF-specific wiring).
- ViewModels live in `Grex365.App/ViewModels/`. Use `ObservableObject` base + `[ObservableProperty]` / `[RelayCommand]` from CommunityToolkit.Mvvm.
- Models live in `Grex365.Core/Models/`. POCOs.
- ViewModels never reference WPF types directly. Use `IDialogService` / `INotificationService` abstractions when UI interaction is needed.

### 4.2 Dependency injection

`App.xaml.cs` builds a `Microsoft.Extensions.Hosting.Host` with services registered:

```csharp
services.AddSingleton<IPowerShellRunner, RunspacePoolRunner>();
services.AddSingleton<IGraphConnection, GraphConnection>();
services.AddSingleton<IExchangeConnection, ExchangeConnection>();
services.AddSingleton<IConnectionStateMonitor, ConnectionStateMonitor>();
services.AddSingleton<IPreferencesStore, JsonPreferencesStore>();
services.AddSingleton<MainWindow>();
services.AddSingleton<ConnectViewModel>();
// ...
```

All ViewModels and services injected via constructor.

### 4.3 Async + cancellation

Every long-running method exposes:
- `CancellationToken` parameter
- `IProgress<LogEntry>` for streaming progress
- Returns `Task<TResult>` or `Task`

Cancel buttons in UI bind to a `CancellationTokenSource` owned by the ViewModel. On Cancel: `cts.Cancel()` propagates to runspaces (`PowerShell.Stop()`) and Graph client (HTTP cancellation).

### 4.4 State observation

`IConnectionStateMonitor` raises `PropertyChanged` events when Graph or EXO state changes. UI bindings update automatically. Internal poll loop (every 1s in background) checks:
- `GraphServiceClient` has a valid token (no actual call needed; check token cache)
- `Get-ConnectionInformation` via runspace returns Connected

No more "click connect → wait silently → status updates after job ends".

### 4.5 Logging

Serilog config:

```csharp
Log.Logger = new LoggerConfiguration()
    .MinimumLevel.Debug()
    .WriteTo.File("logs/grex365-.log", rollingInterval: RollingInterval.Day, retainedFileCountLimit: 30)
    .WriteTo.Sink(new ObservableLogSink())   // pushes entries into ObservableCollection for UI log panel
    .CreateLogger();
```

`ObservableLogSink` exposes an `ObservableCollection<LogEntry>` bound to a `DataGrid` in the UI. Live log without manual queue draining.

### 4.6 Error handling

- Each method: try/catch only where adding context. Otherwise let exceptions bubble.
- ViewModel command handlers wrap calls in try/catch, log with Serilog, show toast / message box.
- `Application.DispatcherUnhandledException` + `AppDomain.CurrentDomain.UnhandledException` + `TaskScheduler.UnobservedTaskException` all wired to a single global handler that logs full stack trace and prompts user.

---

## 5. PowerShell integration

### 5.1 Why we keep PowerShell

Some Microsoft 365 surface area has **no .NET SDK equivalent**:
- Exchange Online cmdlets (`Get-Mailbox`, `Set-MailboxPermission`, `Set-CalendarProcessing`, etc.) — only available in the `ExchangeOnlineManagement` PS module.
- Some Teams + SharePoint admin cmdlets.

For Microsoft Graph we **prefer the native .NET SDK** (`Microsoft.Graph`), not the PowerShell module. Same API, far less overhead, real async.

### 5.2 RunspacePool design

`Grex365.PowerShell.RunspacePoolHost`:
- Single shared `RunspacePool` (min 1, max 4) created on app startup.
- `InitialSessionState`: PSGallery trusted, `ConfirmPreference=None`, `ProgressPreference=SilentlyContinue`, `ErrorActionPreference=Continue`.
- Apartment state **MTA** (avoid MSAL deadlocks).
- Pre-imports `ExchangeOnlineManagement` on first runspace allocation.

`PowerShellRunner.RunScriptAsync(string script, IDictionary<string,object>? args, IProgress<LogEntry>? progress, CancellationToken ct)`:
1. Acquires a `PowerShell` instance from the pool.
2. Subscribes to `Streams.Information / Warning / Error / Verbose / Debug / Progress` and forwards each event to `progress`.
3. `BeginInvoke()` async.
4. Awaits completion or cancellation. On cancel: `ps.Stop()`.
5. Returns `Collection<PSObject>` or throws.

### 5.3 EXO connection lifecycle

Cert flow:
```powershell
Connect-ExchangeOnline `
  -AppId $AppId `
  -CertificateThumbprint $Thumb `
  -Organization $Org `
  -ShowBanner:$false
```
Wrapped in a single runspace invocation. Connection persists across subsequent runspace allocations (same pool, runspaces reused).

Disconnect on app shutdown: `Disconnect-ExchangeOnline -Confirm:$false`.

### 5.4 Graph connection lifecycle (native .NET)

```csharp
var cert = new X509Certificate2(Cert:\CurrentUser\My\{thumbprint}, ...);
var credential = new ClientCertificateCredential(tenantId, clientId, cert);
var graphClient = new GraphServiceClient(credential, scopes);
```

Token caching automatic. No deadlocks. No prompts.

---

## 6. Coexistence with legacy PowerShell toolkit

The legacy `GREX365/` PowerShell scripts stay in the repo during migration. They are **not modified** after the initial connection hang fix.

Migration strategy per feature (Health, Audit, Groups, Offboarding, etc.):

1. **Port** the script logic to a C# method in `Grex365.Core/Services/{Feature}Service.cs`.
2. **Test** with unit tests + manual run against dev tenant.
3. **Wire** to a new ViewModel + View in `Grex365.App`.
4. **Mark** the legacy script as deprecated in `docs/MIGRATION.md`.
5. **Remove** legacy script after one stable release cycle.

This avoids a big-bang rewrite. The new app gradually replaces the old toolkit.

---

## 7. Build, test, release

### 7.1 Local build

```bash
dotnet restore src/Grex365.sln
dotnet build src/Grex365.sln -c Release
dotnet test src/Grex365.sln -c Release
```

### 7.2 CI (GitHub Actions)

`.github/workflows/ci.yml` runs on every push + PR to `main`:
1. Setup .NET 10 SDK
2. Restore
3. Build Release
4. Run tests
5. (On tags) Publish single-file .exe + upload as release artifact

### 7.3 Publish .exe

```bash
dotnet publish src/Grex365.App/Grex365.App.csproj `
  -c Release `
  -r win-x64 `
  --self-contained `
  -p:PublishSingleFile=true `
  -p:IncludeAllContentForSelfExtract=true `
  -p:EnableCompressionInSingleFile=true
```

Result: `bin/Release/net10.0-windows/win-x64/publish/Grex365.exe` (~70-90 MB, no .NET install required on target).

### 7.4 Code signing

Phase 1 (dev): self-signed cert via `New-SelfSignedCertificate -Type CodeSigningCert`. Test users add to trusted publishers.

Phase 2 (release): purchase code-signing cert (Sectigo OV ~$100/yr, EV ~$300/yr). Sign with `signtool sign /f cert.pfx /p pass /tr http://timestamp.sectigo.com /td sha256 /fd sha256 Grex365.exe`.

### 7.5 Auto-update

Velopack:
- `vpk pack` produces a self-extracting installer + delta updates.
- GitHub Releases is the update feed (no extra hosting).
- App checks for update on startup; downloads + applies in background.

---

## 8. Roadmap (hitos sin fechas)

### H0 — Cimientos ✅ (in this commit)
- ARCHITECTURE.md created
- Solution + 4 projects scaffolded
- CI workflow stub
- README updated with migration status

### H1 — Backend core
- `IPowerShellRunner` + `RunspacePoolRunner`
- `IGraphConnection` (native SDK)
- `IExchangeConnection` (PS runspace)
- `IConnectionStateMonitor` (live state)
- Serilog config
- Unit tests for runner stream forwarding + cancellation

### H2 — Connect feature (the bug that started this) ✅ replaces legacy Connect
- WPF shell with Fluent navigation
- ConnectViewModel with live state, cancel button
- Log panel bound to ObservableLogSink
- Settings view (cert path, tenant id, connection method)

### H3 — Migrate features one by one
Order (cheapest first):
1. Tenant health
2. Identity audit
3. Groups workflow (CSV import/export)
4. Mailbox permissions
5. Offboarding wizard
6. Cert wizard (port 29 PS steps to typed C# wizard)

### H4 — UX polish
- Dark / light theme toggle
- Dashboard with summary cards
- Global error boundary
- Keyboard shortcuts

### H5 — Release v1.0
- Self-contained single-file exe
- Code signing (self-signed → real)
- Velopack auto-update
- GitHub Release

### H6 — Iteration
- Bug fixes, new features, automation

---

## 9. Conventions

### 9.1 Naming
- C# follows standard Microsoft conventions: PascalCase types/methods, camelCase locals, `_camelCase` private fields, `Interface` prefix `I`.
- File per type (one public type per `.cs` file).
- Namespaces match folder layout: `Grex365.Core.Connections`, etc.

### 9.2 Style
- File-scoped namespaces.
- `nullable enable` everywhere.
- No `var` for primitive/literal types; OK for complex `new()` expressions where the type is obvious.
- No regions.
- No comments restating code. Comments only for non-obvious *why*.

### 9.3 Async
- Suffix `Async` on every async method.
- `ConfigureAwait(false)` in `Grex365.Core` and `Grex365.PowerShell` (library code, no SyncContext).
- App layer (`Grex365.App`): default context to flow back to UI thread.
- Never `.Result` / `.Wait()`. Use `await`.

### 9.4 Logging
- Structured logging: `Log.Information("Connecting to {Service} as {Account}", service, account);`
- Levels: `Verbose` for noise, `Debug` for dev info, `Information` for user-relevant events, `Warning` for recoverable issues, `Error` for handled failures, `Fatal` for crashes.

### 9.5 Testing
- xUnit. One test class per production class.
- FluentAssertions for readable asserts.
- Moq for interface mocking.
- Integration tests in a separate project later if needed.

---

## 10. Open questions

These need answering before we hit later milestones. Tracked here so they don't get lost.

- [ ] Code-signing cert: self-signed forever (internal use) or buy OV/EV for SmartScreen?
- [ ] Tenant lock: keep the legacy `EnforceTenantLock` preference or trust app-only auth scopes?
- [ ] Offboarding: should it run as a single transactional workflow or step-by-step with checkpoints?
- [ ] Reports: render in-app (DataGrid + export) or open in Excel?
- [ ] Multi-tenant: out of scope for v1?

---

## 11. References

- Live punch list: [`ROADMAP.md`](ROADMAP.md) — comprehensive H0–H6 status, every sub-item, every open decision
- Per-feature migration table: [`MIGRATION.md`](MIGRATION.md)
- Deep research report: [`../deep-research-report.md`](../deep-research-report.md)
- WPF-UI (Fluent for WPF): https://wpfui.lepo.co/
- CommunityToolkit.Mvvm: https://learn.microsoft.com/dotnet/communitytoolkit/mvvm/
- Microsoft.Graph SDK: https://learn.microsoft.com/graph/sdks/sdks-overview
- Velopack: https://velopack.io/
