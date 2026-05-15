# GREX365 — Roadmap & status

> Comprehensive punch list. Updated on every meaningful change.
> Date created: 2026-05-15. Last update: 2026-05-15.
>
> Legend: ✅ done · 🟡 in progress · 🔴 not started · ⚪ deferred · ❌ rejected
>
> Effort scale (solo autodidacta, part-time):
> - S = small (< 1 session)
> - M = medium (1-3 sessions)
> - L = large (3-10 sessions)
> - XL = extra-large (10+ sessions)

---

## Realistic total effort

Initial estimate "3-5 weeks" assumed full-time dev. Solo + part-time autodidacta reality: **3-4 months** to reach a usable v1.0 that fully replaces the PS toolkit. Below is the full breakdown.

Approximate effort summary:
- H0 (cimientos): S ✅
- H1 (backend core): L
- H2 (Connect feature complete): M
- H3 (port 8-10 features): XL
- H4 (UX polish): M
- H5 (release v1.0): M
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
| 1.1.10 | Handle PSGallery first-time install (timeout, fallback) | 🔴 | M | Real-world: 60s install with no output is unacceptable |
| 1.1.11 | Progress events for module install | 🔴 | M | Verbose stream → progress |
| 1.1.12 | Reset runspace state on error (avoid contaminated reuse) | 🔴 | M | Currently uses default thread options |

### 1.2 GraphConnection (native SDK)

| # | Item | Status | Effort | Notes |
|---|------|--------|--------|-------|
| 1.2.1 | `IGraphConnection` interface | ✅ | S | |
| 1.2.2 | `ClientCertificateCredential` flow | ✅ | S | |
| 1.2.3 | Smoke test via `Organization.GetAsync` | ✅ | S | |
| 1.2.4 | Cert loading from CurrentUser\My | ✅ | S | |
| 1.2.5 | Real connection state (token validity, not just IsConnected flag) | ✅ | M | `CheckLiveAsync` probes Graph with 10s cache |
| 1.2.6 | Tenant lock enforcement (compare actual TenantId vs expected) | 🔴 | S | Port from legacy `Assert-TenantLock` |
| 1.2.7 | Scope handling (currently hardcoded `.default`) | 🔴 | S | Keep app-only for cert; explicit scopes for device code |
| 1.2.8 | Device-code / traditional flow | 🔴 | M | `DeviceCodeCredential` |
| 1.2.9 | Connection state cache (avoid re-auth per call) | 🔴 | S | SDK handles via token cache |
| 1.2.10 | Unit tests with Moq | 🔴 | S | |

### 1.3 ExchangeConnection (runspace)

| # | Item | Status | Effort | Notes |
|---|------|--------|--------|-------|
| 1.3.1 | `IExchangeConnection` interface | ✅ | S | |
| 1.3.2 | `Connect-ExchangeOnline` cert flow | ✅ | S | |
| 1.3.3 | Module ensure (install + import) | ✅ | S | |
| 1.3.4 | Disconnect | ✅ | S | |
| 1.3.5 | Persistent session across runspace pool | 🔴 | M | Currently each runspace lacks connection state |
| 1.3.6 | Real `Test-ExchangeOnlineConnected` via runspace | ✅ | S | `CheckLiveAsync` calls Get-ConnectionInformation |
| 1.3.7 | Tenant lock enforcement | 🔴 | S | |
| 1.3.8 | Device-code flow | 🔴 | M | |
| 1.3.9 | Integration test (mock or real tenant) | 🔴 | M | |

### 1.4 ConnectionStateMonitor

| # | Item | Status | Effort | Notes |
|---|------|--------|--------|-------|
| 1.4.1 | `IConnectionStateMonitor` interface | ✅ | S | |
| 1.4.2 | 1s poll loop with cancellation | ✅ | S | |
| 1.4.3 | `INotifyPropertyChanged` plumbing | ✅ | S | |
| 1.4.4 | Real check vs `IGraphConnection.IsConnected` | ✅ | S | Calls `CheckLiveAsync` per tick |
| 1.4.5 | Real check vs `Get-ConnectionInformation` runspace | ✅ | M | 10s cache TTL prevents stampede |
| 1.4.6 | Surface tenant + account info in state | ✅ | S | TenantId, Organization, Account flow to UI |
| 1.4.7 | Unit tests | 🔴 | S | |

### 1.5 Preferences + cert config

| # | Item | Status | Effort | Notes |
|---|------|--------|--------|-------|
| 1.5.1 | `IPreferencesStore` + JSON impl | ✅ | S | |
| 1.5.2 | `ICertConfigStore` + JSON impl | ✅ | S | |
| 1.5.3 | Unit tests roundtrip | ✅ | S | |
| 1.5.4 | Schema version + migration logic | 🔴 | S | Future-proof if shape changes |
| 1.5.5 | Validation on load (corrupted file → default + warn) | 🔴 | S | |
| 1.5.6 | Read legacy paths if found (`GREX365/config/*.json`) | ✅ | S | `LegacyPreferencesImporter` invoked on App startup |

### 1.6 Logging

| # | Item | Status | Effort | Notes |
|---|------|--------|--------|-------|
| 1.6.1 | Serilog rolling file config | ✅ | S | |
| 1.6.2 | `UiLogSink` ObservableCollection | ✅ | S | |
| 1.6.3 | Audit log separate file (who did what when) | 🔴 | M | Compliance-relevant ops only |
| 1.6.4 | Log severity filter in UI | 🔴 | S | Hide Debug by default |
| 1.6.5 | Log export (copy/save to file) | 🔴 | S | |

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
| 2.8 | Settings view (cert path, tenant id, connection method) | 🔴 | M | Bind to `IPreferencesStore` |
| 2.9 | First-run wizard (no config exists) | 🔴 | M | Guide user to set cert + tenant |
| 2.10 | Smoke test against real tenant | 🔴 | M | Manual; documents the flow |
| 2.11 | Replace fake `IsConnected` with real state | ✅ | M | Wired to H1.2.5 + H1.3.6 + dispatcher marshalling |
| 2.12 | Cert config validation UI (warn if cert expired, missing in store) | ✅ | S | `ICertValidator` runs before Connect, blocks if invalid |
| 2.13 | Theme toggle (light/dark) | 🔴 | S | Read from prefs |
| 2.14 | Window restore (size, position) on relaunch | 🔴 | S | |
| 2.15 | App icon + branding | 🔴 | S | |
| 2.16 | About dialog (version, repo link) | 🔴 | S | |

---

## H3 — Migrate features 🔴

One section per legacy feature. Each typically: read PS script → port logic to C# service → unit tests → ViewModel + View → update `MIGRATION.md`.

### 3.1 Tenant health (L)

Legacy: `GREX365/Scripts/Show-TenantHealth.ps1` (20.5 KB, 500+ lines)

| # | Item | Status | Effort | Notes |
|---|------|--------|--------|-------|
| 3.1.1 | Read legacy + identify Graph calls | 🔴 | S | |
| 3.1.2 | Port to `Grex365.Core/Services/TenantHealthService.cs` | 🔴 | M | Use `GraphServiceClient` |
| 3.1.3 | `TenantHealthViewModel` + `TenantHealthView.xaml` | 🔴 | M | DataGrid + summary cards |
| 3.1.4 | Cancellation support | 🔴 | S | |
| 3.1.5 | CSV/clipboard export | 🔴 | S | |
| 3.1.6 | Unit tests with mocked GraphServiceClient | 🔴 | M | |
| 3.1.7 | Mark legacy as DEPRECATED | 🔴 | S | |

### 3.2 Identity audit (L)

Legacy: `GREX365/Scripts/Invoke-IdentityAudit.ps1` (6.8 KB)

| # | Item | Status | Effort | Notes |
|---|------|--------|--------|-------|
| 3.2.1 | Port logic to `IdentityAuditService` | 🔴 | M | |
| 3.2.2 | VM + View | 🔴 | M | |
| 3.2.3 | Tests | 🔴 | M | |
| 3.2.4 | Deprecate legacy | 🔴 | S | |

### 3.3 Groups workflow (XL)

Legacy combined: `Invoke-GroupsWorkflow.ps1` + `Add-GroupMembers.ps1` + `Export-GroupMembers.ps1` + `New-GroupsFromCsv.ps1` (~50 KB)

| # | Item | Status | Effort | Notes |
|---|------|--------|--------|-------|
| 3.3.1 | Port `GroupResolver.ps1` to `GroupResolverService` | 🔴 | M | Search + resolve identities |
| 3.3.2 | Port Add-GroupMembers | 🔴 | M | CSV input |
| 3.3.3 | Port Export-GroupMembers | 🔴 | M | CSV output |
| 3.3.4 | Port New-GroupsFromCsv | 🔴 | L | M365 Groups + DLs, `-WhatIf` support |
| 3.3.5 | Unified Groups VM/View | 🔴 | L | Three tabs: Add / Export / Create |
| 3.3.6 | CSV browse dialog + preview | 🔴 | S | |
| 3.3.7 | Drag-and-drop CSV onto window | 🔴 | S | |
| 3.3.8 | Tests | 🔴 | L | |
| 3.3.9 | Deprecate legacy | 🔴 | S | |

### 3.4 Mailbox permissions (M)

Legacy: `Set-SharedMailboxPermissions.ps1` + `Convert-SharedToUserMailbox.ps1`

| # | Item | Status | Effort | Notes |
|---|------|--------|--------|-------|
| 3.4.1 | Port permissions logic | 🔴 | M | EXO runspace |
| 3.4.2 | Port shared→user convert | 🔴 | M | |
| 3.4.3 | VM + View | 🔴 | M | |
| 3.4.4 | Tests | 🔴 | M | |
| 3.4.5 | Deprecate legacy | 🔴 | S | |

### 3.5 Offboarding wizard (XL)

Legacy: `Invoke-OffboardingWizard.ps1` (22.6 KB, 600+ lines) — most complex feature

| # | Item | Status | Effort | Notes |
|---|------|--------|--------|-------|
| 3.5.1 | Read full wizard, map steps | 🔴 | S | |
| 3.5.2 | Decide: transactional rollback vs checkpoint? | 🔴 | S | OPEN DECISION |
| 3.5.3 | `OffboardingService` with step abstraction | 🔴 | L | Pipeline pattern |
| 3.5.4 | Steps: disable user, revoke licenses, fwd mailbox, transfer OneDrive, etc. | 🔴 | L | |
| 3.5.5 | Multi-page wizard View | 🔴 | L | Step indicator + summary |
| 3.5.6 | Dry-run mode (`-WhatIf` equivalent) | 🔴 | M | |
| 3.5.7 | Progress + cancel mid-pipeline | 🔴 | M | |
| 3.5.8 | Tests | 🔴 | L | |
| 3.5.9 | Deprecate legacy | 🔴 | S | |

### 3.6 Cert wizard (XL)

Legacy: `GREX365/Modules/CertWizard.ps1` (28.3 KB, 628 lines) — 29 interactive steps

| # | Item | Status | Effort | Notes |
|---|------|--------|--------|-------|
| 3.6.1 | Map 29 steps to typed C# state machine | 🔴 | M | |
| 3.6.2 | Self-signed cert generation (`CertificateRequest`) | 🔴 | M | |
| 3.6.3 | App registration creation via Graph | 🔴 | L | Needs Graph admin consent |
| 3.6.4 | Service Principal + directory roles | 🔴 | L | |
| 3.6.5 | AppRoles Graph + Exchange grants | 🔴 | L | |
| 3.6.6 | Multi-page wizard UI | 🔴 | L | |
| 3.6.7 | Save config to `exo-app-params.json` | 🔴 | S | |
| 3.6.8 | Cleanup / uninstall flow | 🔴 | M | Port `Remove-CertConfig` |
| 3.6.9 | Tests (mostly Graph SDK mocked) | 🔴 | L | |
| 3.6.10 | Deprecate legacy | 🔴 | S | |

### 3.7 Roles + UI modes (S)

Legacy: `Roles.ps1` — operator vs admin, support vs advanced UI

| # | Item | Status | Effort | Notes |
|---|------|--------|--------|-------|
| 3.7.1 | Decide if needed in v1 (you're sole operator) | 🔴 | S | OPEN DECISION — probably ⚪ defer |
| 3.7.2 | If yes: port to `IRoleService` + view visibility | ⚪ | M | |

### 3.8 Templates (S)

Legacy: `Templates.ps1` — offboarding templates

| # | Item | Status | Effort | Notes |
|---|------|--------|--------|-------|
| 3.8.1 | Decide if needed v1 | 🔴 | S | OPEN DECISION |
| 3.8.2 | Port if yes | ⚪ | S | |

### 3.9 Reports (M)

Legacy: `Report.ps1` — generates summary reports of runs

| # | Item | Status | Effort | Notes |
|---|------|--------|--------|-------|
| 3.9.1 | Decide format: in-app DataGrid + CSV vs XLSX vs HTML | 🔴 | S | OPEN DECISION |
| 3.9.2 | Implement | 🔴 | M | |
| 3.9.3 | Tests | 🔴 | S | |

---

## H4 — UX polish 🔴

| # | Item | Status | Effort | Notes |
|---|------|--------|--------|-------|
| 4.1 | Dashboard home screen | 🔴 | M | Cards: tenant health summary, recent ops, connection state |
| 4.2 | Theme toggle (light/dark/auto from system) | 🔴 | S | |
| 4.3 | Sidebar navigation (NavigationView with Frame) | 🔴 | M | Replace flat layout |
| 4.4 | Keyboard shortcuts (Ctrl+, settings, Ctrl+L logs, etc.) | 🔴 | S | |
| 4.5 | Empty states for every list / panel | 🔴 | S | |
| 4.6 | Notification toast for completion + errors | 🔴 | M | |
| 4.7 | Confirmation dialogs for destructive ops | 🔴 | S | |
| 4.8 | Per-monitor DPI testing | 🔴 | S | |
| 4.9 | Accessibility pass (keyboard nav, screen reader) | 🔴 | M | |
| 4.10 | App icon + splash | 🔴 | S | |
| 4.11 | Spanish/English locale toggle | 🔴 | M | OPEN DECISION — needed? |

---

## H5 — Release v1.0 🔴

| # | Item | Status | Effort | Notes |
|---|------|--------|--------|-------|
| 5.1 | `dotnet publish` single-file `.exe` works | ✅ | S | Tested locally — pipeline ready |
| 5.2 | Self-signed code-signing cert | 🔴 | S | `New-SelfSignedCertificate -Type CodeSigningCert` |
| 5.3 | `signtool` integrated into publish step | 🔴 | M | |
| 5.4 | Velopack auto-update | 🔴 | M | GitHub Releases as feed |
| 5.5 | GitHub Releases automated on tag | 🔴 | S | CI workflow already stubs this |
| 5.6 | Release notes template | 🔴 | S | |
| 5.7 | Versioning scheme (SemVer + Directory.Build.props) | 🔴 | S | |
| 5.8 | Install / uninstall docs | 🔴 | S | |
| 5.9 | Real (purchased) code-signing cert decision | 🔴 | S | OPEN — SmartScreen reputation |
| 5.10 | Smoke test on clean Win10/Win11 VMs | 🔴 | M | |

---

## H6 — Post v1.0 / iteration ⚪

- Plugin system (Prism / MEF) — only if real demand
- MSIX packaging — only if corporate deployment needed
- Application Insights — only if multi-operator
- Multi-tenant support — out of scope v1
- Cross-platform (Avalonia) — out of scope, Windows-only target

---

## Open decisions tracker

These block later work. Resolve before reaching their hito.

| # | Decision | Blocks | Status |
|---|----------|--------|--------|
| D1 | Branching strategy (trunk vs feature branches) | H0.12 | 🔴 |
| D2 | Offboarding: transactional or checkpoint? | H3.5 | 🔴 |
| D3 | Report format: CSV / XLSX / HTML / DataGrid only? | H3.9 | 🔴 |
| D4 | Roles + UI modes: keep or drop for v1? | H3.7 | 🔴 — recommend drop |
| D5 | Templates: keep or drop for v1? | H3.8 | 🔴 |
| D6 | i18n: Spanish only, or Spanish + English? | H4.11 | 🔴 |
| D7 | Code-signing cert: self-signed forever or buy real? | H5.9 | 🔴 |
| D8 | Min target OS: Win10 1809+ or Win11 only? | many | 🔴 |
| D9 | Tenant lock: keep legacy preference? | H1.2.6 | 🔴 |

---

## What to work on next (single source of truth)

**Right now**: H0 done. The next session should pick one of:

1. **H1.1.6–1.1.11**: solidify `PowerShellRunner` (tests + install progress). Necessary for everything else.
2. **H2.11**: real connection state (replace fake `IsConnected`). Unlocks the actual bug fix demo.
3. **H2.9**: first-run wizard so the app is usable on a fresh machine.

Recommendation: **H2.11 first** (shortest path to "Connect actually works and shows live state"), then H1 tests, then H3 features one-by-one.
