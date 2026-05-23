# Migration log: PowerShell toolkit → .NET 10 app

> Tracks what has been ported from `GREX365/` (legacy PS) to `src/` (new .NET app).
> Live source of truth: [`PROGRESS.md`](../PROGRESS.md) (sessions bitácora + 532 tests detail).
> Update on every feature migration.

---

## Status (2026-05-23)

| Feature | Legacy location | New location | Status | Notes |
|---|---|---|---|---|
| Connect (cert flow) | `GREX365/Modules/Connection.ps1` | `Grex365.Core/Connections/GraphConnection.cs` + `Grex365.PowerShell/ExchangeConnection.cs` | 🟢 Ported | Cert + device-code (Azure CLI client); live `CheckLiveAsync`; tenant lock enforced post-auth |
| Preferences | `GREX365/Modules/Preferences.ps1` | `Grex365.Core/Preferences/JsonPreferencesStore.cs` | 🟢 Ported | + legacy importer |
| Cert config | `GREX365/Modules/Preferences.ps1` | `Grex365.Core/Preferences/JsonCertConfigStore.cs` | 🟢 Ported | Same JSON shape as legacy |
| Logging | `GREX365/Modules/Logging.ps1` | `Grex365.App/Services/UiLogSink.cs` + Serilog + `FileAuditLog` JSONL | 🟢 Ported | Live ObservableCollection, rolling file, audit JSONL, AppInsights opt-in |
| Connection state monitor | `GREX365/Modules/Jobs.ps1` (queue polling) | `Grex365.Core/Connections/ConnectionStateMonitor.cs` | 🟢 Ported | Background poll 2s; `INotifyPropertyChanged`; surfaces tenant/account |
| RunspacePool host | `GREX365/GUI/Start-Gui.ps1` (inline) | `Grex365.PowerShell/RunspacePoolHost.cs` | 🟢 Ported | MTA + InitialSessionState; auto-install EXO via `pwsh.exe` external (WindowsApps ACL workaround) |
| Tenant health | `GREX365/Scripts/Show-TenantHealth.ps1` | `Grex365.Core/Health/TenantHealthService.cs` + `TenantHealthViewModel`/`View` | 🟢 Ported | M365-portal-style license cards via `SkuCatalog` (FriendlyName + Category + utilization) |
| Identity audit | `GREX365/Scripts/Invoke-IdentityAudit.ps1` | `Grex365.Core/Audit/IdentityAuditAnalyzer.cs` + `GraphAuditService.RunIdentityAuditAsync` | 🟢 Ported | + parallel 8x via SemaphoreSlim; resilient if `AuditLog.Read.All` missing |
| Groups workflow | `GREX365/Scripts/Invoke-GroupsWorkflow.ps1` + `Add/Export/New-GroupsFromCsv.ps1` | `Grex365.Core/Groups/GraphGroupsService.cs` + `Grex365.PowerShell/DistributionListsService.cs` + `GroupsViewModel`/`View` | 🟢 Ported | Bulk CSV (M365 + DL) with auto-detect type from CSV column + forward-fill GroupName |
| Mailbox permissions / convert | `GREX365/Scripts/Set-SharedMailboxPermissions.ps1` + `Convert-SharedToUserMailbox.ps1` | `Grex365.PowerShell/SharedMailboxService.cs` + `SharedMailboxViewModel`/`View` | 🟢 Ported | FullAccess/SendAs/SendOnBehalf + CSV bulk |
| Offboarding wizard | `GREX365/Scripts/Invoke-OffboardingWizard.ps1` | `Grex365.Core/Offboarding/OffboardingService.cs` + `OffboardingViewModel`/`View` | 🟢 Ported | Composes Users + SharedMailbox; RBAC-gated |
| Onboarding wizard | (new — was missing in legacy) | `Grex365.Core/Onboarding/OnboardingService.cs` + `OnboardingViewModel`/`View` | 🟢 Ported | Create user + UsageLocation + assign SKUs + add to groups |
| Cert wizard (29 pasos) | `GREX365/Modules/CertWizard.ps1` | `Grex365.Core/Certificates/CertificateGenerator.cs` + `GraphAppRegistrationService.cs` + `CertWizardViewModel`/`View` | 🟢 Ported | Self-signed RSA 2048 + auto App Registration creation (Graph/EXO permissions) + admin-consent URL + ExportPfx |
| DNS check | (new — not in legacy) | `Grex365.Core/DomainChecks/DomainChecker.cs` + `DomainCheckViewModel`/`View` | 🟢 Ported | MX/TXT/SPF/DMARC via nslookup |
| Mail flow rules viewer | (new) | `Grex365.PowerShell/MailFlowRulesService.cs` + `MailFlowRulesViewModel`/`View` | 🟢 Ported | `Get-TransportRule` lister with filter |
| Mailbox rules (OOO/forwarding/calendar) | (extracted from above) | `Grex365.PowerShell/MailboxRulesService.cs` + `MailboxRulesViewModel`/`View` | 🟢 Ported | OOO state + msg interno/externo + range; forwarding SMTP; calendar folder permissions |
| Audit log viewer | (new) | `Grex365.Core/Audit/FileAuditLog.cs` + `MetricsAggregator.cs` + `AuditLogViewModel`/`View` | 🟢 Ported | JSONL persistente + summary cards (totales, error rate, last 24h, top sources) |
| Audit report exports (HTML/JSON/Baseline diff) | (new) | `AuditReportHtmlBuilder.cs` + `AuditReportJsonBuilder.cs` + `AuditBaselineComparer.cs` | 🟢 Ported | HTML standalone con CSS embebido (light + prefers-color-scheme dark), JSON schema `grex365.audit.v1` parseable, baseline diff New/Resolved/Persistent contra export previo |
| Consola PowerShell embebida | (new) | `Grex365.App/ViewModels/PsConsoleViewModel.cs` + `PsConsoleView` | 🟢 Ported | REPL multi-line reusa `IPowerShellRunner` (RunspacePool shared con app); Graph/EXO en scope; history navegable Up/Down max 50 dedupe; Ctrl+Enter run, Esc cancel. Cierra Plantamiento §6 backlog "Terminal PS embebido" sin EasyWindowsTerminalControl |
| First-Run Wizard | (new) | `Grex365.App/FirstRunWizardWindow.xaml` + `FirstRunWizardViewModel.cs` | 🟢 Ported | Modal 5 páginas: Welcome → Connection (device-code/cert) → TenantLock (id+domain) → Theme (Dark/Light/Auto) → Summary. Persist preferences |
| Plugin system | (not in legacy) | `Grex365.Core/Plugins/PluginLoader.cs` + `IModule` + `Settings` enable/disable UI | 🟢 Ported | `AssemblyLoadContext` per DLL from `%LOCALAPPDATA%\Grex365\plugins\*.dll`. SamplePlugin POC en `samples/` |
| RBAC guard | (not in legacy) | `Grex365.Core/Security/RbacGuard.cs` + `GraphMembershipChecker.cs` | 🟢 Ported | `/me/checkMemberGroups` against `AuthorizationGroupId`; gates destructive ops only |
| Tenant lock | (not in legacy) | `Grex365.Core/Connections/TenantLock.cs` | 🟢 Ported | Enforced post-Graph-connect (cert + device-code paths); aborts + disconnects on mismatch |
| Telemetry | (not in legacy) | `Grex365.App/Services/ApplicationInsightsTelemetry.cs` + `ITelemetry`/`NullTelemetry` | 🟢 Ported | Opt-in via conn string en Settings; `UiLogSink` forwards Ok/Warn/Error |
| Role / UI mode | `GREX365/Modules/Roles.ps1` | — | ⚪ Deferred | Solo operator — superseded by RBAC guard (membership-based) |
| Templates | `GREX365/Modules/Templates.ps1` | — | ⚪ Deferred | Open decision — not yet needed |
| Reports | `GREX365/Modules/Report.ps1` | — | 🔴 Pending | Open decision: format (CSV / XLSX / HTML / DataGrid only) |

### Security audits (all new — go beyond legacy)

| Audit | Implementation | Source | Status |
|---|---|---|---|
| Identity (stale/disabled+licensed) | `IdentityAuditAnalyzer` + `RunIdentityAuditAsync` | Graph `/users` | 🟢 |
| Groups (owner-less, empty, guests in private) | `RunGroupsAuditAsync` | Graph `/groups` | 🟢 |
| Group activity (no recent usage) | `GroupActivityAnalyzer` + `RunGroupActivityAuditAsync` | Graph `/reports/getOffice365GroupsActivityDetail` | 🟢 |
| MFA coverage | `MfaCoverageAnalyzer` + `RunMfaCoverageAuditAsync` | Graph `/reports/authenticationMethods/userRegistrationDetails` | 🟢 |
| Conditional Access policies | `CaPolicyAnalyzer` + `RunConditionalAccessAuditAsync` | Graph `/identity/conditionalAccess/policies` | 🟢 |
| Privileged roles | `PrivilegedRoleAuditAnalyzer` + `RunPrivilegedRolesAuditAsync` | Graph `/directoryRoles` + members | 🟢 |
| App credentials expiry | `AppCredentialAuditAnalyzer` + `RunAppCredentialsAuditAsync` | Graph `/applications` | 🟢 |
| Tenant defaults | `TenantDefaultsAnalyzer` + `RunTenantDefaultsAuditAsync` | Graph `/policies/authorizationPolicy` + `identitySecurityDefaultsEnforcementPolicy` | 🟢 |
| OAuth consent grants | `OAuthGrantAnalyzer` + `RunOAuthGrantsAuditAsync` | Graph `/oauth2PermissionGrants` + `/servicePrincipals` | 🟢 |
| Mailbox forwarding externo | `MailboxForwardingAnalyzer` + `ExoForwardingAuditService.ScanForwardingAsync` | EXO `Get-Mailbox` + `Get-AcceptedDomain` | 🟢 |
| Inbox rules (BEC indicator) | `InboxRuleAnalyzer` + `ScanInboxRulesAsync` | EXO `Get-InboxRule` per mailbox | 🟢 |
| Transport rules | `TransportRuleAuditAnalyzer` + `ScanTransportRulesAsync` | EXO `Get-TransportRule` + `Get-OutboundConnector` | 🟢 |
| Shared mailbox sign-in | `SharedMailboxSignInAnalyzer` + `ScanSharedMailboxSignInAsync` | EXO `Get-Mailbox` + `Get-User` | 🟢 |

Legend: 🟢 ported · 🟡 skeleton only · 🔴 pending · ⚪ deferred · ⚫ deleted

---

## Procedure per feature

1. Read legacy script + understand its inputs/outputs/side effects.
2. Write/update C# method in `Grex365.Core/Services/{Feature}Service.cs` (create folder when first service lands).
3. Add unit tests in `tests/Grex365.Core.Tests/`.
4. Wire to a ViewModel + View (or extend existing).
5. Update this table.
6. Mark legacy script as `# DEPRECATED — see src/Grex365.Core/...` (top of file).
7. After one stable release using the new path, delete legacy script.

---

## Open decisions per pending feature

- **Offboarding (D2)**: transactional rollback on partial failure? Or step-by-step with manual recovery? — Currently: step-by-step, fail-soft (every step has try/catch + reports per-step). Decision: keep as-is unless real ops show data inconsistencies.
- **Reports (D3)**: render in-app `DataGrid` + CSV export only, or also XLSX/HTML? — All security audits already export to CSV from `AuditView`. No standalone "reports" module yet. Decision: revisit if user explicitly asks.
- **Roles/UI modes (D4)**: superseded by RBAC guard (membership-based). Closed.
- **Templates (D5)**: keep or drop for v1? — Open. No demand yet.
- **i18n (D6)**: Spanish-only is sufficient for current operator. Open if multi-language teams adopt.
- **Code-signing cert (D7)**: self-signed vs purchased OV/EV. Open — depends on SmartScreen friction in real distribution.
- **Min OS target (D8)**: Win10 1809+ vs Win11-only. Open — `.NET 10` runtime supports both.

### Resolved

- **Cert wizard**: auto-creates App Registration via Graph (`GraphAppRegistrationService.CreateAndConfigureAsync`) with all required AppRoles + cert upload + admin-consent URL. Manual fallback in legacy docs.
- **Tenant lock (D9)**: kept. Enforced post-Graph-connect (cert + device-code).
- **Terminal PowerShell embebido**: shipped vía `PsConsoleView` (Sprint E 2026-05-22) — reusa runspace existente en vez de integrar `EasyWindowsTerminalControl`.
- **Docs internas (Fase 6)**: shipped Sprint P 2026-05-23 — ARCHITECTURE.md refresh + RUNBOOK.md nuevo + ROADMAP.md refresh + MIGRATION.md sync. Closes Plantamiento Fase 6 item "Crear documentación técnica interna".

---

## Renames de nav titles (2026-05-22 / 2026-05-23)

Acumulación de pequeños renames cubierta por `NavTitleMigrator.RenameMap` para preservar `LastSelectedNavigation` de usuarios beta.

| Old | New | Cuando |
|---|---|---|
| Conexion | Conexión | 2026-05-22 |
| Auditoria | Auditoría | 2026-05-22 |
| Reglas buzon | Reglas de buzón | 2026-05-22 |
| Mail flow | Flujo de correo | 2026-05-22 |
| Audit log | Registro de auditoría | 2026-05-22 |
| Cert Wizard | Asistente cert | 2026-05-22 |
| DNS check | Comprobación DNS | 2026-05-22 |
| Salud tenant | Licencias | 2026-05-23 (Sprint O) |
