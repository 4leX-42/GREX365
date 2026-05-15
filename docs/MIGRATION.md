# Migration log: PowerShell toolkit → .NET 10 app

> Tracks what has been ported from `GREX365/` (legacy PS) to `src/` (new .NET app).
> Update on every feature migration.

---

## Status

| Feature | Legacy location | New location | Status | Notes |
|---|---|---|---|---|
| Connect (cert flow) | `GREX365/Modules/Connection.ps1` | `src/Grex365.Core/Connections/GraphConnection.cs` + `src/Grex365.PowerShell/ExchangeConnection.cs` | 🟡 Skeleton | Graph native SDK; EXO via runspace |
| Preferences | `GREX365/Modules/Preferences.ps1` | `src/Grex365.Core/Preferences/JsonPreferencesStore.cs` | 🟢 Ported | JSON roundtrip tested |
| Cert config | `GREX365/Modules/Preferences.ps1` | `src/Grex365.Core/Preferences/JsonCertConfigStore.cs` | 🟢 Ported | Same JSON shape as legacy |
| Logging | `GREX365/Modules/Logging.ps1` | `src/Grex365.App/Services/UiLogSink.cs` + Serilog | 🟢 Ported | Live ObservableCollection, rolling file |
| Connection state monitor | `GREX365/Modules/Jobs.ps1` (queue polling) | `src/Grex365.Core/Connections/ConnectionStateMonitor.cs` | 🟡 Skeleton | Background poll loop, `INotifyPropertyChanged` |
| RunspacePool host | `GREX365/GUI/Start-Gui.ps1` (inline) | `src/Grex365.PowerShell/RunspacePoolHost.cs` | 🟢 Ported | MTA + InitialSessionState |
| Tenant health | `GREX365/Scripts/Show-TenantHealth.ps1` | — | 🔴 Pending | H3 |
| Identity audit | `GREX365/Scripts/Invoke-IdentityAudit.ps1` | — | 🔴 Pending | H3 |
| Groups workflow | `GREX365/Scripts/Invoke-GroupsWorkflow.ps1` | — | 🔴 Pending | H3 |
| Add group members | `GREX365/Scripts/Add-GroupMembers.ps1` | — | 🔴 Pending | H3 |
| Export group members | `GREX365/Scripts/Export-GroupMembers.ps1` | — | 🔴 Pending | H3 |
| New groups from CSV | `GREX365/Scripts/New-GroupsFromCsv.ps1` | — | 🔴 Pending | H3 |
| Mailbox permissions | `GREX365/Scripts/Set-SharedMailboxPermissions.ps1` | — | 🔴 Pending | H3 |
| Shared mailbox convert | `GREX365/Scripts/Convert-SharedToUserMailbox.ps1` | — | 🔴 Pending | H3 |
| Offboarding wizard | `GREX365/Scripts/Invoke-OffboardingWizard.ps1` | — | 🔴 Pending | H3 |
| Cert wizard (29 pasos) | `GREX365/Modules/CertWizard.ps1` | — | 🔴 Pending | H3 — port to typed C# wizard |
| Role / UI mode | `GREX365/Modules/Roles.ps1` | — | 🔴 Pending | H2/H4 |

Legend: 🟢 ported · 🟡 skeleton only · 🔴 pending · ⚪ deprecated · ⚫ deleted

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

- **Cert wizard**: 29-step console wizard → multi-page WPF wizard with `System.Security.Cryptography.X509Certificates.CertificateRequest`. Need to decide if we keep app-registration creation steps (require Graph admin consent) inside the app or document as manual.
- **Offboarding**: transactional rollback on partial failure? Or step-by-step with manual recovery?
- **Reports**: render in-app `DataGrid` + CSV export only, or also XLSX?
