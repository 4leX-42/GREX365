# GREX365 v2.0 — Progress log

> **Documento maestro de seguimiento.** Mapea el estado del proyecto contra `Plantamiento_arquitectura_de_la_herramienta.md` (roadmap arquitectónico) y `deep-research-report.md` (research técnico). Toda feature shipped y todo pendiente vive aquí.

- Branch: `grex365-2.0` · Pushed up to `origin/grex365-2.0`
- Stack actual: **C# · .NET 10 · WPF + wpf-ui (Fluent) · MVVM (CommunityToolkit.Mvvm) · Serilog · Microsoft.Extensions.Hosting · Microsoft.ApplicationInsights**
- Tests: **326 passing** (xUnit + FluentAssertions)
- Última actualización: 2026-05-22

## Auditoría técnica integral 2026-05-22

**Alcance**: estructura proyecto + auth/permisos + dead code + naming + theme/paleta + module consistency. Tres pases en paralelo via Explore agents.

### A. Estructura + dead code · OK
- 3 proyectos: `Grex365.Core` (lib) / `Grex365.PowerShell` (helpers) / `Grex365.App` (WPF) + `Grex365.Core.Tests` (xUnit).
- 9 converters todos referenciados desde XAML/App.xaml.
- 14 ViewModels todos registrados en DI. `SettingsViewModel` no expuesto vía DataTemplate (intencional: window separado).
- NuGet packages todos en uso (CommunityToolkit.Mvvm, Wpf.Ui, Microsoft.Graph, Serilog sinks, ApplicationInsights).
- Config dir `GREX365/config/`: `user_preferences.json` + `exo-app-params.json`. Sin duplicados.

### B. Auth / permisos · 2 hallazgos MEDIUM
- AppReg auto-create permisos: 9 AppRoles Graph + 1 EXO (Exchange.ManageAsApp). Bien cubierto.
- Tenant lock enforced en auto-connect + manual connect (cert + device-code). Pero **device-code allows `organizations` tenant** (multi-tenant login) — el TenantLock es el único safeguard, y puede ser BYPASS si `_graph.TenantId` es null tras login (ConnectViewModel.cs:227). Cerrar.
- RBAC gateado en TODOS los comandos destructivos (Users/Groups/SharedMailbox/MailboxRules). App-only auth bypass RBAC (intencional — no hay contexto `me.CheckMemberGroups`).
- Cert validator usa `validOnly: false` en store lookup, pero check NotBefore/NotAfter en validación — coherente.
- No hardcoded secrets. Client ID Azure CLI hardcoded (intencional, public).

### C. Theme palette / consistencia visual · CRÍTICO
- App.xaml define **solo Dark theme** (`<ui:ThemesDictionary Theme="Dark" />`). Toggle vía `ApplicationThemeManager.Apply()` pero NO respeta hardcoded colors.
- **16+ ubicaciones con colores hex hardcoded** en XAML — no responden al toggle Light/Dark.
- **4 converters con RGB hardcoded** (SeverityToBrush, AuditSeverityToBrush, UtilizationToBrush, BoolToBrush) — congelados, no theme-aware.
- Page title FontSize inconsistente: 24px (9 views) vs 28px (TenantHealth + Connect).
- Page Grid Margin inconsistente: `24` plano vs `32,28,32,16` (PageRoot style).
- NavigationItem titles mix Español/Inglés: Dashboard/Conexion/Salud tenant/Usuarios/Grupos/Buzones/Reglas buzon/Auditoria EN español pero "Mail flow"/"Audit log"/"Cert Wizard"/"DNS check" en inglés.

### Plan de refactor (orden ejecución)
1. Brushes semánticas Severity/Utilization en App.xaml + variantes light/dark via wpf-ui ApplicationTheme aware Color resources.
2. Refactor converters a usar `Application.Current.Resources["BrushSemanticError"]` lookup (dynamic).
3. Limpieza hardcoded en XAML — `DynamicResource` apunta a brushes semánticas.
4. Normalizar nav titles a español 100%.
5. Standardize page margins + FontSize (24px + 32,28,32,16 PageRoot).
6. Cerrar bypass tenant lock device-code null tenant.

### Refactor ejecutado 2026-05-22 (post-audit)
- ✓ Paleta semántica unificada en App.xaml: Color + Brush resources `SemanticError/Warn/Info/Ok/Neutral/Debug` + Soft variants (alpha 0x55) + MutedText variants. 16 Color + 14 Brush keys.
- ✓ 4 converters (SeverityToBrush, AuditSeverityToBrush, UtilizationToBrush, BoolToBrush) refactorizados — lookup dinámico vía `Application.Current.TryFindResource(key)` en lugar de hardcoded RGB. Theme toggle ahora afecta a colores derivados.
- ✓ 19 reemplazos hardcoded → DynamicResource en MainWindow + AuditView + AuditLogView + DashboardView.
- ✓ NavigationItem titles normalizados a español: "Mail flow" → "Flujo de correo", "Audit log" → "Registro de auditoría", "Cert Wizard" → "Asistente cert", "DNS check" → "Comprobación DNS". `RequiresExchangeTitles` set actualizado.
- ✓ 10 views normalizadas: `Grid Margin="24"` / `StackPanel Margin="24"` → `32,28,32,16` (coherencia con `PageRoot` Padding).
- ✓ Tenant lock bypass cerrado: `ConnectViewModel.ConnectByDeviceCodeAsync` ahora aborta si TenantId queda null tras login en lugar de continuar sin enforcement.

326 tests siguen verdes. Build clean.

### Post-refactor sweep
- DashboardView pills Foreground refactor también a `BrushSemanticError/Warn/Info` (eran inline hex en `Run.Foreground`).
- Surface tokens añadidos en App.xaml: `BrushSurfaceDivider` (#33808080 — gris translúcido válido en light/dark). Disponible para dividers consistentes.
- Sweep final `grep "Foreground=\"#\|Background=\"#"`: cero matches restantes — paleta 100% via DynamicResource.

## Bitácora sesiones

### 2026-05-22 (Sprint G) — First-run wizard + DataGrid real fix

User feedback Sprint F: "Auditoría sigue fallando, color de mierda" → diagnosis: WPF `DataGridTextColumn` y `GridViewColumn.DisplayMemberBinding` generan TextBlocks con Foreground hardcoded a `SystemColors.WindowTextBrush` via internal `SyncColumnProperty` — global Style en DataGrid/ListView NO penetra esos TextBlocks generados.

Real fix:
- Nuevos x:Key Styles `DataCellText` + `DataMonoCellText` en App.xaml (Foreground=`TextFillColorPrimaryBrush` + VerticalCenter + Margin + TextTrimming).
- `AuditLogView` 5 DataGridTextColumn: añade `ElementStyle="{StaticResource DataCellText}"`.
- `MailFlowRulesView` 5 DataGridTextColumn: igual.
- `AuditView` GridView Categoría/Identity/Detalle: reemplaza `DisplayMemberBinding` por `CellTemplate` con TextBlock styled.
- `DomainCheckView` GridView Tipo/Estado/Valor: misma conversión.

Sprint G: First-run wizard explícito (H2.9):
- `FirstRunWizardWindow.xaml` FluentWindow modal 780x640 5 páginas + step indicator pills (semantic accent active).
- `FirstRunWizardViewModel` state machine `FirstRunStep` enum + ConnectionMethod/EnforceTenantLock/ExpectedTenantId/ExpectedTenantDomain/Theme props + Next/Back/Skip/Finish commands con CanExecute.
- Pages: Welcome (intro card) + Connection (RadioButton devicecode/cert) + TenantLock (CheckBox + ID/Dominio TextBoxes disabled cuando off) + Theme (Dark/Light RadioButton) + Summary (table valores elegidos).
- Skip: solo set `FirstRunCompleted=true`. Finish: persiste todos los prefs.
- Converters nuevos: `EnumToVisibilityConverter` + `StringEqualsConverter` (último para RadioButton IsChecked binding string).
- `App.xaml.cs.OnStartup` chains `ShowFirstRunWizardIfNeededAsync` → `TryAutoConnectAsync` via ContinueWith UI scheduler. Re-aplica tema tras wizard close.

Tests **401 verdes** (+10 FirstRunWizardViewModelTests: state machine, persistencia, skip, finish con trim, save error, labels). 326 Core + 75 App.

### 2026-05-22 (final) — Sprint F "UI polish a fondo"

User reporta: "se ve muy verde, auditoría resultados como mensajes de error, modo oscuro texto oscuro y azul oscuro no se ve". 4 commits.

Bugs visuales identificados + fix:
- **DataGrid/ListView/ListBox texto invisible en dark theme** — legacy WPF defaults Foreground=Black hardcoded; wpf-ui ThemesDictionary NO restilea estos controles. Fix: global Styles (no x:Key) en App.xaml para DataGrid + DataGridCell + DataGridRow + DataGridColumnHeader + ListView + ListViewItem + GridViewColumnHeader + ListBox, Foreground=`TextFillColorPrimaryBrush`. Normaliza AlternatingRowBackground (`SubtleFillColorTertiary`), header bg (`SubtleFillColorSecondary`), bottom border 1px stroke, RowHeight 32, transparent rows, no grid lines. Esto fix la tabla Hallazgos en AuditView + tabla jsonl en AuditLogView.
- **Severity pills (AuditView + DashboardView) invisible en light theme** — count text Foreground=`#FFFFFF` hardcoded; label usaba `BrushSemantic*MutedText` (#FFEAEA/#FFF8E5/#E6F0FF) diseñado solo para dark. Fix: rediseño pills como Border + StackPanel. Border: Soft bg + strong border 1px. Label "ERROR" (Bold 11px) Foreground=`BrushSemanticError` (theme-independent strong color). Count (Bold 14-16px) Foreground=`TextFillColorPrimaryBrush` (theme-aware blanco/negro). Borrados los brushes MutedText sin uso.

Polish UI adicional:
- `ux(splitter)` GridSplitter invisible — custom ControlTemplate con 40x3 pill central `ControlStrokeColorDefault` Opacity 0.7. Cursor SizeNS hover.
- Card hover state: `CardHover` style toggles `ControlFillColorSecondary` bg + `ControlStrokeColorSecondary` border on IsMouseOver. `MetricCard` border on AccentFillColorTertiary on hover.
- Page transition: ContentControl ControlTemplate envuelve ContentPresenter en Border con Loaded EventTrigger storyboard: Opacity 0→1 + TranslateTransform Y 6→0, 180ms CubicEase EaseOut. Cada nav change re-materializa view → fade-in.
- `PageSubtitleText` ahora usa `TextFillColorSecondaryBrush` (theme-aware) + Margin 20→16 tighter.
- ProgressBar global: custom ControlTemplate CornerRadius=3 (track + indicator). Background `ControlFillColorTertiary`. MinHeight 6. Look Fluent moderno. License cards utilization bars beneficiados.

Tests **391 verdes** (sin tests rotos por refactor styles). Build clean 7 projects 0 errors.

### 2026-05-22 (noche) — Sesión "Sprint C completion + D + E"

Continúa autonomous tras save. **6 commits**. Tests **326 → 391** (+65 App.Tests). Build clean 7 projects 0 errors.

Commits:
- `test(app)` Sprint C completion: añade `App.Tests` csproj a `src/Grex365.slnx` + escribe `UserDetailsViewModelTests` (16 tests). Patrón Harness con Mock<IUsersService> strict + TestDialog/Clipboard/UiLogSink/UserDetailsHost + WaitForAsync helper para fire-and-forget LoadAsync triggered by Host.OpenRequested. Cubre OpenRequested→populate / userNotFound / CloseRequested→reset / Toggle confirm-no/yes / RemoveLicense null/yes / AssignSelectedSku noUser/noSku/valid / ResetPassword confirm-no/yes (clipboard+show) / RevokeSessions / RemoveAllLicenses confirm-no/yes / Close command.
- `test(app)` Sprint D parte 1: OffboardingViewModelTests (6) + UsersViewModelTests (15). RBAC denied / confirm-no / confirm-yes / DisableEnable / AssignLicense seat-available/no-seats / RemoveLicenses zero/positive licenses. Total App.Tests 37.
- `test(app)` Sprint D parte 2: GroupsViewModelTests (5) + OnboardingViewModelTests (5). RemoveMember confirm paths + service-throws / OnboardingRun confirm-no/yes con verificación trim UPN + uppercase usageLocation + GroupIdentifiers split. AddSelectedSku dedup + RemoveSku. Total 47.
- `test(app)` Sprint D parte 3: MailboxRulesViewModelTests (9). ApplyForwarding empty/RBAC/confirm-no/confirm-yes con trim. ClearForwarding confirm-yes resets fields + display. RemoveCalendarPermission noSelection/confirm-yes/no. Cierra cobertura parity de los 6 VMs refactorizados. Total 56 App.Tests.
- `feat(ps-console)` Sprint E: nuevo módulo "Consola PS" en nav Herramientas. `PsConsoleViewModel` reusa `IPowerShellRunner` existente (RunspacePool-backed) — scripts ejecutan en mismo contexto que app (Graph/EXO en scope). Input multi-line + OutputText accumulator PS>-prefixed + Cancel + Clear + History navegable (Up/Down, MaxHistory=50, dedupe consecutivos) + IsBusy + ProgressRing. View con header CommandPrompt glyph + warning banner ("sin sandbox") + KeyBindings Ctrl+Enter→Run, Esc→Cancel. PsConsoleViewModelTests (9) con Mock<IPowerShellRunner>: empty input no-op / valid script appends output+history / errors [ERROR] markers / runner throws [EXCEPCION]+log / cancel [CANCELADO] / Clear empties / history dedup / Prev navigate-back stops oldest / Next advances + wraps past-end empty.
- `fix(nav)` Inyecta glyph U+E756 (CommandPrompt) en nav item "Consola PS" via Python utf-8 r+w (PUA chars unreliable via Edit transport — patrón usado previamente en fix(nav) glyphs faltantes).

**Estado final sesión:** 391 tests (326 Core + 65 App) verdes. Build 7 projects 0 errors. 6 VMs refactorizados (UserDetails/Users/Groups/Offboarding/Onboarding/MailboxRules) con cobertura tests parity completa. Nuevo módulo Consola PS funcional + testeado. Plantamiento §6 backlog "Terminal PowerShell embebido" CERRADO via PS Console module (reusa runner existente; menos invasivo que `EasyWindowsTerminalControl` integration).

### 2026-05-22 (tarde) — Sesión "polish + architectural cleanup"

Sprints A + B completos, Sprint C en progreso (refactor done, tests pendientes). 4 commits. Build clean, 326 tests siguen verdes.

Commits:
- `ux(nav+drawer)` Acentos consistentes en nav titles ("Conexión"/"Auditoría"/"Reglas de buzón") + dict de migración para `LastSelectedNavigation` persistido. Sets `RequiresGraphTitles`/`RequiresExchangeTitles` + `SyncAuditBadge` + `DashboardView.CommandParameter` + strings CertWizard actualizados. Drawer user lateral con slide animation (TranslateTransform X 460↔0, 220ms EaseOut entrada / 180ms EaseIn salida) + `IsHitTestVisible` bound a `UserDrawerOpen` evita hits off-screen. Reemplaza Visibility toggle hard.
- `docs` Sync `ROADMAP.md` + `MIGRATION.md` + `ARCHITECTURE.md` al estado 2026-05-22 (7 días stale). ROADMAP gana snapshot header + H1/H2 sub-items refrescados (1.2.6/1.2.7/1.2.8/1.2.9/1.3.7/1.4.7/1.6.3/1.6.4 marcados ✅). H3 colapsado a tabla single con todos features ported. H4 marcado mostly-done. H5 con MSIX scaffold ✅ + assets/firma pendientes. Decision tracker D4/D9 resolved, D1/D2 marcados de-facto. "What next" actualizado al backlog real (MSIX assets, QA load, test gap, terminal embebido, first-run wizard). MIGRATION añade sección con 13 security audits nuevos (no estaban en legacy PS). ARCHITECTURE mueve Plugin/AppInsights/MSIX de "NOT taken" a nueva sección "Stack decisions later reversed".
- `feat(polish)` Bundle SemVer + About + Window restore + Log export + global shortcuts. `Directory.Build.props` raíz con `<Version>0.2.0-alpha</Version>` (single source). `App.AppVersion` static lee `AssemblyInformationalVersionAttribute`. Sidebar header + status bar bindados via `{x:Static local:App.AppVersion}`. `AboutWindow.FluentWindow` con name + version + runtime + data dir + "Abrir carpeta" (Process.Start). Wired en DI + `MainViewModel.OpenAboutCommand` + botón "Acerca de" en status bar + F1 keybinding. `UserPreferences` gana `WindowWidth/Height/Left/Top/Maximized`. `MainWindow.xaml.cs` Loaded restaura con guard `IsOnScreen` (VirtualScreen check) + Closing persiste; skip size/pos si maximized. `MainViewModel.ExportLogCommand` abre `SaveFileDialog` (.txt default, .csv option), itera `LogView` (respeta filtros severity) y escribe UTF-8 timestamped + CSV escape. Botón "Exportar" junto a "Limpiar". KeyBindings nuevos: Ctrl+, → Settings, F1 → About, Ctrl+L → ToggleLogPanel. Tooltips actualizados.
- `refactor` Extract `IDialogService` + `IClipboardService` a `Grex365.Core.Abstractions`. WPF impls (`WpfDialogService` mapea `DialogIcon` → `MessageBoxImage`; `WpfClipboardService` swallows shell/RDP throws). Registrados Singleton en DI. 6 VMs refactorizadas (`UserDetails`/`Users`/`Groups`/`Offboarding`/`Onboarding`/`MailboxRules`): 17 `MessageBox.Show` + 1 `Clipboard.SetText` extraídos. Cero direct WPF refs en VMs (verificado grep). Restaura regla arquitectónica de `ARCHITECTURE.md` (VMs UI-agnostic). Scaffold `tests/Grex365.App.Tests/Grex365.App.Tests.csproj` (net10.0-windows + WPF + ref App + xUnit/Moq/FluentAssertions) + `TestFakes.cs` (`TestDialogService`/`TestClipboardService`/`TestUiLogSink`/`TestUserDetailsHost`). **NO incluido en `src/Grex365.slnx` aún** — próxima sesión añade + escribe tests.

### Próximo paso (continúa Sprint C)

1. Añadir `<Project Path="../tests/Grex365.App.Tests/Grex365.App.Tests.csproj" />` a `src/Grex365.slnx`.
2. Escribir `UserDetailsViewModelTests`:
   - LoadAsync popula User + Memberships + AssignedLicenses cuando `OpenRequested`
   - LoadAsync user-not-found set `StatusMessage = "Usuario no encontrado."`
   - Reset clears collections + state cuando `CloseRequested`
   - ToggleAccount con `ConfirmResult=false` → no calls al service
   - ToggleAccount con `ConfirmResult=true` → `SetAccountEnabledAsync(uid, !User.AccountEnabled, ...)` + reload
   - RemoveLicense con `null row` → no-op
   - RemoveLicense con row + confirm yes → `RemoveLicenseAsync(uid, skuId)` + reload
   - AssignSelectedSku con no User → no-op; con sku + user → `AssignLicenseAsync` + reload
   - ResetPassword copia a clipboard + show dialog
   - RevokeSessions con confirm yes → `RevokeSignInSessionsAsync`
3. Tests adicionales (opcional Sprint D): `GroupsViewModelTests`, `OffboardingViewModelTests`.
4. Verificar `dotnet test src/Grex365.slnx` corre ambos test projects.

### 2026-05-22 — Sesión "security audits sprint"

Tests **214 → 301** (+87). Siete auditorías de seguridad nuevas + refactor notify.

Commits:
- `feat(audit)` Conditional Access policies audit: `CaPolicyAnalyzer` puro + `RunConditionalAccessAuditAsync` consume `/identity/conditionalAccess/policies`. Categorías: disabled (INFO), report-only stale ≥30d (WARN), report-only fresh (INFO), enabled sin builtInControls (ERROR), enabled con controls débiles (WARN), enabled con `All` users sin exclusiones (WARN), enabled sin user scope (ERROR), tenant sin policies (ERROR). Suma `Policy.Read.All` al App Reg auto-create.
- `feat(audit)` Privileged role assignments audit: `PrivilegedRoleAuditAnalyzer` puro + `RunPrivilegedRolesAuditAsync` enumera `/directoryRoles` + members en paralelo (8x). Categorías: guests con role admin (ERROR), cuentas deshabilitadas con role (ERROR), service principals con role (INFO), 0 Global Admins (ERROR), 1 Global Admin (WARN — sin backup), >5 Global Admins (WARN — sprawl). Identifica GA via templateId `62e90394-69f5-4237-9190-012177145e10`.
- `feat(audit)` App credentials expiry audit: `AppCredentialAuditAnalyzer` puro + `RunAppCredentialsAuditAsync` itera `/applications` y agrega `passwordCredentials` + `keyCredentials`. Severidades: expired (ERROR), expirando ≤30d (WARN), long-lived >2y (INFO). Suma `Application.Read.All` al App Reg auto-create.
- `feat(audit)` Tenant defaults audit: `TenantDefaultsAnalyzer` puro + `RunTenantDefaultsAuditAsync` consume `/policies/authorizationPolicy` + `/policies/identitySecurityDefaultsEnforcementPolicy`. Categorías: users-can-create-apps (WARN), users-can-create-tenants (WARN), email-verified self-join (WARN), allowInvitesFrom=everyone (WARN), SSPR disabled (INFO), email-based subscriptions (INFO), SecurityDefaults on (INFO).
- `feat(audit)` OAuth consent grants audit: `OAuthGrantAnalyzer` puro + `RunOAuthGrantsAuditAsync` enumera `/oauth2PermissionGrants` y resuelve nombres de clientes/recursos via `/servicePrincipals` (8x paralelo). Severidades: AllPrincipals + scope alto-riesgo (ERROR — admin consent peligroso), Principal + scope alto-riesgo (WARN — posible phishing OAuth). Set alto-riesgo: Mail.*/Files.*/Sites.*/Directory.*/User.*/Group.*/Calendars.*/Contacts.*/full_access_as_user/Notes.ReadWrite.All. Cubierto por `Directory.ReadWrite.All` existente.
- `feat(audit)` Transport rules security audit: `TransportRuleAuditAnalyzer` puro + `ScanTransportRulesAsync` en `ExoForwardingAuditService`. Severidades: enabled rule con ForwardTo/BlindCopyTo/RedirectMessageTo a destinatarios externos (ERROR — exfil), outbound connector custom (INFO — verificar hybrid), DeleteMessage broad-scope (WARN), modo Audit (INFO — no enforcement), rules disabled con nombre que contiene `anti/spam/phish/dlp/quarantine/...` (WARN). Match dominios case-insensitive con `Get-AcceptedDomain`.
- `feat(audit)` Shared mailbox sign-in audit: `SharedMailboxSignInAnalyzer` puro + `ScanSharedMailboxSignInAsync` en `ExoForwardingAuditService`. Get-Mailbox -RecipientTypeDetails SharedMailbox + Get-User por UPN para extraer AccountDisabled. Flagea shared mailboxes con `AccountDisabled=false` (WARN — sign-in habilitado, vector password attack sin MFA). Tolerante a Get-User fallo (INFO unknown).
- `ux(audit)` orden findings por severidad (ERROR > WARN > INFO) en AuditView. Helper `AddFindingsSorted` en VM aplicado a todos los `Run*Audit`. Nueva columna Severidad coloreada via `AuditSeverityToBrushConverter` (rojo/ámbar/azul). Hace mucho más legible el listado con 11 audits coexistentes.
- `ux(audit)` filtros + búsqueda + summary badges en AuditView. CollectionView sobre Findings con predicate severity + text-search (Category/Identity/Detail). Checkboxes ERROR/WARN/INFO toggle visibility. Pills coloreadas con counts dinámicos vía `RecomputeCounts()` hooked en CollectionChanged.
- `ux(audit)` glyphs Segoe Fluent en columna Severidad (ErrorBadge U+EA39, Warning U+E7BA, Info U+E946). `AuditSeverityToGlyphConverter` separado del brush converter. Refuerza visualmente sin depender solo de color (accesibilidad).
- `feat(dashboard)` card "Última auditoría" — `IAuditFindingsStore` (singleton) recibe Update tras cada Run*Audit; Dashboard subscribe vía PropertyChanged, muestra nombre + timestamp + pills ERROR/WARN/INFO con botón "Abrir Auditoría". Hace visible el resultado de seguridad desde el landing.
- `ux(sidebar)` badge rojo con count de ERROR findings en la nav item "Auditoría". `NavigationItem` gana `ErrorBadge` + `HasErrorBadge`; `MainViewModel` subscribe a `IAuditFindingsStore` y empuja `_auditStore.ErrorCount` al item correspondiente. Visible globalmente desde cualquier módulo.
- `ux(audit)` empty state con glyph FavoritesList U+E930 + mensaje "Sin hallazgos todavía" cuando `Findings.Count == 0`. `CountToVisibilityConverter` extendido con parámetro `invert` para mostrar/ocultar inverso. Mejora primera impresión antes de lanzar primer audit.
- `ux(audit)` keyboard shortcuts en AuditView via InputBindings: `Ctrl+R` lanza identidad+grupos, `Esc` cancela, `Ctrl+E` exporta CSV. Tooltips actualizados. Export ahora itera `FindingsView` (respeta filtros ERROR/WARN/INFO + búsqueda) en lugar de toda la colección.
- `fix(users)` quita `signInActivity` del `$select` en `GraphUsersService.SearchAsync`/`GetByIdAsync` — provocaba `AuditLog.Read.All required` cuando se buscaba un usuario. Audit identity sí lo sigue usando (sólo donde es imprescindible para stale-detection).
- `ux` Enter dispara búsqueda en Users / Groups / SharedMailbox / MailboxRules / DNS check (KeyBinding en TextBox.InputBindings).
- `ux` log panel inferior **colapsable + oculto por defecto**. `UserPreferences.LogPanelVisible` persiste el estado. `MainViewModel.ToggleLogPanelCommand` + botón "Log" en status bar. RowDefinitions usan Style+DataTrigger para colapsar splitter y border cuando el panel está oculto.
- `feat(groups)` elimina toggle manual "DL (Exchange)". `BulkGroupRow.GroupType` (default `M365`) + `BulkGroupRowPreprocessor` ahora detecta tipo desde columnas `GroupType`/`Type`/`Kind` del CSV (alias: m365/microsoft 365/unified/dl/distribution/exchange) con forward-fill como el nombre. `BulkCreateFromCsvAsync` divide rows por tipo y despacha a `_groups.CreateM365GroupsFromRowsAsync` y/o `_dls.CreateFromRowsAsync` según corresponda.
- `feat(tenant-health)` rediseño completo de cards de licencias estilo portal M365. `SkuCatalog` (mapa curado de SKUs MS) resuelve `SkuPartNumber` → `FriendlyName` + `LicenseCategory` (Enterprise/Business/Frontline/Security/AppOrAddOn/Other) + `Priority`. `LicenseCard` VM presenter calcula utilización (Low/Medium/High/Critical) + colored brush. `UtilizationToBrushConverter` mapea a verde/azul/ámbar/rojo. Vista usa `CollectionViewSource` con `GroupDescription` por categoría y `SortDescription` por priority, `WrapPanel` muestra cards de 280px con header de categoría. 13 tests del catálogo (alias, fallback humanize, case-insensitive, priority order).
- `ux(sidebar)` reestructura nav agrupado por categorías estilo CIPP: Tenant (Dashboard/Conexión/Salud), Identidad (Usuarios/Grupos/Onboarding/Offboarding), Mail (Buzones/Reglas/Mail flow), Seguridad (Auditoría/Audit log), Herramientas (Cert Wizard/DNS), Plugins (módulos externos). `NavigationItem.Category` + `NavigationItemsView` (CollectionViewSource) con `PropertyGroupDescription`. `ListBox.GroupStyle.HeaderTemplate` muestra header pequeño semibold opacity 0.55.
- `feat(user-portal)` mini-portal lateral estilo M365 admin para usuarios: `IUserDetailsHost` (singleton state holder con eventos OpenRequested/CloseRequested + INotifyPropertyChanged). `UserDetailsViewModel` carga perfil + group memberships del usuario y expone comandos rápidos (Toggle account, Remove all licenses, Close). `UserDetailsView` UserControl con header + status + quick-action bar + identity card + memberships list. Doble-click en UsersView ListBox o en Groups members list dispara `host.RequestOpen(id)`. MainWindow muestra overlay Border 460px de ancho a la derecha con DropShadow, animado vía Visibility (mejor: futura animación). `BoolToOnOffConverter` extendido con parámetro `'TrueLabel/FalseLabel'` para reusarlo.
- `fix(mail-flow)` corrige error `Method invocation failed because [System.Net.Http.HttpResponseMessage] does not contain a method named 'GetResponseHeader'` en Mail Flow. Bug conocido ExchangeOnlineManagement 3.x + PS7: ForEach-Object detrás de Get-TransportRule fuerza lazy property fetch que toca path interno con HttpWebResponse. Workaround: usar `Select-Object @{N=...;E={...}}` directamente para forzar evaluación eager. Mismo `param()` declarado para que AddParameter binde si añadimos parámetros futuros.
- `fix(audit-identity)` resilient si falta `AuditLog.Read.All`: try/catch específico sobre Users.GetAsync. En fallo de permiso, retry sin `signInActivity` en `$select` + WARN finding `"Permiso faltante: AuditLog.Read.All"` con guía para conceder admin consent. Stale-user detection se degrada gracefully en vez de abortar todo el audit.
- `ux(logos)` headers consistentes con glyph Segoe Fluent en MailFlowRulesView (U+E715 Mail), MailboxRulesView (U+E71B MailReply) y AuditLogView (U+E7C3 Page2). Mismo patrón que AuditView/UsersView.
- `feat(user-portal)` licencias granulares en mini-portal: `IUsersService` gana `GetAssignedLicensesAsync(userId)` + `RemoveLicenseAsync(userId, skuId)`. UserDetailsViewModel ahora carga licencias asignadas resueltas vía SkuCatalog (FriendlyName + Category) en `AssignedLicenses` y SKUs disponibles del tenant en `AssignableSkus`. UserDetailsView muestra card con lista por licencia (FriendlyName + SkuPart + Category + botón Quitar individual) + ComboBox con SKUs disponibles (PartNumber + asientos libres) + botón Asignar.
- `fix(drawer)` cerrar drawer + botones licencias no respondían: drawer Border estaba `Grid.RowSpan=3` cubriendo la fila TitleBar de la FluentWindow, que interceptaba los clicks en la parte superior (close button). Reposicionado a `Grid.Row=1` (sólo fila de contenido) + `Panel.ZIndex=999`. Eliminado `DropShadowEffect` (afectaba hit-testing). Background a `SolidBackgroundFillColorBaseBrush` (opaco). UserControl `x:Name="Root"` + binding `{Binding ElementName=Root, Path=DataContext.RemoveLicenseCommand}` reemplaza `RelativeSource AncestorType=UserControl` que fallaba dentro del ItemsControl.
- `feat(autoconnect)` conexión automática al arrancar si certificado + config válidos: nuevo `TryAutoConnectAsync()` en `App.OnStartup` después de `MainWindow.Show()`. Carga cert config, valida con `ICertValidator`, conecta Graph, fuerza `TenantLock`, conecta EXO (best-effort, log Warn si falla). Errores no rompen la app — el usuario puede conectarse manualmente en Conexión.
- `fix(exo)` workaround bug `HttpResponseMessage does not contain GetResponseHeader` que afectaba Get-TransportRule / Get-Mailbox / Get-InboxRule en EXO 3.x + PS7. Set `$env:DISABLE_REST_API_USE_BY_DEFAULT = "true"` antes de `Connect-ExchangeOnline` fuerza path legacy que no tiene el bug. Aplicado en `ExchangeConnection.ConnectByCertificateAsync` + redundante en `MailFlowRulesService`.
- `feat(user-portal)` quick actions extendidas: Reset password (genera contraseña 16 chars criptográficamente segura con cada categoría: uppercase/lowercase/digits/symbols, sin caracteres ambiguos, copia al portapapeles + MessageBox) y Revoke sessions (POST /users/{id}/revokeSignInSessions). `ResetPasswordAsync` y `RevokeSignInSessionsAsync` en `IUsersService` y `GraphUsersService`.
- `ux(empty-states)` UsersView + GroupsView con placeholder cuando no hay resultados (icono Contact U+E77B / People U+E716 + texto guía "Busca... — pulsa Enter"). Listas con avatares con iniciales (`InitialsConverter`) — círculo 32px Background AccentFill para usuarios, square con CornerRadius=6 para grupos, badge GroupKind como pill.
- `ux(connect)` header con logo Globe U+E703, nuevo card "Certificado" muestra App ID, Thumbprint, Expira, status con pill coloreada válido/inválido. `ConnectViewModel.LoadCertInfoAsync` se ejecuta en el constructor para mostrar info inmediatamente al abrir la vista.
- `ux(logos)` headers consistentes en TODAS las views: Dashboard (Home U+E80F), TenantHealth (Health U+E9D9), Audit (Shield U+E9D5), SharedMailbox (Mail U+E715), Onboarding (AddFriend U+E8FA), Offboarding (BlockContact U+E8F8), CertWizard (Lock U+E72E), DomainCheck (Globe U+E774). Patrón StackPanel Orientation=Horizontal con glyph FontSize=22 + título FontSize=24-28.
- `ux(audit-dark)` contraste pills + checkboxes en dark theme: alpha background 0x22 → 0x55 + BorderThickness=1 + Foreground brighter shades (`#FFEAEA/#FFF8E5/#E6F0FF` con text negro `#FFFFFF` para count). `AuditSeverityToBrushConverter` ahora usa tonos brighter `#F87171/#FBBF24/#60A5FA` que funcionan bien en dark Y light. Checkboxes filter FontWeight=SemiBold.
- `perf(audit)` paraleliza `RunIdentityAuditAsync` + `RunGroupsAuditAsync` con `Task.WhenAll` en `RunAsync` VM. Antes secuencial — ahora corren simultáneamente.
- `feat(audit-scenarios)` 3 escenarios pre-canned con filtro automático de findings:
  - **Cuentas en riesgo (deshab+licencia)**: identity + groups en paralelo, filtra findings con keywords disabled/license/stale/inactive
  - **Acceso privilegiado**: privileged roles + MFA + CA en paralelo
  - **Higiene mail (BEC)**: forwarding externo + inbox rules + transport rules + shared mailbox sign-in
- `ux(audit-layout)` reorganiza botones en 4 secciones con headers: ESCENARIOS RÁPIDOS / IDENTIDAD / SEGURIDAD-TENANT / MAIL-EXO. Cada sección en Border CornerRadius=8 con label small caps opacity 0.55. Pasa de 2 filas planas confusas a 4 grupos claros por dominio.
- `ux(powertoys-style)` adopta patrón **PowerToys Settings** (CommunityToolkit SettingsCard portado a WPF): nuevos estilos `SettingsCard` (Border CornerRadius=6, MinHeight=64, hover state), `SettingsCardIcon` (TextBlock 32px ancho, glyph FontFamily), `SettingsCardTitle` (14px SemiBold), `SettingsCardDescription` (12px opacity 0.6), `SettingsGroupHeader` (11px small-caps opacity 0.55). Aplicado a license cards (single-row con icon + name + progress bar + % en lugar de chunky cards de 280px) y Connect cert info (4 rows: Estado / App ID / Thumbprint / Expiración). Sidebar nav con accent left rail 3px estilo Files en lugar de fondo accent completo — selección más sutil y visual estilo Windows Settings.
- `refactor(audit-vm)` extrae `NotifyAllCommands()` helper en AuditViewModel — elimina ~80 líneas de notify chains repetidas y evita bugs cuando se añade un nuevo command.

Plantamiento status: Fase 6 sigue **6/8 hechos**; backlog seguridad amplía cobertura M365 baseline. Toolbar fila 1 ahora con 8 audits Graph (Identidad+grupos, MFA, CA policies, Privileged roles, App credentials, Tenant defaults, OAuth grants, Actividad grupos) y fila 2 con 4 audits EXO (Forwarding externo, Inbox rules, Transport rules, Shared mailbox sign-in).

---

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
- `feat(rbac)` extiende gate a Users/Groups/SharedMailbox y MailboxRules (OOO/forwarding/calendar perms) — patrón uniform `RequireAuthorizedAsync(context)`. Constructive ops (Enable, AssignLicense, AddMembers, Onboarding, CertWizard) sin gate.
- `feat(audit)` detector forwarding externo (security signal): `MailboxForwardingAnalyzer` puro + `ExoForwardingAuditService` que combina `Get-AcceptedDomain` con `Get-Mailbox` para flagear `ForwardingSmtpAddress` hacia dominios fuera del tenant.
- `feat(audit)` scanner de inbox rules (BEC indicator): `InboxRuleAnalyzer` puro detecta DeleteMessage, MoveToFolder en set sospechoso (Deleted/Junk/RSS/Archive…), Forward/Redirect externo. Keywords BEC en EN+ES elevan a WARN. `ScanInboxRulesAsync(maxMailboxes)` itera Get-InboxRule por buzón con tope configurable.
- `feat(audit)` cobertura MFA: `MfaCoverageAnalyzer` puro + `RunMfaCoverageAuditAsync` consume `/reports/authenticationMethods/userRegistrationDetails`. Admin sin MFA = ERROR crítico, member sin MFA = WARN, guest sin MFA = INFO. Summary con admin coverage %.

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
- [x] **Permisos por rol (RBAC)** — `IRbacGuard` + `IMembershipChecker` (`GraphMembershipChecker` envuelve `/me/checkMemberGroups`). Settings textbox `AuthorizationGroupId`. Gateado en TODAS las acciones destructivas: Offboarding, Users (Disable/RemoveLicenses/BulkCsv), SharedMailbox (Convert/ApplyPermission/BulkCsv), Groups (RemoveMember/BulkCreate). Acciones constructivas (Enable, AssignLicense, AddMembers) sin gate. Cache invalidado en Disconnect.
- [ ] Documentación técnica interna (arquitectura, manual operación)
- [ ] QA escenarios reales (100+ ops simultáneas)

---

## Backlog funcional (no asociado a una fase concreta)

### Features útiles pendientes
- [x] **Auth interactivo Graph sin cert preexistente** — device-code via cliente público de Azure CLI (`ConnectByDeviceCodeAsync`); valida acceso real con `Me` + `Organization` antes de marcar conectado
- [x] **Mail flow rules viewer** — nuevo modulo de navegacion ("Mail flow") que lista `Get-TransportRule` de EXO (Name/State/Priority/Mode/Description) con filtro libre; gated por RequiresExchange
- [x] **Auditoría: grupos sin actividad reciente** — `RunGroupActivityAuditAsync` consume `/reports/getOffice365GroupsActivityDetail` (period D7/D30/D90/D180), UI con NumberBox umbral, exporta junto al resto de findings
- [x] **Auditoría: forwarding externo (security)** — `ExoForwardingAuditService` combina `Get-AcceptedDomain` + `Get-Mailbox -ResultSize Unlimited | Where ForwardingSmtpAddress` para flagear forward hacia dominios fuera del tenant (vector típico de phishing/exfiltración). Botón "Forwarding externo" en AuditView
- [x] **Auditoría: inbox rules (BEC indicator)** — `InboxRuleAnalyzer` puro + `ScanInboxRulesAsync(maxMailboxes)`. Categorías: delete, hide (move-to Deleted/Junk/RSS/Archive…), external forward/redirect. Keywords BEC EN+ES (invoice/factura/wire/payment/password…) elevan a WARN. Botón "Inbox rules" + NumberBox tope buzones en AuditView
- [x] **Auditoría: MFA coverage** — `MfaCoverageAnalyzer` puro + `RunMfaCoverageAuditAsync` consume `/reports/authenticationMethods/userRegistrationDetails`. Admin sin MFA = ERROR crítico, member = WARN, guest = INFO. Status muestra cobertura admin %. Requiere `Reports.Read.All`
- [x] **Auditoría: Conditional Access policies** — `CaPolicyAnalyzer` puro + `RunConditionalAccessAuditAsync` consume `/identity/conditionalAccess/policies`. Detecta policies disabled (INFO), report-only stale ≥30d (WARN), enabled sin builtInControls (ERROR), enabled con controls débiles sin MFA/compliantDevice/block (WARN), enabled con includeUsers=All sin exclusiones (WARN), enabled sin user scope (ERROR), y tenant sin policies (ERROR). Requiere `Policy.Read.All`
- [x] **Auditoría: Privileged role assignments** — `PrivilegedRoleAuditAnalyzer` puro + `RunPrivilegedRolesAuditAsync` enumera `/directoryRoles` y members en paralelo. Detecta guests con role admin (ERROR), cuentas deshabilitadas con role (ERROR), service principals con role (INFO), 0 Global Admins (ERROR), 1 GA (WARN — sin backup), >5 GAs (WARN — sprawl). Cubierto por `Directory.ReadWrite.All` existente
- [x] **Auditoría: App credentials expiry** — `AppCredentialAuditAnalyzer` puro + `RunAppCredentialsAuditAsync` itera `/applications` y agrega `passwordCredentials` + `keyCredentials`. Severidades: expired (ERROR), expirando ≤30d (WARN), long-lived >2y (INFO). Requiere `Application.Read.All`
- [x] **Auditoría: Tenant defaults / authorization policy** — `TenantDefaultsAnalyzer` puro + `RunTenantDefaultsAuditAsync` consume `/policies/authorizationPolicy` + `/policies/identitySecurityDefaultsEnforcementPolicy`. Detecta users-can-create-apps (WARN), users-can-create-tenants (WARN), allowInvitesFrom=everyone (WARN), email-verified self-join (WARN), SSPR disabled (INFO), SecurityDefaults on (INFO). Requiere `Policy.Read.All`
- [x] **Auditoría: OAuth consent grants** — `OAuthGrantAnalyzer` puro + `RunOAuthGrantsAuditAsync` enumera `/oauth2PermissionGrants` y resuelve nombres de clientes/recursos via `/servicePrincipals`. AllPrincipals + scope alto-riesgo (ERROR — admin consent), Principal + scope alto-riesgo (WARN — posible phishing OAuth). Set alto-riesgo cubre Mail/Files/Sites/Directory/User/Group/Calendars/Contacts/Notes + full_access_as_user. Cubierto por `Directory.ReadWrite.All`
- [x] **Auditoría: Transport rules** — `TransportRuleAuditAnalyzer` puro + `ScanTransportRulesAsync` en `ExoForwardingAuditService`. Detecta rules enabled con forwarding/BCC/redirect a destinatarios externos (ERROR — vector exfil), outbound connectors custom (INFO), DeleteMessage broad-scope (WARN), modo Audit (INFO), y rules disabled cuyo nombre incluye keywords de seguridad anti/spam/phish/dlp/quarantine (WARN). Match dominios contra `Get-AcceptedDomain`
- [x] **Auditoría: Shared mailbox sign-in** — `SharedMailboxSignInAnalyzer` puro + `ScanSharedMailboxSignInAsync` en `ExoForwardingAuditService`. Get-Mailbox SharedMailbox + Get-User por UPN. Flagea shared boxes con `AccountDisabled=false` (WARN — sign-in habilitado, vector password attack). Tolerante a Get-User fallo (INFO unknown)
- [x] **Cert export PFX con password** — `ICertificateGenerator.ExportPfx`, panel "Exportar PFX" en CertWizardView con PasswordBox + tests de validacion (no encontrado, password vacio, etc.)
- [x] **Auto-create App Registration vía Graph** — `GraphAppRegistrationService.CreateAndConfigureAsync` aplica todos los AppRoles (User/Group/GroupMember/Organization/AuditLog/Directory + Exchange.ManageAsApp + Reports.Read.All), sube cert como `KeyCredential`, crea ServicePrincipal y devuelve admin-consent URL clickable. Reemplaza los 29 pasos manuales del legacy.
- [x] **Auto-install módulo EXO** — `ExchangeConnection.InstallModuleAsync` lanza `pwsh.exe` externo (Start-Process) para esquivar el ACL de WindowsApps que niega `Microsoft.PackageManagement.dll` en runspaces embebidos. UI muestra estado del módulo + botones Comprobar/Instalar.

### Polish UI
- [ ] Terminal PowerShell embebido (`EasyWindowsTerminalControl`)
- [x] **Theme toggle desde sidebar** — botón "Tema" junto a "Ajustes" persiste y aplica al instante
- [x] **Disable nav items cuando Graph/Exchange desconectado** — `NavigationItem.RequiresGraph/RequiresExchange`, `MainViewModel.UpdateNavEnabledStates` reactivo al `ConnectionStateMonitor`

---

## Tests (326 passing)

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
| BulkGroupRowPreprocessor | 25 | Forward-fill, skip orphans, trim, IsEmail theory, NormalizeType (M365/DL aliases case-insensitive), GroupType detection from CSV column, forward-fill type |
| OnboardingValidator | 16 | UPN/password/usage/mail-nickname validation + derive |
| MailboxRulesValidator | 15 | OOO state transitions, date ranges, forwarding SMTP shape |
| BulkUserActionParser | 17 | enable/disable/remove-licenses + assign:&lt;SKU&gt; parse + lookup |
| PluginLoader | 4 | empty dir / corrupt dll / whitespace path |
| FileAuditLog | 4 | roundtrip / append-jsonl / missing-month / concurrent-writes |
| MetricsAggregator | 6 | Totales, error rate, last 24h, top sources, recientes |
| CertificateGenerator | 4 | Self-signed + ExportPfx |
| NullTelemetry | 3 | IsEnabled=false + no-throw para TrackEvent/Exception/Flush |
| RbacGuard | 8 | sin-grupo short-circuit, whitespace, miembro/no, checker-throws, cache, Invalidate, trim |
| MailboxForwardingAnalyzer | 10 | sin-fwd, interno, externo flagged, SMTP prefix, case-insensitive, UPN vacío, malformado, ForwardingAddress no flagged, trailing dot, multi-rows |
| InboxRuleAnalyzer | 13 | disabled skip, UPN vacío, delete plain/keyword, move-deleted/regular/RSS, forward externo/interno, redirect externo, SMTP en brackets, 3 findings combinados, keyword español |
| MfaCoverageAnalyzer | 9 | empty, admin sin/con MFA, member sin MFA, guest sin MFA, UPN vacío, IsAdmin precedence, mixed pop, capable=false |
| CaPolicyAnalyzer | 16 | empty (ERROR), enabled strong OK, disabled, report-only stale/fresh, sin controls, weak controls, All sin exclusions, exclusions valid, sin user scope, empty name, unknown state, compliantDevice/block strong, case-insensitive, mixed |
| PrivilegedRoleAuditAnalyzer | 14 | empty (no GA), 1 GA (backup warn), 2 GAs OK, >5 GAs (sprawl), guest admin (ERROR), disabled admin (ERROR), SP admin (INFO), disabled-SP no false positive, multi-roles same user, guest in 2 roles, empty memberId, no-GA-other-role, template case-insensitive, identity from displayName |
| AppCredentialAuditAnalyzer | 11 | empty, expired (ERROR), expiring soon (WARN), not-soon OK, long-lived (INFO), null endDate, empty appId, credential name in detail, mixed counts, Key type, threshold exact |
| TenantDefaultsAnalyzer | 10 | safe baseline, create-apps (WARN), create-tenants (WARN), email-verified join, invitesFrom=everyone, invitesFrom=adminsOnly OK, SSPR disabled, email-based subs, SecurityDefaults on, case-insensitive invitesFrom |
| OAuthGrantAnalyzer | 13 | empty, low-risk only, tenant-wide high-risk (ERROR), user-consented (WARN), empty clientId skip, scope trim, case-insensitive, mixed scopes detail, name fallback, unique clients dedup, full_access_as_user, IsHighRiskScope helper, mixed counts |
| TransportRuleAuditAnalyzer | 17 | empty, no-actions OK, forward externo (ERROR), forward interno OK, BCC externo, redirect externo, outbound connector (INFO), delete broad scope (WARN), delete narrow OK, modo Audit (INFO), disabled (INFO), disabled security keyword (WARN), smtp prefix strip, case-insensitive domain, no-domain skip, empty name, multi-findings same rule |
| SharedMailboxSignInAnalyzer | 6 | empty, disabled OK, enabled (WARN), unknown (INFO), empty UPN skip, mixed counts |
| SkuCatalog | 13 | resolve theory (E5/E3/SPB/F3/EntraID/Visio), fallback humanize, empty SKU graceful, case-insensitive, priority order Ent<Bus<Frontline<Sec, CategoryLabel mapping |

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
1. **Documentación técnica interna** — arquitectura + manual operación (ARCHITECTURE.md + RUNBOOK.md) [requiere petición explícita del usuario]
2. **Asset definitivo MSIX** — reemplazar PNG placeholders por branding + smoke test instalación end-to-end con cert real
3. **QA escenarios reales** — 100+ ops simultáneas bajo carga + scripted bulk CSV grande
4. **Terminal PowerShell embebido** (`EasyWindowsTerminalControl`) — útil para troubleshooting in-app
