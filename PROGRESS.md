# GREX365 v2.0 — Progress log

> **Documento maestro de seguimiento.** Mapea el estado del proyecto contra `Plantamiento_arquitectura_de_la_herramienta.md` (roadmap arquitectónico) y `deep-research-report.md` (research técnico). Toda feature shipped y todo pendiente vive aquí.

- Branch: `grex365-2.0` · Pushed up to `origin/grex365-2.0` (Sprints S–Y on remote, Sprint Z + AA local pre-push)
- Stack actual: **C# · .NET 10 · WPF + wpf-ui (Fluent) · MVVM (CommunityToolkit.Mvvm) · Serilog · Microsoft.Extensions.Hosting · Microsoft.ApplicationInsights**
- Tests: **1330 passing** (xUnit + FluentAssertions) — 516 Core + 814 App
- Última actualización: 2026-05-31 (sesión · Sprint AO — Offboarding: backbone + EXO + fixes + plantillas + layout + auto-reply grupo/excepción + delegación por EXO externo)

## Sprint AO · 2026-05-31 — Offboarding: backbone de seguridad y observabilidad

Mejora del módulo Offboarding a partir de un review técnico (9 áreas). Esta tanda cubre los items **sin nuevos scopes Graph**: pre-checks bloqueantes, dry-run, idempotencia y export auditable. Los items que requieren consentimiento nuevo en el App Reg productivo (limpieza de métodos MFA, OneDrive, Teams ownership, notificaciones `sendMail`, retirada de roles) quedan pendientes de aprobación.

- **Pre-checks como paso 0 (bloqueante)**: `OffboardingService` lee `MailboxInfo` (tamaño/holds/archivo vía EXO externo — `GetMailboxFactsAsync`, que existía pero no se invocaba) antes de tocar nada y emite el paso "Verificaciones previas". Detecta: cuenta ya deshabilitada, buzón ya compartido, >50 GB, hold activo.
- **Gate de licencias ampliado** (cierra hueco real de pérdida de datos): además de "no quitar licencia si la conversión a shared falló", ahora tampoco la quita si el buzón resultante seguiría necesitando licencia — >50 GB (límite de shared sin licencia, requiere EXO Plan 2) o hold activo (litigation/in-place). Resultado OMITIDO + `success=false`.
- **Idempotencia** (re-run seguro): cuenta ya deshabilitada → omite el disable pero revoca sesiones; buzón ya compartido → omite la conversión y **permite** liberar licencia.
- **Aviso de licencia heredada de grupo**: si tras `RemoveAllLicensesAsync` siguen asignadas licencias (solo se quitan las directas), se marca AVISO con la remediación (sacar del grupo de licencias) en vez de fingir éxito.
- **Dry-run** (`OffboardingOptions.DryRun`): ejecuta los pre-checks read-only y simula cada paso mutador (estado SIMULADO) sin tocar el tenant. Checkbox en la vista. Para rehearsal seguro en producción sobre `testeo*`.
- **Export auditable**: `OffboardingReport.ToJson/ToCsv` (puro, testeado, escape RFC-4180); botón "Exportar CSV" copia al portapapeles los resultados del último run (1 fila por paso, con timestamps).
- **Observabilidad**: `OffboardingStep.At` (timestamp), `OffboardingResult.StartedAt/EndedAt/DryRun`.
- **Validación en vivo (read-only, sin mutaciones)**: probe headless efímero conectó por cert al tenant productivo y corrió dry-run contra los 6 `testeo*`. Confirmado: pre-check lee facts EXO reales (testeo224 → SharedMailbox 0 GB → convert OMITIDO idempotente), already-disabled skip (Testeo3.1), degradación sin buzón, todos los pasos SIMULADO, 0 cambios.

### Pasos EXO de finalización (mismo sprint, commit 2)
Portados de `Invoke-OffboardingWizard.ps1` al servicio .NET como pasos **opcionales** post-conversión (sin scopes nuevos — `Exchange.ManageAsApp`):
- **Auto-reply** (#5 del review): `IExternalExoOps.SetAutoReplyAsync` → `Set-MailboxAutoReplyConfiguration` (OOO permanente interno+externo). Opción `OffboardingOptions.AutoReplyMessage`. TextBox en la vista.
- **Forwarding** (#6 del review): `SetForwardingAsync` → `Set-Mailbox -ForwardingSmtpAddress -DeliverToMailboxAndForward`. Opción `ForwardTo`; checkbox "Reenviar al delegado" usa el mismo destino que el FullAccess.
- **Hide-from-GAL**: `HideFromGalAsync` → `Set-Mailbox -HiddenFromAddressListsEnabled` con manejo del caso híbrido on-prem (SKIP con guía, como el legacy). Checkbox.
- Best-effort: fallo de finalización = AVISO **no fatal** (no deshace el offboarding). Solo corren con buzón presente + EXO externo; si no → OMITIDO.
- Tests: **+4 Core** (corren/dry-run/sin-EXO/fallo-no-fatal) **+1 App** (opciones fluyen al servicio).
- Total tras ambos commits: **1313** (503 Core + 810 App).
- **Pendiente offboarding**:
  - Sin scope nuevo: SendAs al delegado (FullAccess ya en VM), quitar de DLs/M365 groups (`GroupMember.ReadWrite.All` ya concedido — vía Graph, no EXO).
  - **Requiere scopes nuevos + decisión del usuario** (consentimiento Global Admin en App Reg productivo): limpieza de métodos MFA (`UserAuthenticationMethod.ReadWrite.All`), OneDrive delegación+tamaño (`Files/Sites`), Teams ownership transfer, notificaciones (`Mail.Send`), retirada de roles Entra.

### Fixes tras prueba real (commit 3)
Detectado al ejecutar offboarding real sobre un `testeo*` sin buzón → `Convertir a buzón compartido — ERROR pwsh exit 1`.
- **Error EXO real ahora se ve**: `ExternalExoOps.RunAsync` envuelve el body en try/catch y emite el mensaje real con marker `###GREX-ERR###`; antes el error CLIXML se filtraba y solo quedaba "pwsh exit 1". Ahora muestra p.ej. "el objeto no se encontró".
- **Sin buzón → se omiten pasos de buzón** (paridad legacy `if ($mbox)`): si EXO está conectado pero no hay buzón legible, `OffboardingService` marca convert OMITIDO "El usuario no tiene buzón…" y **permite** liberar licencia (no hay buzón que dejar huérfano). `willBeShared` ahora exige `convertSucceeded`; el gate de convert-fallido excluye el caso sin-buzón.
- Diagnóstico EXO read-only confirmó: de los 6 `testeo*`, solo **testeo224** tiene buzón (SharedMailbox); los otros 5 no tienen buzón → de ahí el error original.
- **App**: el assembly es `Grex365.exe` y el usuario corre **Release** (`bin/Release/net10.0-windows`). Rebuild Release + relanzada para que se vean los cambios de UI (antes se compilaba Debug).
- Tests: **+1 Core** (`NoMailbox_SkipsConvert_AllowsLicenseRemoval`, repro del bug). Total **1314**.

### Plantillas de auto-reply + UI más grande (commit 4)
- **Plantillas**: `OffboardingAutoReply.Render` (puro) substituye `{usuario}` / `{delegado}` por-usuario al ejecutar. VM expone catálogo `AutoReplyTemplates` ("Baja — ya no trabaja aquí", "Contactar con el delegado", "Personalizado"); ComboBox pre-selecciona la primera (templates-first), al elegir rellena el textbox (editable). Render por-target: {usuario}=DisplayName, {delegado}=delegado (batch) / Upn+DelegateToAll (single).
- **UI más grande** (queja "todo demasiado pequeño"): checkboxes 13→15, títulos de sección 16→18, label auto-reply 16, textbox auto-reply MinHeight 54→110, ComboBox FontSize 15, botones FontSize 15 + padding. Hint con los tokens disponibles.
- Tests: **+5 Core** (render: substitución, case-insensitive, vacíos, sin tokens) **+3 App** (preselección, Personalizado limpia, tokens renderizados a opciones). Total **1322** (509 Core + 813 App).

### Layout horizontal + auto-reply grupo/excepción (commit 5)
- **Layout en columnas** (queja "todo demasiado vertical, paneles largos"): listas de candidatos y de cola en `UniformGrid Columns=2` (2 por fila); checkboxes de acciones (Deshabilitar/Convertir/Quitar lic./Dry-run) y de finalización (forward/hide) en 2 columnas. Reduce la altura de los paneles ~50%.
- **Auto-reply inteligente** `OffboardingAutoReply.Resolve` (puro) — modelo grupo + excepción:
  1. override por usuario (`OffboardingTarget.AutoReplyOverride`, campo por fila en la cola) **siempre gana**;
  2. si hay delegado (por-usuario `DelegateTo`, o el global `DelegateToAll`) → plantilla global con `{delegado}` substituido;
  3. **sin delegado → mensaje "sin reemplazo" automático** (`Offboarding.AutoReply.NoDelegate`);
  4. auto-reply global vacío + sin override → `null` (paso omitido).
  Resuelve el escenario real (5→Pepe / 4→Marta / 1→Fernando / 2 sin delegado): una sola config compartida (plantilla global) + delegado por usuario/grupo + excepciones individuales. El delegado por-fila ya alimentaba `{delegado}`; ahora además el caso sin-delegado y el override propio.
- Tests: **+6 Core** (Resolve: override gana, con/sin delegado, off, theory por delegado) **+1 App** (batch: per-target delegado + fallback sin-delegado). Total **1329** (515 Core + 814 App).

### Fix: delegación FullAccess/SendAs por EXO externo (commit 6)
Error real en ejecución: `[HttpResponseMessage] does not contain a method named 'GetResponseHeader'` — el bug clásico de EXO **in-proc**. Causa: la delegación FullAccess post-éxito en el VM usaba `ISharedMailboxService.ApplyPermissionAsync` (RunspacePool interno). Fix:
- `IExternalExoOps.GrantDelegateAsync(mailbox, delegate, sendAs)` → pwsh externo: `Add-MailboxPermission FullAccess -AutoMapping:$false` + `Add-RecipientPermission SendAs`.
- `OffboardingOptions.DelegateMailboxTo`; el servicio concede FullAccess+SendAs como paso de finalización (vía `_externalExo`), AVISO no-fatal.
- VM: pasa `DelegateMailboxTo = del` por target/single; **eliminada** la delegación in-proc (`_mailboxes`) y su campo/param del ctor. Ahora TODO EXO de offboarding va por pwsh externo.
- Tests: **+1 Core** (`Delegate_GrantsFullAccessAndSendAs_ViaExternalExo`, con `ISharedMailboxService` strict que NO debe tocarse). Total **1330**.
- **Validado en vivo contra testeo224** (real, con revert): FullAccess + SendAs + auto-reply → **OK por pwsh externo** (las 3 que petaban con GetResponseHeader in-proc). Confirma el fix. Hide-GAL → error legítimo "objeto sincronizado desde organización interna" = testeo224 es **híbrido** (sync AD on-prem) → `HideFromGalAsync` lo maneja como SKIP con guía. Endurecido el regex de detección híbrida (`ámbito de escritura|organizaci.n interna|local organization|synchroniz…`) para no depender solo de "sincroniz". Aviso: muchos buzones de Andersen pueden ser híbridos → hide-GAL se aplica en AD local.

## Sprint AN · 2026-05-30 — Remate funcional + UX

- **Usuarios**: panel de detalle rico (`UserDetailsView`) ahora embebido inline en la columna derecha, se carga al seleccionar (1 clic). Drawer se mantiene solo para Groups; se suprime en la página Usuarios (`MainViewModel.SyncUserDrawerVisibility`). `UserDetailsView.ShowClose` DP nueva.
- **Licencias**: arreglada la lupa de búsqueda — `NullToCollapsedConverter` trataba `""` como no-nulo y dejaba la "×" siempre visible; ahora colapsa con string vacío. Foreground/CaretBrush explícitos en los TextBox de filtro (legibilidad modo oscuro).
- **Auditoría hardening**: guards de respuesta null en Graph (/users, /groups, miembros); pre-check de conexión Exchange en los 4 scans EXO (mensaje accionable en vez de error cmdlet críptico).
- **Offboarding guiado** (acción "Corregir" en hallazgos Disabled+License): flujo recomendado MS — bloquear sign-in + revocar sesiones → convertir buzón a compartido → liberar licencias. **Bug corregido**: `OffboardingService` quitaba licencias ANTES de convertir (pérdida de datos: borrado a 30 días). Ahora convierte primero y omite la liberación de licencia si la conversión falla. `MailboxInfo` enriquecido (hold/archive/tamaño) para avisos pre-acción (límite 50 GB, retención).

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

### 2026-05-28 (sesión autónoma) — Sprint AM · i18n ViewModels (cierre D6 al 100%)

User: "Remata por completo el proyecto, que quede todo perfecto, sobre todo a nivel de interfaz y scripts". Auditoría inicial (Explore agent) detectó el último hueco real de interfaz: **D6 i18n estaba migrado en las 20 vistas XAML pero NO en los ViewModels** — `StatusMessage`, títulos de diálogo `ConfirmAsync/ShowAsync`, y títulos de `OpenFileDialog/SaveFileDialog` seguían siendo literales ES hardcoded.

**Sprint AM — i18n ViewModels** `<commit AM>`:
- **16 ViewModels migrados** a `L10n.Get` / `L10n.Format`: Users, Groups, MailboxRules, Offboarding, Onboarding, PsConsole, DomainCheck, MailFlow, SharedMailbox, UserDetails, AuditLog, Audit, TenantHealth, Connect, CertWizard + MainViewModel (export-log dialog title).
- **L10n.cs +~230 keys ES+EN** bajo namespaces `Common.Status.*` (Cancelled/CancelledByUser/Error/Exported/NoResultsToExport compartidas), `Common.Confirm.Title`, `Common.Dialog.SaveResults`, y per-VM `<VM>.Status.*` / `.Confirm.*` / `.Dialog.*`. Cubre: validaciones ("Buzón vacío.", "UPN vacío.", "Selecciona un usuario."), progress ("Cargando...", "Buscando..."), summaries (incl. los 8 readouts computados de Audit MFA/OAuth/AppCreds/TenantDefaults/Privileged/CA/Shared/Transport vía format-string multi-arg), bodies+títulos de confirmación destructiva, y títulos de file-dialogs.
- **Test infra race-safe** (`TestAssemblyConfig.cs`): `[assembly: CollectionBehavior(DisableTestParallelization = true)]` + `[ModuleInitializer]` → `L10n.Initialize("es")`. Los VMs ahora resuelven L10n en runtime, y las 4 clases que mutan estado L10n estático (`L10nTests`/`L10nExtensionTests`/`FirstRunWizardVMTests`/`BoolToOnOffConverterTests`) restauran "es" en `Dispose` en lugar de `Reset`. Las aserciones literales ES de los VM tests existentes siguen pasando sin cambios.
- **`VmStatusL10nTests` +24**: wording exacto ES de 18 keys representativas + formateo (`Common.Status.Error` con arg, `Exported`, summaries con count, `ToggleBody` compone verbo localizado, EN difiere de ES). La paridad EN total ya la cubre `L10nTests.EnDict_HasTranslationForEveryEsKey` (itera TODAS las keys).
- **Decisión scope**: literales de log Serilog, headers CSV y status-codes de datos (`AGREGADO`/`OK`/`INVALIDO`) NO se localizan (son tokens internos/de datos, no UI). El placeholder `"—"` de `SettingsViewModel.CertStatusMessage` se deja (em-dash, no traducible).
- Sweep final `grep` sobre `src/.../ViewModels/*.cs`: cero literales ES user-facing restantes. **D6 i18n cerrado 100% (vistas + ViewModels)**.

**Scripts**: parse-check de los 28 `.ps1` de `GREX365/` (Parser AST, sin ejecutar) → **0 errores de sintaxis**. Son source-of-truth validada; `Test-AllScripts.ps1` + `Invoke-SelfTest.ps1` requieren tenant real (objetos `testeo*`), no automatizable headless.

**Estado**: build limpio 7 proyectos 0 errores/0 warnings. Tests **1291** (485 Core + 806 App, +43 desde 1248).

### 2026-05-27 (cont. autónoma) — Sprints AG–AI · converter L10n + prefs fail-soft + release scaffold

**Sprint AG — BoolToOnOffConverter L10n** `<commit AG>`:
- Cierra los 3 ConverterParameter hardcoded ES detectados en Sprint AB (`'Habilitado/Deshabilitado'`, `'Guest/Member'`, default `conectado/desconectado`).
- Converter detecta dotted tokens como L10n keys (`Status.Enabled/Status.Disabled`) y resuelve via `L10n.Get` en convert time. Backwards-compat preservada: tokens sin dot tratados como literales.
- +8 keys ES+EN — Common.Connected/Disconnected, Status.Valid/Invalid/Enabled/Disabled, UserType.Guest/Member.
- 4 callsites XAML migrados (ConnectView CertIsValid, UserDetailsView 2x, UsersView 2x).
- `BoolToOnOffConverterTests`: 11 tests cubren no-param + dotted + plain + mixed + unknown-key fallback + non-bool guard + ConvertBack throws.
- **D6 i18n cerrado 100%** sin strings hardcoded ES restantes en surfaces migradas.

**Sprint AH — H1.5.5 prefs fail-soft** `<commit AH>`:
- Bug antes: launching con `user_preferences.json` corrupto (edit manual / disk truncation / version drift) → `JsonException` → app dead on arrival.
- Fix: `JsonPreferencesStore` + `JsonCertConfigStore` `LoadAsync` catch `JsonException` + `IOException`, cuarentenan archivo via `CorruptFileQuarantine.MoveAside` → `*.corrupted-yyyyMMddHHmmss.bak`, retornan defaults.
- `+5 tests` Core — corrupt-returns-defaults, quarantine-preserves-bytes, corrupt-then-save-overwrites, cert-corrupt-quarantines, empty-file-defaults.
- ROADMAP H1.5.5 cerrado.

**Sprint AI — H5.6 release scaffold** `<commit AI>`:
- `CHANGELOG.md` root (Keep a Changelog ES) — scope a v2.0 rewrite, separa del legacy CHANGELOG en `GREX365/Seguimiento Claude/`. Documenta "how to cut a release" 6-step.
- `.github/release-template.md` con `${TAG_NAME}`/`${TAG_NAME_ANCHOR}` placeholders. Cubre highlights, assets table (portable EXE / MSIX / appinstaller), instalación, verificación, breaking changes section, next steps.
- CI workflow: nuevo job `release` (ubuntu-latest, permissions: contents:write) gated por `refs/tags/v*`, depende de publish + msix. Render template via sed, crea draft GH release via `softprops/action-gh-release@v2`, attachea Grex365.App.exe + .msix + .appinstaller. Prerelease auto-detect cuando tag contiene `-`.
- ROADMAP H5.6 cerrado.

**Total 9 commits this round** (Sprints AA–AI): tests 855 → **1248** (+393).

### 2026-05-27 (sesión cont. autónoma) — Sprints AC–AF · cerrar D6 i18n (20/20 surfaces)

Tras Sprints AA+AB (Shell+Dashboard+Users+UserDetails), continuar batch hasta agotar las 12 superficies restantes. **D6 i18n cerrado en MVP scope.**

**Sprint AC — 4 small views** `<commit AC>`:
- MailFlow, Offboarding, DomainCheck, PsConsole (~365 líneas combined)
- +45 keys ES+EN (11+11+9+14)
- +45 tests Theory en SmallViewsBatch. Total tests: **994**

**Sprint AD — 4 medium views** `<commit AD>`:
- Onboarding, SharedMailbox, MailboxRules, CertWizard (~700 líneas combined)
- +89 keys ES+EN (22+17+22+28). Multi-line NextSteps body cubre 4 instrucciones admin
- +89 tests Theory en MediumViewsBatch. Total tests: **1083**

**Sprint AE — Groups + TenantHealth** `<commit AE>`:
- GroupsView (330) + TenantHealthView (245) — big surfaces
- +46 keys ES+EN. Groups incluye bulk-create panel con tooltips por RadioButton + Auto/M365/DL
- Surgical edits (no rewrite) para preservar animaciones + license card ProgressBar templates
- +46 tests. Total tests: **1129**

**Sprint AF — AuditView + AuditLogView** `<commit AF>`:
- Cerrar las 2 últimas superficies (339 + 182 líneas)
- +95 keys ES+EN (70 Audit + 25 AuditLog) — la batch más grande del proyecto. Cubre 13 security audits con tooltips, 3 quick-scenario buttons, 4 sections (Identity/Security/Mail), 4 KPI summary cards, findings filter + pills + columns, AuditLog monthly metrics chips + DataGrid headers
- AuditSeverityToGlyph/Brush converters preservados — pure visual mappings sin texto
- +95 tests Theory en AuditBatch. Total tests: **1224**

**Coverage final L10n D6: 20/20 superficies migradas**:
About · Wizard · Settings · ConnectView · MainWindow shell · Dashboard · Users · UserDetails · MailFlow · Offboarding · DomainCheck · PsConsole · Onboarding · SharedMailbox · MailboxRules · CertWizard · Groups · TenantHealth (Licencias) · Audit · AuditLog

**Pendientes (deferred to dedicated sprint)**:
- 3 ConverterParameter strings hardcoded ES en `BoolToOnOffConverter` (`'Habilitado/Deshabilitado'` x2, `'Guest/Member'`) — requieren rewrite del converter para aceptar L10n keys o multi-binding

### 2026-05-26 (sesión cont. autónoma) — Sprint AB · i18n Users + UserDetails

**Sprint AB — UsersView + UserDetailsView** `c70f32d`:
- L10n.cs: +48 keys ES+EN — `Users.*` (22): eyebrow/title/subtitle, search panel (label/placeholder/2 buttons), empty-state (title+hint), none-selected fallback, 4 detail row labels (UPN/Mail/Cuenta/Licencias), 7 action buttons (Enable/Disable/RemoveLic/BulkCsv/ExportBulk/LoadSkus/AssignLicense), "Pertenece a" section. `UserDetails.*` (26): eyebrow + Close, 4 quick-action pairs (Toggle/ResetPassword/RevokeSessions/RemoveAll con tooltips funcionales), Identidad section + 5 field labels, Licencias card (title/total suffix/per-license Remove), Assign-new flow (label + filter tooltip + clear tooltip + " libres)" suffix + Asignar button), Memberships card + " grupos" suffix.
- UsersView.xaml: xmlns l + 14 reemplazos. Selected-user fallback usa `FallbackValue + TargetNullValue` pair para que cubra both null y unset.
- UserDetailsView.xaml: 22 reemplazos. Tooltips de quick-actions migrados (incluye texto técnico API "POST /users/{id}/revokeSignInSessions...").
- L10nTests: +48 Theory cases. **949 tests total** (480 Core + 469 App).

Pendiente identificado: 3 sitios con strings hardcoded en `BoolToOnOffConverter` ConverterParameter — `'Habilitado/Deshabilitado'` y `'Guest/Member'` quedan ES-only. Requieren rewrite del converter (interpretar key vs literal) o multi-binding. Diferido a sprint dedicado.

Coverage L10n actualizado: 8/20 superficies migradas — Users + UserDetails añadidas a (About, Wizard, Settings, ConnectView, Shell, Dashboard).

### 2026-05-26 (sesión cont. autónoma) — Sprint AA · i18n shell+dashboard + GoTo NavKey + DPI manifest hygiene

User request: continuar el proyecto, usar plugins/skills disponibles. Rebase + push de 11 commits locales (Sprints U–Z + license filter) tras pull --rebase de README update upstream.

**Sprint AA — i18n shell + Dashboard** `<pending push>`:
- L10n.cs: +46 nuevas keys ES+EN — `Shell.Sidebar.*` (Theme/Settings + tooltips), `Shell.Log.*` (title + 5 filtros + 2 botones + tooltip), `Shell.Status.*` (Graph/Exchange/Tenant/Account + About/Log/Disconnect buttons), `Dashboard.*` (eyebrow/title/subtitle/quick-actions/status/metrics/last-audit/pills).
- MainWindow.xaml: xmlns `l="clr-namespace:Grex365.App.Xaml"` + reemplazo `{l:L10n Key=…}` en sidebar 4 botones, log panel título + 5 checkboxes + 2 botones, status bar 6 labels + 3 botones. 22 strings migrados.
- DashboardView.xaml: refactor completo via `Write` — eyebrow/title/subtitle, 5 quick-action buttons, 4 metric cards (Graph/Exchange/Tenant/Account) con captions, Last-audit panel (botón "Abrir Auditoría" + 3 pills ERROR/WARN/INFO + timestamp formato). 24 strings.
- DashboardViewModel.GoTo: bug fix — match por stable `NavKey` primero, fallback a `Title`. Antes solo matcheaba por Title localizado → CommandParameter "Salud tenant" silently broken tras rename a "Licencias" (Sprint O). Dashboard CommandParameter values switched to NavKey ("Nav.Connection"/"Nav.Licenses"/etc.) → i18n-safe.
- L10nTests: +46 nuevos casos Theory (22 Shell keys + 24 Dashboard keys) ES+EN dual-check. **901 tests** (480 Core + 421 App).

**Build hygiene**:
- Grex365.App.csproj: `<NoWarn>$(NoWarn);WFO0003</NoWarn>` — WinForms-only DPI advisory que pide `ApplicationHighDpiMode`. app.manifest sigue canónico para WPF (`true/PM`), property conflictaría. Comentario explica decisión.

**Repo cleanup**:
- 13 PNGs root-level eliminados — capturas dev ad-hoc, sin referencias en docs/code (grep verified). PROGRESS/docs/README text-only.
- `.gitignore`: `_bmad/` (replaces narrow `_bmad/custom/config.user.toml`) + `.claude/skills/` añadidos — checkouts per-developer.

**Coverage L10n actualizado** (de 4 superficies a 6):
- ✓ AboutWindow, FirstRunWizardWindow, SettingsWindow, Views/ConnectView (Sprints S–U)
- ✓ MainWindow shell (Sprint AA)
- ✓ Views/DashboardView (Sprint AA)
- ⏳ Pendientes: Users, Groups, TenantHealth, Onboarding, Offboarding, SharedMailbox, MailboxRules, MailFlowRules, Audit, AuditLog, PsConsole, CertWizard, DomainCheck, UserDetails (14 views restantes).

### 2026-05-24 (sesión cont. autónoma — UX overhaul) — Sprints V–Y · paleta azul + contrast + tray + silent notifications

User feedback explícito sobre calidad visual + UX: "fase seria de edición, presentación y rediseño completo... Sustituir el morado actual por una línea visual más moderna en tonos azulados/celestes. Corregir problemas de legibilidad: en ajustes, por ejemplo, el modo oscuro deja partes en gris y el texto casi no se puede leer." + lifecycle: "X = minimizar a tray, mantener en segundo plano, reconexión silenciosa, no toasts molestos".

Decisiones tomadas vía AskUserQuestion:
- Paleta: **Azul corporativo sólido** `#1E40AF → #2563EB → #3B82F6` (indigo deep, no púrpura).
- Lifecycle: **Minimize-to-tray** (X oculta a tray, menú tray: Abrir/Estado/Reconectar/Salir).
- Notifications: **Solo status bar silencioso** (sin toasts ni dialogs en reconnect/auto-connect).

**Sprint V — Paleta azul corporativo** `444c1a5`:
- App.xaml: BrandAccent gradient `5B8DEF→7A78F0→A66CF4` (azul→púrpura) → `1E40AF→2563EB→3B82F6` (indigo→sky). BrandPurple eliminado, nuevo BrandSky #7DD3FC. AccentFill* wpf-ui tokens override migrados a Mid (#2563EB) en lugar de Start.
- GlassBorderHighlight gradient migrado a blue indigo. DropShadowEffect colors usan BrandAccentMid.
- GxLogoMark + GxLogoBadge: canvas 42x28 (era 40x28), stroke 3 (era 3.2), proporciones tightened. Badge añade AccentGlowSoftEffect drop-shadow.

**Sprint W — Theme-aware contrast** (commit conjunto V+W):
- App.xaml: nuevos styles `FieldHintText`/`MutedText`/`CaptionText` con `Foreground=TextFillColorSecondaryBrush` (wpf-ui token auto-adjust per theme).
- SettingsCardDescription / SettingsCardIcon / SettingsGroupHeader: Opacity reemplazado por Foreground=Secondary brush. SettingsCardIcon ahora BrandAccentSolid (azul).
- Batch sweep 17 XAML files: `Opacity="0.5"`/`0.55"` → TextFillColorTertiaryBrush; `Opacity="0.6"`/`0.65"`/`0.7"` → TextFillColorSecondaryBrush. Auto-adjust per theme via DynamicResource. Border con Opacity 0.7 (GridSplitter handle) preservado (no es TextElement).
- Fixes user feedback dark mode "gris-sobre-gris" gracias a wpf-ui tokens que tienen contrast correctness baked-in.

**Sprint X — Background lifecycle** `40f0f67`:
- Grex365.App.csproj: `UseWindowsForms=true` (System.Windows.Forms.NotifyIcon). Implicit usings `System.Windows.Forms` + `System.Drawing` removidos vía `<Using Remove="...">` para evitar collision con WPF UserControl/Brush. TrayIconService usa aliases D=/D2=/WinForms=.
- `TrayIconService` nuevo (Grex365.App.Services): NotifyIcon con ContextMenuStrip (Abrir GREX365/Estado/Reconectar ahora/Salir). Icono procedural 32x32: gradient `1E40AF→3B82F6` 45° + letra "G" white SemiBold + status dot overlay (rojo/ámbar según graphConnected+exchangeConnected). UpdateConnectionState re-renderiza icon dynamic.
- App.xaml.cs: `ShutdownMode=OnExplicitShutdown` (proceso sobrevive cierre main window). `WireTrayIcon(monitor)`: suscribe monitor.PropertyChanged → tray.UpdateConnectionState. OpenRequested → Show+Activate+focus. ReconnectRequested → TryAutoConnectAsync. ExitRequested → flag `_explicitExitRequested` + Shutdown(). OnExit dispose tray.
- MainWindow.xaml.cs: `OnMainWindowClosing` cancela close + Hide() si !ExplicitExit. SaveWindowState sigue persistiendo. Topmost ping (true→false) en OpenRequested para foreground.

**Sprint Y — Silent notifications** (commit conjunto X+Y):
- UiLogSink.OnEntry ahora gated por `ShouldNotify` policy:
  - Severity Info/Debug nunca emit toast.
  - Severity Ok nunca emit toast (success surface vía status bar).
  - `SilentSources` HashSet (case-insensitive): {AutoConnect, Connect, ConnectionMonitor, TenantLock, Settings, EXO, Graph} bypass toast incluso para Error/Warning. Status bar + log panel canónicos.
  - Resto source/severity combos siguen disparando toast (e.g. RBAC denial en Users/Groups VM).
- UiLogSinkNotificationTests +19: ShouldNotify gate cases (null notifier, silent sources case-insensitive ×9, non-silent error/warning, info/debug suppressed theory, Ok gate).

**Verificación arranque**: app launched OK, log confirma flujo silencioso:
```
[AutoConnect] Cert válido detectado, conectando...
[Graph] OK · Conectado. Organización: Andersen
[EXO] OK · Exchange Online conectado.
[AutoConnect] OK · Conectado automáticamente.
```
Cero MessageBox / Snackbar disparado durante auto-connect ✓.

**Estado final Sprints V–Y**: 855 tests verdes (480 Core + 375 App, +19 UiLogSink). Build clean 7 projects 0 errors (2 warnings WFO0003 high-DPI manifest non-fatal). 2 commits locales (`444c1a5` + `40f0f67`). Total sesión: 7 commits desde S.

**Pendiente verificación visual humana**: usuario debe abrir app y validar:
1. Paleta azul corporativo aplicada (no purple) en sidebar/botones/glow/logo
2. Settings dark mode legibilidad (no gray-on-gray) ambos temas
3. X cierra a tray (no shutdown), tray menú functional
4. Tray icon visible con status dot rojo/ámbar/verde
5. Logo blue gradient en sidebar + AboutWindow

### 2026-05-24 (sesión continuación autónoma) — Sprints S + T + U · i18n expansion + BulkGroupPlanner + ConnectView

**Sprint U** (cont. autónoma) — 1 commit `6f96dab`:
- L10n +75 keys `Connect.*`: Eyebrow/Title/Subtitle, labels MsGraph/ExchangeOnline, Section.Actions/Tenant, 4 botones, ExoModule box (Title/Hint/Probe/Install), DeviceCode box (Title/3 steps), Tenant fields (Id/Domain/Account), Cert section (GroupHeader/Status/AppId+Hint/Thumbprint+Hint/Expiry+Hint/ValidLabel), 8 Status templates. EN parity.
- `ConnectView.xaml`: 30 strings hardcoded ES → l:L10n. VM StatusMessage L10n diferido (dynamic error strings ya logged via Serilog, lower-prio).
- L10nTests +58 (Connect 36 dual-lang theory + Wizard 22 dual-lang theory).



Usuario: "continua con el proyecto, no pares hasta que termines, recuerda usar los ultimos plugins que instale y demas, saludos". Plugins nuevos detectados: BMAD (`_bmad/`) + Claude skills (`.claude/skills/`) instalados. Skip ceremonia BMAD formal (requiere sprint-status.yaml + epics.md sharded ausentes); aplico patrón directo `extract pure helper + tests` + `wire MarkupExtension XAML`. Backlog autónomo identificado: D6 i18n in-progress (Sprint R scaffold solo 39 keys) + memoria pendiente `GroupsVM BulkCreate tests`.

**Sprint S (i18n masivo + XAML MarkupExtension)** — 1 commit `8b7fbaa`:

1. `feat(i18n)` L10n expansion + L10nExtension:
   - `L10n.cs` dicts amplios ES+EN: 39 → 90 keys. Settings.* completo (40+ keys cubriendo Section.Connection/Plugins/Certificate, Connection.Cert/Traditional, Tenant.IdLabel/Placeholder/DomainLabel/Placeholder, EnforceTenantLock, LanguageLabel, Theme.Dark/Light/Auto/Hint, LogLevel.Debug/Information/Warning/Error/Hint, AppInsights/Placeholder/Hint, Rbac/Hint, Plugins.Hint/Empty/ModulesSuffix/Status.Loaded/Disabled/ErrorPrefix, Cert.AppId/TenantId/Organization/Thumbprint/Browse/Validate, Button.Reload/Save, SaveStatus.Prefix/ErrorPrefix/SavedLog). About.* (8 keys: Window.Title/TitleBar/Description/Version/Runtime/DataDir/Button.OpenDataDir/Close). Dialog.Close añadido. Common ampliado (Save/Delete/Edit/Add/Remove/Back/Next/Finish/Skip/Loading/Empty).
   - `L10n.Format(key, args)` helper nuevo: `string.Format(template, args)` con fallback a template crudo si FormatException o args null/empty.
   - `L10n.KnownKeys` exposed `EsStrings.Keys` para tests parity loop.
   - `L10nExtension : MarkupExtension` (nuevo `Grex365.App.Xaml`): `{l:L10n Key=...}` en XAML — habilita binding-less translation. ProvideValue → `L10n.Get(Key)`, empty/null key → empty string.
   - `SettingsWindow.xaml`: 28 strings hardcoded ES → `{l:L10n Key=...}`. `AboutWindow.xaml`: 13 strings → l:L10n. `SettingsViewModel.cs` localizado: plugin status strings (Cargado/Deshabilitado/"Error: " prefix) + SaveStatus prefix vía Format con timestamp + Error prefix via Format.
   - `L10nTests +75` (37 → 112): theories Settings (42 ES + 8 EN), About (8 ES + 3 EN), Dialog/Common parity (23 ambos idiomas), Format (7: placeholder substitution, multi-args, no-args, invalid placeholder, null-args, ES/EN SaveStatus timestamp), KnownKeys + EN-parity-loop sobre KnownKeys.
   - `L10nExtensionTests +6` (nuevo): ProvideValue known/null/empty/unknown key, Constructor Key property, defaults to null.

**Sprint T (FirstRunWizard L10n + BulkGroupPlanner)** — 1 commit `7086892`:

1. `feat(i18n+core)`:
   - **L10n.cs +50 keys** `Wizard.*`: Window.Title/TitleBar, 5 step badges (Welcome→Summary), Welcome content (title/subtitle/SectionTitle/Bullet1-3/SkipHint), Connection (Title/Subtitle/DeviceCode/Cert + .Hint variantes), TenantLock (Title/Subtitle/Enable/IdLabel/DomainLabel/Hint), Theme (Title/Subtitle + Dark/Light/Auto + .Hint variantes), Summary (Title/Subtitle/Connection/TenantLock/Theme), Button.Skip, Status.Skipping/Saving/Saved, label templates ConnectionMethodLabel/TenantLockLabel (Cert/DeviceCode/Enabled/Disabled/Empty). EN parity total.
   - `FirstRunWizardWindow.xaml`: 35 strings hardcoded ES → `{l:L10n Key=...}`. Botones nav (Atrás/Siguiente/Finalizar) usan Common.* keys (ya disponibles en Sprint S).
   - `FirstRunWizardViewModel.cs`: `ConnectionMethodLabel` + `TenantLockLabel` ahora via L10n.Get + L10n.Format (template "Activado · ID/Dominio: {0} / {1}"). Status (Skipping/Saving/Saved/Error) localizadas. Error path usa Settings.SaveStatus.ErrorPrefix Format compartido.
   - **`BulkGroupPlanner`** extract (`Grex365.Core.Groups`): refactor pure helper de `GroupsViewModel.BulkCreateFromCsvAsync`. Resuelve memoria backlog `GroupsVM BulkCreate tests` (item 3 pendiente Sprint S sessions previas): OpenFileDialog blocking hace VM no testable; planner puro sí. API: `Plan(rows, choice) → BulkGroupPlan` record (M365Rows/DlRows/distinct counts/breakdown/typeHint). `BuildConfirmMessage(plan, count, domain)` para `_dialogs.ConfirmAsync` text. `Summarize(results)` para StatusMessage final. VM ahora 3 líneas vs 25 inline.
   - **Bug fixed planner**: empty/whitespace/unknown choice ahora cae a Auto path en lugar de "Forzado por usuario: TODOS los grupos como ." con string vacío (VM legacy comportamiento equivalente vía else clause; planner explicit guard mejora robustez).
   - `L10nCollection.cs` xUnit serial fixture: tests L10n + L10nExtension + FirstRunWizardVM ahora `[Collection("L10n")]`. L10n estático no soporta parallel execution (Reset corre durante otro test). Fix race: 2 fails → 0.
   - `FirstRunWizardVMTests +5`: English labels translated, empty-marker em-dash, mid-flight Saving status TCS, localized Saved status ES/EN. `BulkGroupPlannerTests` nuevo 16: Plan (8 Auto/ForcedM365/ForcedDl/case-insensitive/empty/null/unknown choice/distinct case-insensitive/breakdown omits zero-sides), BuildConfirmMessage (3: includes-fields, trims domain, null-guard), Summarize (4: null-guard, empty all-zero, counts by action, unknown ignored).

**Estado final Sprints S+T**: 778 tests verdes (480 Core +21, 298 App +5 sobre 470/108 baseline Sprint R). Build clean 7 projects 0 errors. 2 commits locales (`8b7fbaa` + `7086892`), no pushed pendiente. D6 i18n status `🟡 in-progress` → close-ish: Settings + About + FirstRunWizard externalizados; remaining hardcoded ES principalmente en StatusMessage de page-VMs (Users/Groups/Audit/MailboxRules/etc) — backlog continuable bajo demanda.

**Memoria nueva sesión**: actualizar `project_migration_progress.md` con `BulkGroupPlanner` + L10n + 778 tests si user pide. Sprint pattern aprendido: WPF MarkupExtension `MarkupExtensionReturnType(typeof(string))` simple + static singleton L10n.Get; tests deben serializarse via `[Collection]` cuando comparten estado estático.

### 2026-05-23 (sesión noche cont.) — Sprint R · i18n ES/EN scaffold + decisiones cerradas

User pidió continuar con backlog restante. Estimación weeks: "¿cuántas semanas más le echas?" → respondí 2-4 weeks bloqueada por arte+tenant+decisiones. User reforzó "si sigue con todo esto" + después "claro metele idiomas en la parte de setings / ajustes ingles y espa;ol". También aviso seguridad importante: "solo puedes hacer pruebas con usuarios testeo y grupos y listas de distribucion de testeo" — guardado como `feedback_only_testeo_objects` memoria.

**Sprint R (i18n + cierre decisiones)** — 3 commits:

1. `docs(roadmap)` cerrar D3 + D8:
   - **D3 Reports format** ✅ Closed — CSV + HTML + JSON ya shipped en Sprint J-K. XLSX descartado (sin Excel dep). DataGrid in-app cubierto por AuditView.
   - **D8 Min target OS** ✅ Closed — Win10 1809+ de-facto via .NET 10 runtime + app.manifest supportedOS GUIDs. Win11 inherita.
   - **D6 i18n** status 🔴 → 🟡 in progress.
   - **D5 Templates** notation: no demand observado, deferred post-v1.0.

2. `feat(i18n)` L10n + Settings + nav localized:
   - `L10n` static (Grex365.App): dicts hardcoded ES (canonical) + EN (parcial con fallback automatic a ES). API Configure/Initialize/Get/Reset. 39 keys: Nav.* (15) + NavCategory.* (7) + Settings.* + Dialog.* + Common.*.
   - Por qué hardcoded vs embedded JSON: WPF SDK strip JSON embedded resources (g.resources collision). Hardcoded más portable + reliable.
   - `UserPreferences.Language = "es"` default. Aplicado en App.OnStartup → L10n.Initialize(lang) antes de DI build.
   - Settings UI: nuevo ComboBox "Idioma / Language" (Español/English) antes del Tema. LanguageRestartHint property se llena tras Save si idioma cambió → muestra "Reinicia la aplicación..." (ES/EN según L10n).
   - MainViewModel refactor: NavigationItem ahora tiene `NavKey` + `CategoryKey` (identity stable across language). Helper `BuildNav(navKey, glyph, vmType, categoryKey)` resuelve title+category via L10n.Get. RequiresGraphKeys/RequiresExchangeKeys match contra NavKey en lugar de Title. LoadLastNavigation dual-match (NavKey directo + legacy Title fallback via NavTitleMigrator). PersistNavAsync guarda NavKey en lugar de Title.
   - L10nTests (36): null/empty/whitespace/unknown lang → DefaultLanguage, Initialize ES/EN, fallback chain, case-insensitive, Configure overload, Reset, SupportedLanguages, theory ×15 (es) + ×5 (en) cubre todas Nav.* keys.

**Memoria nueva**: `feedback_only_testeo_objects` — REGLA CRÍTICA: solo objetos `testeo*` para pruebas; tenant Andersen productivo, blast radius alto en ops accidentales. Trabajo local code/tests/docs no requiere check; aplica solo cuando hay comando contra tenant real.

**Estado Sprint R**: 652 tests verdes (459 Core + 193 App, +36 L10n). Build clean 7 projects 0 errors. 3 commits pushed. D3 + D8 closed. D6 in-progress (scaffold complete; expansion futura externalizar más strings VMs/views).

**Resumen sesión 2026-05-23 noche acumulada (P + Q + R)**:
- 19 commits pushed.
- Tests 469 → 652 (+183).
- Docs refrescadas (ARCHITECTURE / RUNBOOK / ROADMAP / MIGRATION / README).
- 8 helpers puros extraídos: NavTitleMigrator, LicenseFilterMatcher, NavRequirements, WindowPlacementGuard, GraphPermissionErrorDetector, PSModulePathSanitizer, ExeResolver, CsvEscaper + L10n.
- 1 bug latente fixed (CSV \r en 4 VMs).
- Decisiones cerradas: D3, D8 (cumulativas: 4 closed total ✅, 1 in-progress 🟡 D6).
- Idiomas ES/EN scaffold operativo en Settings.

Plantamiento status final: Fase 1-4 DONE, Fase 5 (MSIX scaffold + CI release) DONE — falta arte + smoke test, Fase 6 (telemetría + enterprise) **7/8 DONE** — solo QA escenarios reales restante (necesita tenant). Backlog autónomo agotado salvo iteración i18n expansion (externalizar más strings) que se puede continuar bajo demanda.

### 2026-05-23 (sesión noche autónoma cont.) — Sprint Q · Extract sweep + bug fix CRLF

Continuación autónoma tras Sprint P. User feedback "sigue con lo que falta" tras inicial cierre. Trabajo `extract pure helpers + tests` pattern aplicado consistentemente sobre code-behind / VMs / PS service helpers no testados.

**Sprint Q (refactor + bug fix)** — 6 commits:

1. `refactor(app)` NavRequirements + WindowPlacementGuard (+26 tests):
   - `NavRequirements` (`Grex365.App.ViewModels`): `IsEnabled(requiresGraph, requiresExchange, graphConnected, exchangeConnected) → bool` + overload sobre `NavigationItem`. MainViewModel.UpdateNavEnabledStates delega. Tests 15 (no-req always, requires-graph/exchange/both, overload on item real, null-item throws, truth table theory 9 casos).
   - `WindowPlacementGuard` (`Grex365.App`) + `VirtualScreenBounds` readonly record struct: `IsOnScreen(left, top, width, height, virtBounds)` con EdgeMargin 50px. MainWindow.IsOnScreen wraps `SystemParameters.VirtualScreen*`. Tests 11 (centered, top-left, fully-off cada lado, dual-monitor secondary, secondary-disconnected scenario, partially-visible 50px margin, edge-case-exactly-at-margin strict <, negative-top).

2. `refactor(audit)` GraphPermissionErrorDetector (+21 Core):
   - GraphAuditService.IsAuditLogPermissionError era private sin tests. Extract a `Grex365.Core.Audit.GraphPermissionErrorDetector` + amplía API con `IsAnyPermissionError` genérico (detect "Insufficient privileges", "Authorization_RequestDenied", "Forbidden" además de existing). Tests cubren null, empty, contains scope, generic phrase, case-insensitive theory 5, unrelated → false; IsAnyPermissionError theory 6 phrases, unrelated 4.

3. `refactor(ps)` PSModulePathSanitizer (+10 Core):
   - RunspacePoolHost.SanitizeModulePath + BuildSafeDefault eran private sin tests. Extract a `Grex365.PowerShell.PSModulePathSanitizer` con overloads test-friendly (userModulesPath inyectable). Tests cubren empty input default, WindowsApps drop (ACL workaround core), mixed paths only WindowsApps removed, UserPath prepended si missing, no duplicate si already, case-insensitive UserPath match, empty parts filtered, BuildSafeDefault 3 entries shape.

4. `refactor(ps)` ExeResolver (+11 Core):
   - ExchangeConnection.ResolvePwshExe era private inline. Extract a `Grex365.PowerShell.ExeResolver` con API genérica `Resolve(candidates, pathEnv, fileExists, fallback) → string` + default `ResolvePwsh()`. Tests cubren first-found wins, second cuando first missing, no-candidate → fallback, null/empty PATH → fallback, empty segments skipped, candidate-order respected, ArgumentNullException theory 3 (null candidates / null fileExists / null fallback).

5. `refactor(core)` CsvEscaper unify (+16 Core, **bug fix latente**):
   - 5 VMs (Users/Groups/SharedMailbox/Audit/Main) tenían private static Escape duplicado. Variation: 4 omitían `\r` check, solo MainViewModel.CsvEscape lo incluía. **Bug latente real**: si campo CSV contenía CRLF de Windows, los 4 VMs producían CSV mal formado (split por LF dentro de campo durante reload/parse).
   - Fix: extract `CsvEscaper.Escape(string?)` a `Grex365.Core.Csv` (canonical location junto a FlexibleCsvReader). RFC 4180-compliant detection (`,` `"` `\n` `\r`). 5 VMs ahora 1-line delegate. Mismo comportamiento + fix `\r` para los 4 que faltaba.
   - Tests 16: null/empty/simple/spaces no-quote, comma quoted, double-quote escaped, newline/CR/CRLF quoted, all-specials combined, theory 4 specials solos, quote-only edge case, tab NOT special.

**Estado final Sprint Q**: 616 tests verdes (459 Core + 157 App, +84 desde Sprint P start 532). Build clean 7 projects 0 errors. 6 commits pushed.

**Resumen sesión completa 2026-05-23 noche** (Sprints P + Q juntos):
- 16 commits pushed.
- Tests 469 → 616 (+147).
- Docs cerradas: ARCHITECTURE refresh + RUNBOOK nuevo + ROADMAP refresh + MIGRATION sync + README refresh.
- Code health: `.Result` purge + nav rename map ampliado + ComboBox focus.
- 7 nuevos helpers puros extraídos: NavTitleMigrator, LicenseFilterMatcher, NavRequirements, WindowPlacementGuard, GraphPermissionErrorDetector, PSModulePathSanitizer, ExeResolver, CsvEscaper.
- 2 nuevos tests para componentes Core previamente sin tests: InMemoryAuditFindingsStore, UserDetailsHost, LicenseCard.
- 1 bug latente fixed: 4 VMs CSV escape sin \r check.

Plantamiento status final: Fase 1-4 DONE, Fase 5 (MSIX scaffold) + CI release DONE — falta arte + smoke test (blocked external), Fase 6 (telemetría + enterprise) **7/8 DONE** — solo QA escenarios reales (necesita tenant) bloqueado externo.

Backlog autónomo agotado completamente. Pendiente todo bloqueado por entrada externa o decisiones D3/D5/D6/D7/D8 que requieren input usuario.

### 2026-05-23 (sesión noche autónoma) — Sprint P · Docs internas + code health + cobertura tests

Usuario delega trabajo ausente: "sigue con el proyecto, es un proyecto de semanas, hasta que no terminas todas las fases no pares". Memoria `feedback_keep_shipping` confirma push autónomo OK. Sesión cierra el item "docs internas" pendiente de Fase 6 y limpia code health.

**Sprint P (docs + code health + tests)** — 4 commits:

1. `docs(ARCHITECTURE + RUNBOOK)`:
   - ARCHITECTURE.md refresh contra estado real del repo (era stale 8 días, marcaba Fase 4-6 como NOT taken). Stack tabla versiones reales (WPF-UI 4.3.0, CommunityToolkit.Mvvm 8.4.2, Microsoft.ApplicationInsights 2.23.0, .NET 10), layout completo (3 proyectos + tests + samples + packaging/), 8 secciones nuevas: connections+auth (cert/device-code/AppReg auto-create/TenantLock/RBAC/Graph replica retry), audit subsystem (12 analyzers + HTML/JSON/CSV + baseline diff), plugin system (IModule + PluginLoader + AssemblyLoadContext), telemetry (ITelemetry/NullTelemetry/AppInsights + UiLogSink forward), theme auto-system (WindowsRegistryThemeProvider + SystemEvents), 14 módulos nav tabla con RequiresGraph/Exchange, packaging MSIX + portable single-file.
   - RUNBOOK.md nuevo (14 secciones): manual operacional desde install hasta troubleshoot. Instalación 3 paths (MSIX AppInstaller / portable / build), primer arranque (FirstRunWizard + device-code + AppReg auto-create + auto-connect), operación día a día (status bar, tema, atajos), 15 módulos detallados, plugins install/enable/troubleshoot, troubleshooting 10 escenarios (EXO no conecta, cert no encontrado, TenantLock mismatch, permisos faltantes Graph, bulk groups replica, audit vacío, tema no sigue, RBAC bloquea, MSIX no arranca), datos disco + backup/restore + desinstall, config avanzada (logging level, AppInsights, proxy), telemetría + privacidad, soporte + escalación, limitaciones conocidas.
   - **Cierra Fase 6 item "Crear documentación técnica interna (arquitectura, manual de operación)"**. Fase 6 ahora 7/8 (queda solo QA escenarios reales — no automatizable sin tenant real).

2. `chore(health)`:
   - AuditViewModel: 7 instancias de `task.Result` tras `Task.WhenAll` → reemplazadas por `await task.ConfigureAwait(true)`. Tras `WhenAll` la task ya está completada, `await` es zero-cost y desenvuelve excepciones limpias (sin AggregateException wrap). Tres callsites: RunIdentityAuditAsync, RunScenarioAtRiskAccountsAsync, RunScenarioPrivilegedAsync.
   - MainViewModel.RenamedNavTitles: añadidos 4 renames del refactor 2026-05-22 que faltaban en el dict (Mail flow → Flujo de correo, Audit log → Registro de auditoría, Cert Wizard → Asistente cert, DNS check → Comprobación DNS). Sin esto, usuarios beta con `LastSelectedNavigation` pre-rename perdían su selección al arrancar tras update.

3. `test(app)` — extract pure helpers + cobertura Sprints L-O (+48 tests):
   - `NavTitleMigrator` (nuevo public static, `Grex365.App.ViewModels`): extracted desde MainViewModel.LoadLastNavigation. RenamedNavTitles dict + `Resolve(saved) → string?`. MainViewModel delega al helper.
   - `LicenseFilterMatcher` (nuevo public static): extracted desde TenantHealthViewModel.LicenseFilterPredicate. `Matches(card, filter)` case-insensitive substring sobre FriendlyName/SkuPartNumber/CategoryLabel. VM delega ahora (1 línea).
   - `NavTitleMigratorTests` (24): null/empty/whitespace, 8 renames theory, case-insensitive, no-renamed pass-through, RenameMap contains-expected-keys + new-titles-resolve-to-self + no-cycles invariant.
   - `LicenseFilterMatcherTests` (11): null/empty/whitespace pass-through, match por friendly/sku/category, case-insensitive, trim, no-match, short-circuit any-field, null-card ArgumentNullException.
   - `LicenseCardTests` (13): LicenseCard.From factory sin tests dedicados. Known SKU → friendly+category resolved, unknown SKU → humanize fallback + Other category, zero-enabled → 0% Low + Available zero, half-used → 50% Low, utilization levels theory (60/79/80/94/95/100 → Low/Medium/High/Critical boundaries), over-consumed → Available clamped to 0 via Math.Max, seat counts preserved, priority > 0 para known SKU.

4. `test(core)` — cobertura componentes sin tests dedicados (+15 tests):
   - `InMemoryAuditFindingsStoreTests` (5): initial defaults nulls + 0, Update sets all five fields, Update raises PropertyChanged para LastAuditName+ErrorCount+WarnCount+InfoCount+LastRunAt, Update twice overwrites + LastRunAt advances, accepts zero counts.
   - `UserDetailsHostTests` (10): initial closed+null, RequestOpen sets state + fires OpenRequested + PropertyChanged para IsOpen y CurrentUserId, RequestOpen con whitespace/null no-op theory, RequestOpen idempotente (PropertyChanged value-gated no duplica), RequestClose desde open resets + fires CloseRequested, RequestClose desde closed no PropertyChanged pero CloseRequested sí, open→close→open ciclo.

**Polish UI Sprint L pendiente cerrado**: focus-ring accent en ComboBox (paridad con TextBox/PasswordBox shipped previamente). `IsKeyboardFocusWithin` trigger → border `BrandAccentSolid` + `AccentGlowSoftEffect`. App.xaml.

**Estado final sesión**: 532 tests verdes (401 Core + 131 App, +63 desde 469). Build clean 7 projects 0 errors. 4 commits pushed. Plantamiento status: Fase 1-4 DONE, Fase 5 MSIX scaffold DONE (blocked: arte definitivo + smoke test), Fase 6 7/8 DONE (blocked: QA escenarios reales con tenant). Backlog Plantamiento §6 cerrado (Terminal PS embebido shipped vía Consola PS). Backlog "Features útiles" 100% shipped.

### 2026-05-23 (sesión tarde) — Sprints M-O · UX overhaul + bug stomping

**Sprint M** — Sally UX pass (desaturar verde + restraint):
- BoolToBrushConverter: `BrushSemanticOk` (verde) → `BrandAccentSolid` (azul). Connection dots ya no verdes.
- Paleta brand re-tuned: `#5B8DEF → #A66CF4` (menos saturado). `BrandCyan` teal `#22D3EE` → sky `#7AB7FF`.
- Sidebar gradient overlay quitado (tintaba Mica verdoso).
- Glass borders + glow effects retirados de status dots, ProgressBar, severity pills. Glow reservado a focus TextBox/PasswordBox + hover interactivos.
- Card radius 12→8, padding 20→14. PageRoot 32,28→22,18. MetricCard sin hover translate.
- HeroCard colapsa a alias Card. CardHover sin animación translate.
- Page transition 320ms fade+slide+scale → 120ms fade only.
- 14 views: hero badge → eyebrow + título Light pattern (Linear/Raycast estilo). Dashboard también tonado.
- Sidebar logo: bare GX 22px monochrome (sigue TextFillColorPrimary) + wordmark + version Cascadia Mono. NO pill, NO glow, NO gradient.
- Nuevo `GxLogoBadge` style retiene el pill gradient solo para AboutWindow/splash.
- GX path refinado: G como D estirada stub 60% altura, X cruce off-center 55%.

**Sprint N** — Bulk groups bugs (3):
- `GroupName` con `@` (SMTP completo, ej. `Testeo774@es.andersen.com`) → doble `@` al concatenar dominio. Fix: si `rawName` contiene `@`, lo usa como email y deriva displayName del local part. Aplica M365 + DL paths.
- Graph eventual consistency: tras `POST /groups`, members read/write fallaba con `Resource ... does not exist`. Nuevo `WithGraphReplicaRetryAsync` exponential backoff (1.5s→3s→6s→12s→15s cap). 8 reintentos `justCreated=true`, 2 en grupos existentes. Captura `Request_ResourceNotFound`, `ResourceNotFound`, msg "does not exist".
- UI: RadioButtons `Auto / M365 / DL` en GroupsView header bulk. `BulkTypeChoice` property en VM. Override per-row GroupType si user elige M365/DL.

**Sprint state persistence + typeahead**:
- 15 page VMs `AddTransient` → `AddSingleton` en `App.xaml.cs`. State (query, results, selección, drafts) ahora persiste entre tabs. Settings + FirstRun siguen Transient (modales one-shot).
- `OnSearchQueryChanged` debounce 250ms en `UsersViewModel` + `GroupsViewModel`. Cancela previous con `CancellationTokenSource`. Min 2 chars antes de pegar a Graph. Snapshot guard descarta callbacks stale. Backend ya usaba startswith.
- Killed 4 instancias zombie de Grex365.exe que retenían DLLs bloqueadas — el código Singleton llevaba sin ejecutarse desde el commit anterior.

**Sprint O** — Licencias module overhaul:
- Nav rename `Salud tenant` → `Licencias`. RenamedNavTitles migration map para preferencias guardadas.
- Auto-load: `TenantHealthViewModel` inyecta `IConnectionStateMonitor`. PropertyChanged listener + `_autoLoadAttempted` flag → refresh automático al conectar Graph.
- View-driven trigger fallback: `TenantHealthView.xaml.cs` Loaded event llama `vm.TriggerLoadIfNeeded()`. Garantiza fetch cuando entras al módulo con lista vacía + Graph conectado, sin depender solo del monitor PropertyChanged.
- Search lupa: TextBox con glyph search + `×` clear button. `LicenseFilter` property + `LicensesView.Filter` predicate match FriendlyName + SkuPartNumber + CategoryLabel case-insensitive. `OnLicenseFilterChanged` refresca la view.
- Gestionar button → `GoToUsersCommand` resuelve `MainViewModel` singleton, switch a Usuarios.
- Header refactor a eyebrow pattern (`LICENCIAS / Licencias del tenant`).

**Estado**: build clean 7 projects 0 errors. 469 tests verdes (386 Core + 83 App). Commits pushed origin/grex365-2.0 (8 commits sesión tarde).

### 2026-05-23 (Sprint L) — Visual overhaul futurista (light + dark)

User directiva: dirección visual obligatoria — minimalista futurista premium tecnológico, glass surfaces, accent gradient azul→púrpura, cyan/neon hints, animaciones suaves, microinteracciones, depth premium. NO corporate gris muerto.

**Paleta nueva (`App.xaml`)** — coherente light+dark:
- `BrandAccentStart #4F8CFF` (azul eléctrico) · `BrandAccentMid #6A6CFF` · `BrandAccentEnd #9B6CFF` (púrpura) · `BrandCyan #22D3EE` (neon hint).
- `BrandAccentGradient` (diagonal 0,0→1,1) + `BrandAccentGradientSoft` (alpha 0.18) + `BrandAccentGradientVertical` para rails.
- Override de tokens wpf-ui `AccentFillColor*` y `AccentTextFillColor*` → toda la app hereda el azul eléctrico (botones Primary, items seleccionados).
- `GlassBorderHighlight` LinearGradientBrush (accent → púrpura alpha) para bordes sutiles iluminados.
- DropShadowEffects: `AccentGlowEffect` (blur 18, opacity 0.55), `AccentGlowSoftEffect` (blur 10, opacity 0.35), `CyanGlowEffect`, `CardElevation` (blur 24 ShadowDepth 2), `CardElevationSoft` (blur 12).

**Estilos globales actualizados**:
- `Card`: CornerRadius 10→12, añadido `CardElevationSoft` como Effect por defecto.
- `CardHover`: Storyboard MouseEnter→TranslateTransform.Y=-2 (180ms CubicEase) + GlassBorderHighlight como borde + CardElevation profunda. MouseLeave revierte.
- `MetricCard`: TranslateY=-3 + AccentGlowSoftEffect en hover.
- `HeroCard` NUEVO: GlassBorderHighlight border + AccentGlowSoftEffect por defecto. Para landing/featured cards.
- `NavListBoxItem`: rail con `BrandAccentGradientVertical` + glow + Opacity 0→1, animación TranslateX +2 en hover (150ms). Tooltip pasado a Español ("Requiere conexión").
- `ProgressBar`: indicator usa `BrandAccentGradient` + glow soft + CornerRadius 4.
- `SidebarBorder`: gradient overlay vertical alpha azul→púrpura (#0C / #08) sobre wpf-ui surface.
- `StatusBarSurface` NUEVO: borde superior `GlassBorderHighlight` (línea acento sutil).
- `BrandLogoGlyph` NUEVO: glyph con `BrandAccentSolid` + glow.

**MainWindow shell**:
- Logo del sidebar refactor a Border 36×36 con `BrandAccentGradient` background + glyph blanco + glow. Subtitle "M365 toolkit" con divisor cyan dot (glow).
- Status bar conectores: ellipses reemplazados por Border CornerRadius=5 con accent glow soft. Labels FontWeight=SemiBold para legibilidad.
- Page transition refactor: fade 280ms + TranslateY 10→0 + ScaleX/Y 0.985→1 (todo 320ms CubicEase). RenderTransformOrigin 0.5,0.4. Más sustancial que el anterior 180ms simple.

**Dashboard hero**:
- Header de página: Border 44×44 CornerRadius=12 con `BrandAccentGradientSoft` bg + `GlassBorderHighlight` border + `AccentGlowSoftEffect`. Glyph accent inside. Title 26px + subtitle inline.
- Card de acciones rápidas convertida a `HeroCard` (glow base permanente).

**Estado**: build clean 7 projects 0 errors. 469 tests siguen verdes (no rotos por refactor visual). Cards en todas las vistas heredan glow/elevation automáticamente vía `{StaticResource Card}`.

**Sprint L extension** — Hero headers en TODAS las views:
- Script Python `hero-headers.py` (efímero) reemplaza el `TextBlock Text=glyph FontSize=22 Opacity=0.85` por Border 44×44 hero badge (BrandAccentGradientSoft + GlassBorderHighlight + AccentGlowSoftEffect) en 12 vistas: ConnectView/AuditView/AuditLogView/TenantHealthView/SharedMailboxView/MailFlowRulesView/MailboxRulesView/OnboardingView/OffboardingView/CertWizardView/DomainCheckView/PsConsoleView.
- UsersView + GroupsView con patrón header diferente (sin glyph TextBlock previo): wrap inline con `Edit` tool — añade el hero badge antes del título.
- Resultado: 14 views con header hero badge coherente. Dashboard sirve de exemplar (badge + title + subtitle).

Pendiente iteración usuario futura: tune intensidad glow según feedback, focus-ring accent en TextBox/ComboBox, headers más grandes en hero pages tipo Dashboard.

### 2026-05-23 (Sprint K) — Audit JSON export + baseline diff

**Sprint K** — Machine-readable export + drift detection contra baseline previo:
- `AuditReportJsonBuilder` puro en `Grex365.Core.Audit`: serializa `AuditReportEnvelope` (Schema + Title + GeneratedAt + Tenant + Operator + Counts + Findings) via `System.Text.Json`. Constante `SchemaVersion = "grex365.audit.v1"` para automation forward-compat. Helper `Parse(string) → AuditReportEnvelope?` para reload. Encoder `JavaScriptEncoder.UnsafeRelaxedJsonEscaping` evita escapar acentos UTF-8.
- `AuditReportCounts` record (Errors/Warnings/Info/Total) precomputado en build, NO recalculado en consumer.
- `AuditBaselineComparer` puro: `Compare(baseline, current) → AuditBaselineDiff` (New/Resolved/Persistent). Identidad por `(Category, Identity, Detail, Severity-case-insensitive)` via custom `IEqualityComparer<AuditFinding>`. Dedupe rows duplicadas en current. `HasChanges` flag + count helpers.
- `AuditViewModel`: `ExportFindingsJsonCommand` (SaveFileDialog .json, UTF-8 sin BOM para tooling compatibility) + `LoadBaselineCommand` (OpenFileDialog, calcula diff contra `FindingsView` actual, status string `nuevos: X · resueltos: Y · persistentes: Z`) + `ClearBaselineCommand`. Props `BaselineSummary` + `HasBaseline` para UI.
- `AuditView.xaml`: botones "Exportar JSON..." / "Cargar baseline..." / "Limpiar baseline" (visible solo si HasBaseline) en toolbar acciones.
- 13 tests `AuditReportJsonBuilderTests`: roundtrip preserva findings, schema embebido, counts case-insensitive, ISO8601 timestamps, pretty-print default, custom options, parse null/whitespace/invalid, tenant+actor.
- 12 tests `AuditBaselineComparerTests`: empty inputs, nulls, persistent vs new vs resolved, severity case-insensitive, category case-sensitive, dedup duplicates, mixed state partition, null severity = empty, different detail = different finding.

**Estado**: 469 tests verdes (386 Core +25 + 83 App). Build clean 7 projects 0 errors.

### 2026-05-23 (Sprint J) — Audit HTML report export

User feedback "continúa" + memoria `feedback_keep_shipping`: trabajar backlog autónomo. Audit técnico previo: backlog `panel_overhaul_pending` ya **completamente shipped** en sesión 2026-05-22 (líneas PROGRESS 178-183) — memoria estale borrada del index.

**Sprint J** — Audit HTML report builder (stakeholder-friendly):
- `AuditReportHtmlBuilder` puro en `Grex365.Core.Audit`: `Build(IEnumerable<AuditFinding>, AuditReportContext) → string`. Standalone HTML con CSS embebido (light + `prefers-color-scheme: dark`), pills resumen (ERROR/WARN/INFO/TOTAL), tablas por severidad con conteos, escape HTML completo (`<`, `>`, `&`, `"`, `'`), filas omitidas si severidad vacía, empty state "Sin hallazgos".
- `AuditReportContext` record: Title + GeneratedAt + TenantDomain (opcional) + GeneratedBy (opcional). Omite metadatos si null/whitespace.
- `Esc(string?)` público (testeable) — usa StringBuilder con switch por char para perf.
- `AuditViewModel.ExportFindingsHtml` command: respeta filtros (FindingsView), SaveFileDialog .html, persiste UTF-8 con BOM, log Ok en UiLogSink. Inyecta `IGraphConnection?` opcional para resolver TenantId en el header.
- `AuditView.xaml` botón "Exportar HTML..." junto al CSV existente.
- 18 tests `AuditReportHtmlBuilderTests`: empty/null findings, group by severity, hide section if empty, HTML entity escape theory (4 inputs), pills 4 counts, tenant+actor opcionales, severidad case-insensitive, timestamp invariant format, columns ×3, Esc helper edge cases, unknown severity excluded, embeds CSS dark.

**Estado**: 444 tests verdes (361 Core +18 + 83 App). Build clean 7 projects 0 errors.

PROGRESS backlog autónomo agotado: docs internas (necesita petición explícita), MSIX assets definitivos (necesita arte), QA escenarios reales (no automatizable sin tenant real). Trabajo futuro autonomous lo deja a discreción del usuario.

### 2026-05-23 (Sprint H + I) — Theme auto-system + AppReg refactor

**Sprint H** — Theme auto-from-system:
- `ISystemThemeProvider` en `Grex365.Core.Abstractions` (`bool IsDarkTheme()`).
- `WindowsRegistryThemeProvider` en `Grex365.App.Services` lee `HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme` (DWORD 0=dark/1=light). Default dark si registry inaccesible.
- `SettingsViewModel.SystemThemeProvider` static settable + nuevo `ResolveActualTheme()` helper. `ApplyTheme` handle "Auto" via provider.
- `App.OnStartup` subscribe `Microsoft.Win32.SystemEvents.UserPreferenceChanged`: cuando user cambia tema en Windows Personalization, si pref="Auto" re-aplica via dispatcher (otherwise respect explicit). Unsubscribe en `OnExit`.
- `SettingsWindow` ComboBox + `FirstRunWizardWindow` RadioButton: añade "Auto (seguir sistema)" option.
- 8 tests `SettingsViewModelThemeTests`: ResolveActualTheme Dark/Light/Auto+darkprov/Auto+lightprov/null prov defaults/null pref/unknown/case-insensitive. IDisposable fixture restaura SystemThemeProvider.

**Sprint I** — GraphAppRegistrationService refactor:
- Service crítico (auto-crea AppReg con 9 Graph AppRoles + Exchange.ManageAsApp) tenía 0 tests + lógica mezclada SDK construction + Graph calls.
- Nueva `AppRegistrationSpec` static en `Grex365.Core.Connections`: pure helpers SDK-typed (`BuildRequiredResourceAccess`, `BuildApplication`, `BuildAdminConsentUrl`, `BuildCertLabel`) + tablas roles + validation. Sin dependencia `IGraphConnection`.
- Service trimmed 120→50 líneas, usa spec para object construction, conserva Graph SDK calls + progress + result mapping.
- 17 tests `AppRegistrationSpecTests`: shape (2 entries Graph+EXO), counts (9+1 roles), Type="Role", IDs únicos parseables GUID, permisos críticos presentes, URL formatter theory + reject empty, BuildApplication shape + validations (empty/null/empty-array), BuildCertLabel truncate.

Tests **426 verdes** (343 Core +17 + 83 App). Build clean 7 projects.

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
- [x] **Terminal PowerShell embebido** — cerrado vía módulo "Consola PS" (Sprint E 2026-05-22): REPL multi-line en runspace compartido con la app, reusa `IPowerShellRunner` (RunspacePool-backed), history navegable, Ctrl+Enter run, Esc cancel. Menos invasivo que `EasyWindowsTerminalControl` integration y reutiliza infra existente
- [x] **Theme toggle desde sidebar** — botón "Tema" junto a "Ajustes" persiste y aplica al instante
- [x] **Disable nav items cuando Graph/Exchange desconectado** — `NavigationItem.RequiresGraph/RequiresExchange`, `MainViewModel.UpdateNavEnabledStates` reactivo al `ConnectionStateMonitor`
- [x] **Focus-ring accent en TextBox/PasswordBox/ComboBox** — borde `BrandAccentSolid` + `AccentGlowSoftEffect` en IsKeyboardFocused/IsKeyboardFocusWithin trigger (App.xaml)

---

## Tests (532 passing)

**Core.Tests**: 401 · **App.Tests**: 131. Suites destacadas (Core abajo):

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
| InMemoryAuditFindingsStore | 5 | initial nulls + 0, Update sets all five fields + raises 5 PropertyChanged, Update twice overwrites, accepts zero counts |
| UserDetailsHost | 10 | initial closed+null, RequestOpen sets state + fires events, whitespace/null no-op theory, idempotent (value-gated), RequestClose from open/closed, open→close→open cycle |
| AppRegistrationSpec | 17 | shape, counts, Type=Role, IDs únicos GUID, permisos críticos, URL formatter, BuildApplication validations, BuildCertLabel truncate |
| AuditReportHtmlBuilder | 18 | empty/null findings, group by severity, hide section if empty, HTML entity escape, pills 4 counts, tenant+actor, severity case-insensitive, timestamp invariant, columns ×3, Esc edge cases |
| AuditReportJsonBuilder | 13 | roundtrip preserves findings, schema embedded, counts case-insensitive, ISO8601, pretty-print, custom options, parse invalid, tenant+actor |
| AuditBaselineComparer | 12 | empty inputs, nulls, persistent vs new vs resolved, severity case-insensitive, category case-sensitive, dedup duplicates, mixed state partition, null severity, different detail |

**App.Tests (131)** — VMs + UI helpers con mocks:

| Suite | Tests | Cubre |
|-------|-------|-------|
| UserDetailsViewModel | 16 | OpenRequested populate, userNotFound, Reset, ToggleAccount, RemoveLicense, AssignSku, ResetPassword (clipboard+show), RevokeSessions, RemoveAllLicenses, Close |
| Users | 15 | RBAC denied, confirm-no/yes paths, DisableEnable, AssignLicense seat-available/no-seats, RemoveLicenses zero/positive |
| Groups | 5 | RemoveMember confirm paths + service-throws |
| Offboarding | 6 | RBAC, confirm paths, per-flag execution |
| Onboarding | 5 | OnboardingRun confirm paths + trim UPN + uppercase usageLocation + GroupIdentifiers split, AddSku dedup + Remove |
| MailboxRules | 9 | ApplyForwarding empty/RBAC/confirm paths con trim, ClearForwarding, RemoveCalendarPermission |
| PsConsole | 9 | empty input no-op, valid script appends + history, errors marker, exception marker, cancel, Clear, history dedupe + nav up/down wrap |
| FirstRunWizard | 10 | state machine, persistencia, skip, finish con trim, save error, labels |
| SettingsViewModelTheme | 8 | ResolveActualTheme Dark/Light/Auto+dark/Auto+light/null prov/null pref/unknown/case-insensitive |
| NavTitleMigrator | 24 | null/empty/whitespace, 8 renames theory, case-insensitive, no-renamed pass-through, RenameMap invariants (contains-keys + new-resolve-to-self + no-cycles) |
| LicenseFilterMatcher | 11 | null/empty/whitespace, match por friendly/sku/category, case-insensitive, trim, short-circuit, null-card ArgumentNullException |
| LicenseCard | 13 | From factory known/unknown SKU, zero-enabled, half-used, utilization levels boundaries (60/79/80/94/95/100), over-consumed Available clamp, seat counts, priority |

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

**Backlog autónomo restante (todos blocked por entrada externa):**
1. **Asset definitivo MSIX** — reemplazar PNG placeholders por branding (requiere arte definitivo de Andersen) + smoke test instalación end-to-end con cert real (requiere entorno test con cert firmado)
2. **QA escenarios reales** — 100+ ops simultáneas bajo carga + scripted bulk CSV grande (requiere tenant real con datos representativos)

**Status fases vs Plantamiento:**
- Fase 1-4: DONE
- Fase 5 (packaging): MSIX scaffold + CI release job DONE — pendiente assets definitivos + smoke test real
- Fase 6 (telemetría + features enterprise): 7/8 done. Falta solo QA escenarios reales (no automatizable sin tenant real)

**Backlog Plantamiento §6 cerrado:** Terminal PowerShell embebido shipped via Consola PS module. Toda feature útil del backlog también shipped.
