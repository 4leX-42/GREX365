# GREX365 v2.0 — RUNBOOK (manual de operación)

> Manual operativo para sysadmins que despliegan, usan o dan soporte sobre GREX365 v2.0.
> Mapeo arquitectónico: [`ARCHITECTURE.md`](ARCHITECTURE.md). Estado fase: [`../PROGRESS.md`](../PROGRESS.md).
> Última revisión: 2026-05-23.

---

## 1. Audiencia y propósito

Este RUNBOOK cubre **uso operacional** de la herramienta una vez instalada: cómo conectarse, qué hace cada módulo, dónde están los datos, qué hacer cuando algo falla.

Si buscas:
- Arquitectura técnica / decisiones de stack → [`ARCHITECTURE.md`](ARCHITECTURE.md).
- Empaquetado / firma / distribución → [`../PACKAGING.md`](../PACKAGING.md).
- Mapeo legacy PS → nuevo .NET → [`MIGRATION.md`](MIGRATION.md).
- Estado de fases del Plantamiento → [`../PROGRESS.md`](../PROGRESS.md).

---

## 2. Instalación

### 2.1 MSIX vía AppInstaller (recomendado para producción)

Distribución típica en Andersen ES: Intune Win32 → MSIX → AppInstaller.

1. El admin sube `Grex365_X.Y.Z_x64.msix` + `Grex365.appinstaller` firmados a un feed HTTPS (Azure Blob / share UNC).
2. Cliente Windows 10/11 instala vía el link `.appinstaller`. AppInstaller verifica firma + comprueba updates al arrancar.
3. Updates: subir nuevo `.msix` + bump versión en `.appinstaller`. Cliente actualiza al siguiente arranque.

Requisitos cliente:
- Windows 10 1809+ o Windows 11.
- Trusted Publishers debe incluir el cert con el que se firmó (Intune lo despliega).
- No requiere .NET runtime preinstalado (MSIX lleva runtime self-contained).

### 2.2 Portable single-file (ad-hoc / soporte)

Para uso puntual o troubleshooting:

```powershell
# Descargar Grex365.exe desde GitHub Releases
.\Grex365.exe
```

Ventajas: no requiere admin local, no se instala, deja huella mínima.
Limitaciones: sin auto-update; arranque +3-5 s extra al descomprimir a `%TEMP%`.

### 2.3 Build desde código fuente

Para devs / QA:

```powershell
git clone https://github.com/4leX-42/GREX365.git
cd GREX365
git checkout grex365-2.0     # branch activa de la rewrite
dotnet build src/Grex365.slnx -c Release
dotnet run --project src/Grex365.App -c Debug
```

Requiere: .NET SDK 10, Windows 10/11, Windows SDK 10.0.17763+ para MSIX (opcional).

---

## 3. Primer arranque

### 3.1 First-Run Wizard

Si `%LOCALAPPDATA%/Grex365/config/user_preferences.json` no existe o `FirstRunCompleted=false`, arranca el wizard modal (5 páginas):

1. **Welcome** — intro card.
2. **Connection** — RadioButton:
   - `Certificate (app-only)` — recomendado para sysadmins.
   - `Device code (interactive)` — para sesiones puntuales.
3. **Tenant Lock** — Checkbox + TextBoxes:
   - `EnforceTenantLock` (recomendado ON).
   - `ExpectedTenantId` (GUID) — bloquea conexión a cualquier otro tenant.
   - `ExpectedTenantDomain` — backup verification.
4. **Theme** — RadioButton: Dark / Light / Auto (sigue sistema).
5. **Summary** — tabla de valores elegidos + botón Finish.

`Skip` salta el wizard pero marca `FirstRunCompleted=true` (no vuelve a salir). `Finish` persiste todos los prefs.

### 3.2 Connect inicial (sin cert preexistente)

Camino recomendado: **Connect view → Device code** primero, después usar **Cert Wizard** para automatizar futuras sesiones.

1. **Conexión** view → seleccionar `Iniciar sesión con código de dispositivo`.
2. Login interactivo en navegador (cuenta Global Admin del tenant).
3. App valida acceso real con `Me` + `Organization` antes de marcar conectado.
4. Si tenant resultante no matchea `ExpectedTenantId` → conexión abortada (TenantLock).

### 3.3 Setup automatizado de App Registration

Una vez conectado vía device code:

1. **Asistente cert** view → `Crear App Registration + cert`.
2. La herramienta llama `GraphAppRegistrationService.CreateAndConfigureAsync`:
   - Crea cert self-signed en `CurrentUser\My`.
   - Crea App Registration con 9 AppRoles Graph + Exchange.ManageAsApp + Reports.Read.All.
   - Sube cert público como KeyCredential.
   - Crea ServicePrincipal.
3. Devuelve URL de admin-consent clickable. **El admin debe abrirla y aprobar consent** (no se puede automatizar).
4. Guarda `exo-app-params.json` con `AppId`, `TenantId`, `Organization`, `Thumbprint`.
5. Reinicia o pulsa `Conectar (certificado)` — debe levantar Graph + EXO sin login.

### 3.4 Auto-connect en arranques posteriores

`App.OnStartup → TryAutoConnectAsync()`:
- Si existe `exo-app-params.json` + cert válido → conecta Graph (cert) + EXO (cert).
- Si TenantLock no matchea → aborta, muestra error en status bar.
- Si cert expirado → log Warning, abre vista Asistente cert.

---

## 4. Operación día a día

### 4.1 Estado de conexión

Status bar inferior muestra dots:
- **Graph** verde-acento = conectado.
- **Exchange** verde-acento = conectado.
- Gris = desconectado / inicializando.

Hover muestra Tenant + AppId.

### 4.2 Cambio de tema

Sidebar → botón `Tema` cicla Dark/Light/Auto. Aplicación instantánea + persistida en prefs.

Modo Auto: la app subscribe `SystemEvents.UserPreferenceChanged`. Cuando cambia el tema de Windows, la app re-aplica (solo si pref="Auto"; explícito Dark/Light no se sobrescribe).

### 4.3 Restauración de tamaño/posición ventana

La app guarda `WindowWidth/Height/Left/Top/Maximized` en `user_preferences.json` al cerrar. Al arrancar restaura con guard `IsOnScreen` (verifica `SystemParameters.VirtualScreenBounds` por si el monitor secundario fue desconectado).

### 4.4 Atajos de teclado

| Atajo | Acción |
|---|---|
| `Ctrl + ,` | Settings |
| `F1` | About |
| `Ctrl + L` | Toggle panel logs |
| `Ctrl + Enter` | Run script (Consola PS) |
| `Esc` | Cancel script en ejecución (Consola PS) |
| `Up/Down` | Navegar historial (Consola PS) |

### 4.5 Cancelación de operaciones

Todos los flujos largos (audit, bulk import, etc.) tienen botón `Cancelar` que dispara `CancellationToken`. Propaga a runspaces PS (`PowerShell.Stop()`) y HTTP de Graph.

---

## 5. Módulos por uso

### 5.1 Dashboard

Landing. Quick actions hero card + summary cards de findings críticos (errores del último audit run).

### 5.2 Conexión

Panel central de conexión manual. Botones:
- `Conectar (certificado)` — usa `exo-app-params.json` actual.
- `Iniciar sesión (device code)` — interactivo, login en navegador.
- `Desconectar` — limpia tokens + EXO disconnect.
- `Comprobar módulo EXO` / `Instalar módulo` — autoinstall `ExchangeOnlineManagement` (lanza `pwsh.exe` externo para esquivar ACL de WindowsApps).

### 5.3 Licencias (renamed 2026-05-23 desde "Salud tenant")

Auto-carga al conectarse Graph (`PropertyChanged` listener + view-driven `Loaded` fallback).

- Tabla `LicenseCard` por SKU: FriendlyName, SkuPartNumber, Category (E5 / E3 / Business / Frontline / Security / etc.), Assigned/Total, utilization bar.
- Search lupa: filtra inline por FriendlyName / PartNumber / Category (case-insensitive).
- Botón `Gestionar` → salta a Usuarios con `MainViewModel.GoToUsersCommand`.

### 5.4 Usuarios

- Search con typeahead debounced (250 ms), min 2 chars. Cancela calls stale.
- Detalle de usuario (UserDetailsView modal lateral): info básica + memberships + licencias asignadas.
- Acciones (gateadas por RBAC):
  - Disable/Enable
  - Assign license (autocomplete SKU; chequeo seats disponibles)
  - Remove license
  - Reset password (genera + copia clipboard + dialog)
  - Revoke sign-in sessions
  - Remove all licenses (confirm dialog)
- **Bulk CSV**:
  - Columnas reconocidas: `Email/UPN`, `Action` (`enable`/`disable`/`remove-licenses`/`assign:<SkuPartNumber>`).
  - Validación pre-ejecución (BulkUserActionParser).

### 5.5 Grupos

- Search debounced igual que Usuarios.
- Acciones por grupo: Add member / Remove member / Export members CSV.
- **Bulk CSV**:
  - Columns: `GroupName`, `Email`, `Type` (opcional: M365/DL/Auto).
  - **Override radio**: `Auto` (detecta por column Type / GroupType) / `M365` (force) / `DL` (force).
  - `GroupName` con `@` (SMTP completo, ej. `sales@es.andersen.com`) → la app deriva displayName del local part, evita doble `@`.
  - **Graph replica retry**: tras `POST /groups`, los `members` ops reintentan con backoff exp (1.5 → 3 → 6 → 12 → 15 s cap). 8 retries en `justCreated`, 2 en grupos existentes.

### 5.6 Buzones compartidos

- Apply shared mailbox (Convert UserMailbox → SharedMailbox).
- Convert SharedMailbox → UserMailbox (habilita visibilidad en Teams).
- Manage permissions (FullAccess / SendAs / SendOnBehalf).
- Errores típicos: `Get-Mailbox not recognized` → sesión EXO caída, reconectar.

### 5.7 Reglas de buzón

- Set Out-of-Office (state on/off, internal/external message, start/end date range).
- Apply forwarding (ForwardingAddress vs ForwardingSmtpAddress; validation SMTP shape).
- Clear forwarding (resets fields y display).
- Remove calendar permission (per-user).

### 5.8 Flujo de correo

Visor read-only de `Get-TransportRule`:
- Columnas: Name, State (Enabled/Disabled/InAudit), Priority, Mode, Description.
- Filtro libre (case-insensitive).
- Gated por RequiresExchange — disabled si EXO desconectado.

### 5.9 Auditoría (12 analizadores de seguridad)

Toolbar fila 1 (Graph audits):
1. **Identidad + grupos** — stale users + grupos sin actividad reciente (D7/D30/D90/D180 selector).
2. **MFA coverage** — admins/members/guests sin MFA.
3. **CA policies** — Conditional Access weak/disabled/report-only stale.
4. **Privileged roles** — guests admin, disabled admins, 0/1/>5 GAs.
5. **App credentials** — secrets/keys expired/expiring/long-lived.
6. **Tenant defaults** — users-can-create-apps, invitesFrom=everyone, SSPR off, etc.
7. **OAuth grants** — AllPrincipals high-risk + user-consented phish-OAuth.
8. **Actividad grupos** (D7-D180 picker).

Toolbar fila 2 (EXO audits):
9. **Forwarding externo** — mailboxes forwarding fuera del tenant.
10. **Inbox rules** — BEC indicators (delete/hide/external forward + ES/EN keywords). NumberBox tope buzones.
11. **Transport rules** — forwarding/BCC/redirect externos, custom connectors, broad-scope deletes, disabled-security-keyword rules.
12. **Shared mailbox sign-in** — shared boxes con `AccountDisabled=false`.

Exports (respeta filtros activos):
- **Exportar CSV** — formato tabular para Excel.
- **Exportar HTML** — standalone con CSS embebido (light + dark via `prefers-color-scheme`).
- **Exportar JSON** — schema `grex365.audit.v1` para automation.

Baseline diff:
- `Cargar baseline...` — carga JSON previo, computa New/Resolved/Persistent contra findings actuales.
- Status string `nuevos: X · resueltos: Y · persistentes: Z`.
- `Limpiar baseline` — descarta diff actual.

### 5.10 Registro de auditoría

Visor del `audit/grex365-YYYY-MM.jsonl` (append-only log de privileged actions):
- Filtros: severity (Ok/Warn/Error), source (componente), date range.
- Métricas en header: total, error rate, last 24h, top sources, recent errors.
- Auto-refresh cuando se añade nueva línea.

### 5.11 Onboarding

Crear usuario nuevo:
- Campos: UPN, password (validation), display name, usage location (uppercase), groups (multi-select via GroupIdentifiers split).
- Validation pre-exec (`OnboardingValidator`): UPN format, password strength, usage 2-letter ISO, mail-nickname derive.
- Idempotente: detecta UPN existente.

### 5.12 Offboarding

Wizard con flags:
- Disable account
- Revoke sign-in sessions
- Convert mailbox to shared (preserva access para manager)
- Remove all licenses
- Set out-of-office message
- Add manager as FullAccess on mailbox
- Remove from all groups
- Hide from GAL

Cada flag indep; en `Ejecutar` corre secuencial + escribe resultado por flag.

### 5.13 Asistente cert

- Crear App Registration + cert (ver §3.3).
- Renovar cert existente.
- **Exportar PFX con password** — `ICertificateGenerator.ExportPfx` (validación: thumbprint existente, password no vacío).
- Borrar cert configurado (limpia `exo-app-params.json` + cert del store).

### 5.14 Comprobación DNS

DNS check externo (no toca tenant):
- MX records
- SPF (TXT lookup `v=spf1`)
- DKIM (selectors comunes: selector1 / selector2)
- DMARC (TXT `_dmarc.<domain>`)

### 5.15 Consola PS

REPL embebido en mismo runspace que la app (Graph/EXO en scope).
- Input multi-line.
- Output prefix `PS>` por bloque.
- Cancel via Esc (PowerShell.Stop()).
- History navegable Up/Down (max 50, dedupe consecutivos).
- Ctrl+Enter → Run.

⚠ **Warning banner**: ejecuta scripts en contexto pleno de la app, sin sandbox.

---

## 6. Plugins

### 6.1 Instalación

```powershell
$dst = Join-Path $env:LOCALAPPDATA 'Grex365\plugins'
New-Item -ItemType Directory -Force -Path $dst | Out-Null
Copy-Item MiPlugin.dll $dst -Force
```

Reiniciar la app — el plugin aparece en la navegación si implementa `IModule` correctamente.

### 6.2 Habilitar / deshabilitar

Settings → sección Plugins → checkbox por DLL.
Persistido en `user_preferences.json` como `DisabledPluginAssemblies` (HashSet de nombres de assembly).

### 6.3 Plugin no carga

Log típico:
```
WARN PluginLoader: Failed to load plugin "MiPlugin.dll" — System.IO.FileNotFoundException: ...
```

Causas comunes:
- Plugin built contra runtime distinto (.NET 9 en plugin, .NET 10 host).
- Dependencias del plugin no copiadas junto al DLL.
- Plugin no expone tipo `IModule` público.

Ver `samples/Grex365.SamplePlugin/README.md` para el patrón correcto de empaquetado (`CopyLocalLockFileAssemblies=false` + `ExcludeAssets=runtime`).

---

## 7. Troubleshooting

### 7.1 EXO no conecta

**Síntoma**: dot Exchange gris, log `Connect-ExchangeOnline … not recognized`.

Causas:
1. Módulo `ExchangeOnlineManagement` no instalado.
2. Cert thumbprint en `exo-app-params.json` ya no existe en el store.
3. Org domain incorrecto.

Solución:
- Connect view → `Comprobar módulo EXO` → `Instalar módulo` (lanza `pwsh.exe` externo).
- Asistente cert → renovar cert si caducó.
- Verificar `Organization` en Settings (`xxxxx.onmicrosoft.com`).

### 7.2 Cert no encontrado

**Síntoma**: log `Certificate {thumb} not found in CurrentUser\My`.

Causas:
- Cert borrado del store.
- App ejecutándose como otro usuario (cert es `CurrentUser`).

Solución: Asistente cert → `Crear App Registration + cert` (genera nuevo + actualiza `exo-app-params.json`).

### 7.3 Tenant Lock mismatch

**Síntoma**: `Tenant esperado {expected} no coincide con conectado {actual}. Conexión abortada.`

Solución:
- Si fue cambio legítimo de tenant: Settings → actualizar `ExpectedTenantId` + `ExpectedTenantDomain`.
- Si fue accidente / phishing: NO permitir conexión; reportar.

Para deshabilitar lock (lab / dev): Settings → uncheck `EnforceTenantLock`. **NO recomendado en producción**.

### 7.4 Audits Graph fallan con "Permiso faltante"

Cada audit declara su scope (ver tabla §7.2 de [`ARCHITECTURE.md`](ARCHITECTURE.md)).

**Síntoma**: log `Permiso faltante: AuditLog.Read.All. Conceder admin consent.`

Solución:
- Asistente cert ya añade los 9 AppRoles + Reports.Read.All a la nueva App Registration. Si la App Reg es legacy (creada manualmente):
  - Azure Portal → Entra ID → App registrations → tu app → API permissions → Add → Graph → Application permissions → añadir `Reports.Read.All` (o el que falte).
  - **Grant admin consent for {tenant}**.
- Verificar via Graph Explorer: `https://graph.microsoft.com/v1.0/applications?$filter=appId eq '{appId}'` → revisar `requiredResourceAccess`.

**Resiliencia**: el audit de identidad detecta `AuditLog.Read.All` faltante y degrada gracefully (retry sin `signInActivity` en `$select`) — no aborta todo el audit por un solo permiso.

### 7.5 Bulk groups: "Resource not found" tras crear grupo

**Síntoma**: log `... Request_ResourceNotFound after POST /groups`.

Causa: replica lag de Graph entre `POST /groups` y `POST /groups/{id}/members`.

Solución: ya implementada (`WithGraphReplicaRetryAsync` 8 retries con exp backoff). Si persiste tras 90s de retries → Microsoft tenant issue, abrir ticket.

### 7.6 Audit log vacío

**Síntoma**: vista Registro de auditoría sin entries.

Causas:
- Mes/año no matchea filtro. Verificar dropdown month.
- `%LOCALAPPDATA%/Grex365/audit/` no existe (primera ejecución sin ops privilegiadas aún).
- Permisos NTFS — la app no puede escribir.

Solución: comprobar `%LOCALAPPDATA%/Grex365/audit/grex365-YYYY-MM.jsonl` existe + tiene lines. Si no, lanzar una op privilegiada (e.g. assign license) y refrescar.

### 7.7 Plugin POC no aparece tras copia

Ver §6.3.

### 7.8 Theme no sigue al sistema

**Síntoma**: cambias Windows a dark/light, la app no cambia.

Causa: pref no está en `Auto`.

Solución: sidebar → botón `Tema` cicla hasta `Auto`. O Settings → Tema → `Auto (seguir sistema)`.

### 7.9 RBAC bloquea op a admin legítimo

**Síntoma**: dialog `Operación no autorizada. Tu cuenta no pertenece al grupo RBAC requerido.`

Causa: usuario no está en el grupo Entra configurado como `RbacRequiredGroup`. App-only auth (cert) bypassea RBAC por design (no hay contexto delegado).

Solución:
- Si conectado via device code: añadir UPN al grupo Entra (puede tardar minutos en propagarse — el `RbacGuard` cachea por sesión; reiniciar la app o `Invalidate()` desde Settings).
- Si quieres bypass total: cambiar a Cert auth (app-only) en Conexión.

### 7.10 App no arranca después de update MSIX

**Síntoma**: doble-click sin efecto. Event Viewer muestra `.NET runtime` exception.

Solución:
- Logs Serilog NO se escriben si app crashea antes de bootstrap. Mira Event Viewer → Applications → `.NET Runtime` source.
- Reinstalar MSIX: `Get-AppxPackage *Grex365* | Remove-AppxPackage` + reinstalar.

---

## 8. Datos en disco

```
%LOCALAPPDATA%\Grex365\
├── config\
│   ├── user_preferences.json     # tenant lock, theme, last nav, window pos/size, disabled plugins, log level
│   └── exo-app-params.json       # AppId, TenantId, Org, cert thumbprint (NO clave privada)
├── logs\
│   └── grex365-YYYY-MM-DD.log    # Serilog rotativo (30 días)
├── audit\
│   └── grex365-YYYY-MM.jsonl     # append-only privileged-action log
└── plugins\
    └── *.dll                     # plugins externos
```

### 8.1 Backup

Backup full = copiar `%LOCALAPPDATA%\Grex365\` + exportar cert desde `CurrentUser\My`:

```powershell
# Backup carpeta
Copy-Item $env:LOCALAPPDATA\Grex365 D:\backups\grex365_$(Get-Date -Format yyyyMMdd) -Recurse

# Export cert (Asistente cert → Exportar PFX con password)
# O manual:
$cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object Thumbprint -eq <thumbprint>
$pwd = ConvertTo-SecureString -String '<pwd>' -Force -AsPlainText
Export-PfxCertificate -Cert $cert -FilePath D:\backups\grex365.pfx -Password $pwd
```

### 8.2 Restore en máquina nueva

```powershell
# Restaurar carpeta
Copy-Item D:\backups\grex365_YYYYMMDD\* $env:LOCALAPPDATA\Grex365 -Recurse -Force

# Importar cert
$pwd = ConvertTo-SecureString -String '<pwd>' -Force -AsPlainText
Import-PfxCertificate -FilePath D:\backups\grex365.pfx -CertStoreLocation Cert:\CurrentUser\My -Password $pwd
```

Arrancar app → debería auto-connect.

### 8.3 Desinstalación limpia

```powershell
# Cerrar app
Get-Process Grex365 | Stop-Process -Force

# Borrar datos
Remove-Item $env:LOCALAPPDATA\Grex365 -Recurse -Force

# Borrar cert
Get-ChildItem Cert:\CurrentUser\My | Where-Object Subject -like 'CN=Grex365*' | Remove-Item

# Desinstalar (MSIX)
Get-AppxPackage *Grex365* | Remove-AppxPackage
```

---

## 9. Configuración avanzada

### 9.1 Nivel de logging

Settings → Logging level: `Debug` / `Information` / `Warning` / `Error`.
Aplicado vía `LoggingLevelSwitch` — afecta inmediatamente, sin reiniciar.

Debug recomendado solo para troubleshooting (gran volumen, retention 30 días).

### 9.2 Application Insights (telemetría)

Empty/null → `NullTelemetry` (no-op).

Para habilitar:
```json
// %LOCALAPPDATA%\Grex365\config\user_preferences.json
{
  "AppInsightsConnectionString": "InstrumentationKey=...;IngestionEndpoint=..."
}
```

O env var `APPLICATIONINSIGHTS_CONNECTION_STRING` (sobreescribe pref).

Reiniciar app. `UiLogSink` empezará a forward Ok/Warn/Error a AppInsights (`TrackEvent` / `TrackException`).

### 9.3 Tenant Lock

Settings → checkbox `Enforce Tenant Lock`. Off para lab/dev.

### 9.4 Plugins enable/disable

Settings → Plugins → toggle por DLL. Persiste en `DisabledPluginAssemblies` (HashSet de nombres de assembly sin extensión).

Cambios efectivos al siguiente arranque (plugins se cargan en `App.OnStartup`).

### 9.5 Proxy

Si tu red exige proxy:

```powershell
# Antes de lanzar la app
$env:HTTPS_PROXY = 'http://proxy.local:8080'
$env:HTTP_PROXY  = 'http://proxy.local:8080'
$env:NO_PROXY    = 'localhost,127.0.0.1'
Start-Process Grex365.exe
```

.NET HttpClient (Graph SDK) y MSAL respetan estas vars.

---

## 10. Telemetría + privacidad

### 10.1 Qué se recolecta

Con `ApplicationInsightsTelemetry` activado:
- `UiLog.Ok` events — operaciones exitosas. Propiedades: source (componente), action.
- `UiLog.Warn` events — issues recoverables.
- `TrackException` — excepciones manejadas + unhandled (via global handler).
- Default AppInsights auto-collection: trace + dependency calls + page views (off en desktop).

### 10.2 Qué NO se recolecta

- Contenido de mensajes / mailboxes.
- Datos personales de usuarios del tenant (UPN, email).
- Claves / cert thumbprints.
- Contenido de CSVs bulk.

### 10.3 Opt-out

Empty `AppInsightsConnectionString` → `NullTelemetry`, cero IO de telemetría.

Logs Serilog locales siguen activos para troubleshooting (no salen del equipo).

---

## 11. Soporte + escalación

### 11.1 Antes de abrir ticket

Recoge:
1. Build/version (About → F1).
2. Logs últimos 24h: `%LOCALAPPDATA%\Grex365\logs\grex365-YYYY-MM-DD.log`.
3. Audit log mes actual: `%LOCALAPPDATA%\Grex365\audit\grex365-YYYY-MM.jsonl`.
4. `user_preferences.json` (anonimiza TenantId si es sensible).
5. Steps to reproduce.

### 11.2 Canales

- Internal Andersen ES IT: ver canal #it-soporte / mail `support@es.andersen.com`.
- Bugs reproducibles: GitHub Issues `4leX-42/GREX365`.

### 11.3 Logs sensibles

Antes de compartir:
- Audit log puede contener UPNs de usuarios → anonimizar si va fuera del equipo.
- Logs Serilog raramente contienen PII (logging estructurado solo añade campos explícitos), pero revisa antes.
- `exo-app-params.json` contiene AppId + TenantId — share OK con MS Support (públicos) pero NO publicar en foros.

---

## 12. Limitaciones conocidas

- **No multi-tenant**: una App Reg / un Tenant Lock por instalación. Para gestionar múltiples tenants, usar múltiples user profiles Windows (cada uno con su `%LOCALAPPDATA%\Grex365`).
- **App-only auth bypassea RBAC** by design (no hay `me.CheckMemberGroups` context).
- **MSIX firma**: el `Publisher` del manifest debe matchear el `Subject` del cert que firma. Si cambias de cert (renew), re-firmar manifest.
- **Auto-update MSIX**: requiere feed HTTPS accesible. Sin feed, el `.appinstaller` no actualiza.
- **Plugins**: cargan in-proc con full perms — no hay sandbox. No instalar plugins de fuentes no confiables.
- **Audit Reports.Read.All**: report endpoints de Graph tienen ventana de 1-2 días lag. El audit de actividad de grupos refleja datos de hace 24-48h, no en vivo.

---

## 13. Cambios recientes (changelog operacional)

Ver [`../PROGRESS.md`](../PROGRESS.md) "Bitácora sesiones" para detalle por commit. Resumen entries 2026-05:

- **2026-05-23 Sprint O** — "Salud tenant" → "Licencias" + auto-load + search filter + deep-link a Usuarios.
- **2026-05-23 Sprint N** — Bulk groups: SMTP-as-GroupName + Graph replica retry + RadioButtons M365/DL/Auto.
- **2026-05-23 Sprints L-M** — UX overhaul (futurista + restraint), hero badges 14 views, palette refinada.
- **2026-05-23 Sprint K** — Audit JSON export + baseline diff.
- **2026-05-23 Sprint J** — Audit HTML report (stakeholder-friendly).
- **2026-05-23 Sprint H** — Theme auto-from-system (Windows live follow).
- **2026-05-22 Sprint G** — First-Run Wizard + DataGrid theme real-fix.
- **2026-05-22** — 7 security audits nuevos (CA, Privileged, AppCred, TenantDefaults, OAuth, Transport, SharedSignIn).
- **2026-05-20** — Mail Flow viewer.

---

## 14. Referencias

- [`ARCHITECTURE.md`](ARCHITECTURE.md) — arquitectura técnica.
- [`../PACKAGING.md`](../PACKAGING.md) — empaquetado MSIX/AppInstaller/Intune.
- [`../PROGRESS.md`](../PROGRESS.md) — estado fase + bitácora sesiones.
- [`MIGRATION.md`](MIGRATION.md) — mapeo legacy PS → nuevo .NET.
- [`ROADMAP.md`](ROADMAP.md) — punch list H0-H6.
