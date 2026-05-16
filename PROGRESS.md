# GREX365 v2.0 — Progress log

Branch: `grex365-2.0` · Pushed up to `origin/grex365-2.0`.
Última actualización: 2026-05-16.

> Última feature: Bulk creation de grupos M365 desde CSV (Email + GroupName forward-fill).

## Done

### Plataforma
- [x] Scaffold .NET 10 (`Grex365.Core`, `Grex365.App` (WPF), `Grex365.PowerShell`, `Grex365.Core.Tests`)
- [x] MVVM con CommunityToolkit.Mvvm
- [x] DI con `Microsoft.Extensions.Hosting`
- [x] Serilog → archivo rotativo en `%LOCALAPPDATA%\Grex365\logs\`
- [x] Global exception handlers (UI / AppDomain / TaskScheduler)
- [x] Legacy config importer (importa `GREX365/config/*.json` heredado)
- [x] Runspace pool (`RunspacePoolHost`) y `IPowerShellRunner` con streams

### Conexión M365
- [x] `IGraphConnection` (cert-based) + smoke test contra `/organization`
- [x] `IExchangeConnection` (cert-based) vía `Connect-ExchangeOnline`
- [x] `IConnectionStateMonitor` con polling 2s y `INotifyPropertyChanged`
- [x] `ICertValidator` (existencia + validez del thumbprint)
- [x] `ITenantLock` — bloquea conexión si tenant ID no coincide
- [x] Disconnect Graph + EXO

### Módulos UI (10 en navegación)
- [x] Dashboard — status cards + quick action buttons
- [x] Conexion — cert auth Graph + EXO
- [x] Salud tenant — org + counts + SKUs consumidos
- [x] Usuarios — buscar, perfil, membresías, enable/disable, quitar licencias, **asignar licencia (picker SKU)**, bulk CSV
- [x] Grupos — buscar, miembros, añadir (texto/CSV), eliminar, exportar CSV, **bulk create M365 desde CSV (forward-fill GroupName)**
- [x] Buzones — lookup + permisos actuales, convertir Regular↔Shared, FullAccess/SendAs/SendOnBehalf, CSV import/export
- [x] Auditoria — identidades (stale members/guests + disabled+licensed) + grupos (sin owner / vacíos), paralelizado 8x
- [x] Offboarding — wizard compuesto (deshabilitar + quitar licencias + convertir a shared)
- [x] Cert Wizard — generar self-signed RSA 2048, instalar en CurrentUser\My, exportar .cer
- [x] DNS check — nslookup MX/TXT/SPF/DMARC (no requiere auth)
- [x] Settings (modal) — tenant lock, cert config con picker, tema persistido

### UX/QoL
- [x] Sidebar nav con persistencia del último seleccionado
- [x] Status bar global (Graph/EXO/Tenant/Cuenta + botón Desconectar todo)
- [x] Log panel con filtros por severidad + Limpiar
- [x] Cancellation tokens + ProgressRing en toda operación larga
- [x] MessageBox confirm en operaciones destructivas (disable, remove licenses, remove member, offboarding)
- [x] Tema Dark/Light persistido en `UserPreferences.Theme`
- [x] Cert picker dialog desde Settings (lista certs `CurrentUser\My`)

### Tests (89 passing)
- [x] LogEntry, PreferencesStore, PowerShellRunner, CertValidator
- [x] LegacyPreferencesImporter
- [x] TenantLock (5 escenarios)
- [x] SharedMailboxService (12: apply, convert, get permissions, errores)
- [x] FlexibleCsvReader (8: delimitadores, quoted, BOM, edge cases)
- [x] ConnectionStateMonitor (4: estado inicial, polling, fallos, dispose)
- [x] IdentityAuditAnalyzer (9: stale, disabled+lic, totales)
- [x] OffboardingService (6: empty UPN, missing user, per-flag, errores)
- [x] SkuInfo (6: math available, ordering, display, fallback guid)
- [x] BulkGroupRowPreprocessor (13: forward-fill, skip orphans, trim, IsEmail theory)

## Pending

### Funcional
- [ ] Auth tradicional/UPN interactivo (MSAL)
- [ ] Set-OutOfOffice / forwarding rules
- [ ] Calendar permission view/set
- [ ] Mail flow rules viewer
- [ ] Audit: groups without recent activity, externos en grupos privados, etc.
- [ ] Bulk groups CSV: soportar también DL (Exchange Online — script legacy lo hace)
- [ ] Cert export PFX con password
- [ ] Auto-update App Registration permisos via Graph (legacy CertWizard hace los 29 pasos)

### Roadmap doc (fases)
- [ ] **Fase 4 — Plugins/MEF**: arquitectura modular dinámica
- [ ] **Fase 5 — Packaging**: MSIX, firma código, AppInstaller, Intune
- [ ] **Fase 6 — Telemetría**: AppInsights, audit DB, métricas

### Polish UI
- [ ] Charts/gráficos (TenantHealth podría tener pie chart de SKUs)
- [ ] Terminal PowerShell embebido (`EasyWindowsTerminalControl` o similar)
- [ ] Theme toggle accesible desde título / dashboard (ahora solo en Settings)
- [ ] Disable nav items cuando Graph desconectado (gating UX)
- [ ] Toast notifications (wpf-ui Snackbar) en éxitos/fallos largos

## Por dónde voy

**Capas estables.** Próximo bloque natural sería:
1. Onboarding wizard (espejo del Offboarding: crear user + asignar licencia + agregar a grupos)
2. Set-OutOfOffice / forwarding rules en buzones
3. Bulk groups CSV — extender a DL via EXO runner

**Por validar:**
- Render visual real (no he podido lanzar la UI desde sesión headless)
- Conexión M365 real (necesita tenant + cert reales)
- Comportamiento offline de cada vista (qué pasa al pulsar Buscar sin conexión — debería mostrar "Graph no está conectado.")

## Cómo lanzar

```powershell
dotnet run --project src/Grex365.App/Grex365.App.csproj
```

Primer arranque: 5-15s para JIT. Después abre `MainWindow` (FluentWindow Mica, 1280x800).

Datos persistidos en `%LOCALAPPDATA%\Grex365\`:
- `config/preferences.json` — tenant lock, theme, last nav
- `config/exo-app-params.json` — cert config (AppId, TenantId, Org, Thumbprint)
- `logs/grex365-YYYY-MM-DD.log` — Serilog rotativo (30 días)
