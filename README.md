# GREX365

Toolkit para operaciones administrativas sobre Microsoft 365: Exchange Online, Microsoft Graph y Entra ID. Autenticación por certificado (app-only) o device code (delegado).

> **Estado del proyecto (2026-05)**: migración en curso de PowerShell + WPF a **C# .NET 10 + WPF + WPF-UI (Fluent)**. La versión PS sigue funcional en `GREX365/` mientras se porta.
>
> Documentación clave:
> - [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) — stack, capas, patrones, decisiones rechazadas con razones
> - [`docs/ROADMAP.md`](docs/ROADMAP.md) — punch list completo H0–H6, estado por sub-item, decisiones abiertas
> - [`docs/MIGRATION.md`](docs/MIGRATION.md) — tabla feature-by-feature legacy→nuevo
> - [`deep-research-report.md`](deep-research-report.md) — research inicial (raíz del repo)

## Build (nueva app .NET)

```bash
# Restore + build
dotnet build src/Grex365.slnx -c Release

# Run tests
dotnet test src/Grex365.slnx -c Release

# Run app
dotnet run --project src/Grex365.App -c Debug

# Publish single-file .exe (self-contained, ~70-90 MB, no .NET install required)
dotnet publish src/Grex365.App/Grex365.App.csproj `
  -c Release -r win-x64 --self-contained `
  -p:PublishSingleFile=true `
  -p:IncludeAllContentForSelfExtract=true `
  -p:EnableCompressionInSingleFile=true
```

Datos de usuario (config, logs): `%LocalAppData%\Grex365\`

---

## Legacy: PowerShell toolkit (modo previo)

```
GREX365 v2.0
─────────────────────────────────────────────
Tenant   : <tenant>.onmicrosoft.com
Auth     : Certificate
Graph    : Connected
Exchange : Connected
─────────────────────────────────────────────
```

## Capacidades

| # | Operación | Servicios | Descripción |
|---|-----------|-----------|-------------|
| 1 | Agregar miembros a grupo/DL | Graph + EXO | Alta masiva desde CSV con resolución automática de identidades, deduplicación, soporte de usuarios sin licencia. |
| 2 | Exportar miembros de grupo/DL | Graph + EXO | Exportación a CSV `Email;Id` reutilizable para auditoría o reinyección. |
| 3 | Crear grupos/DL desde CSV | Graph + EXO | Creación masiva de M365 Groups o Distribution Lists con asignación inicial de miembros. Soporta `-WhatIf`. |
| 4 | Convertir SharedMailbox a UserMailbox | EXO | Habilita visibilidad en Microsoft Teams. Confirmación post-cambio automática. |
| 5 | Asistente de certificado | Graph + Entra | Provisión automatizada en 29 pasos: cert self-signed, App Registration, Service Principal, roles directorio, AppRoles Graph + Exchange. |
| 6 | Preferencias | — | Gestión de método de conexión, UPN admin, eliminación segura de cert configurado. |

## Requisitos

- Windows 10 / 11
- **PowerShell 7.4 LTS o superior** (Windows trae 5.1 por defecto, no es suficiente)
- Conexión a internet con acceso a `*.microsoft.com`
- Módulos auto-instalados en primera ejecución (scope `CurrentUser`):
  - `ExchangeOnlineManagement`
  - `Microsoft.Graph.Authentication`
  - `Microsoft.Graph.Users`
  - `Microsoft.Graph.Groups`
  - `Microsoft.Graph.Applications`
  - `Microsoft.Graph.Identity.DirectoryManagement`
  - `Microsoft.Graph.Identity.SignIns`

Para la provisión inicial del certificado (asistente) se requiere una cuenta **Global Administrator** del tenant.

## Instalar PowerShell 7

Si solo tienes PS 5.1 (default Windows):

```powershell
winget install --id Microsoft.PowerShell --source winget
```

O descarga el MSI desde [aka.ms/powershell](https://aka.ms/powershell).

Verifica:

```powershell
pwsh -NoLogo -Command '$PSVersionTable.PSVersion'
```

Debe mostrar `7.4.x` o superior.

## Inicio rápido

```powershell
git clone https://github.com/4leX-42/GREX365.git
cd GREX365
pwsh .\Main.ps1
```

> **Importante**: usa `pwsh` (PowerShell 7), no `powershell` (5.1).

En la primera ejecución se preguntará por el método de conexión. Si eliges `Certificate` y no hay certificado configurado, se ofrecerá lanzar el asistente. Tras la primera ejecución, el método queda persistido en `GREX365/config/user_preferences.json` y se puede cambiar desde el menú con las teclas `C` / `T`.

## Autenticación

| Método | Descripción | Cuándo usar |
|--------|-------------|-------------|
| Certificate (app-only) | App Registration + cert X.509 self-signed almacenado en `CurrentUser\My`. Sin login interactivo en cada sesión. | Producción, automatizaciones, scripts desatendidos. |
| Device code (delegado) | Login interactivo con la cuenta del operador en cada sesión. | Soporte ad-hoc, pruebas, equipos compartidos. |

La clave privada nunca sale del equipo. El `.cer` público se sube automáticamente a la App Registration durante el asistente.

## Modelo de seguridad

- **Cert local**: clave privada en `CurrentUser\My` con ACL restringida al usuario actual + SYSTEM.
- **JSON metadata**: `config/exo-app-params.json` contiene `TenantId`, `AppId`, `Thumbprint`, `Organization`. Sin secretos (sin client secret, sin clave privada).
- **Gitignore**: `*.cer`, `*.pfx`, `exo-app-params.json`, `user_preferences.json` y `logs/` excluidos del repositorio.
- **Validación post-conexión**: la cabecera principal refleja en tiempo real el estado de Graph y Exchange Online.

## Permisos otorgados al SP durante el asistente

- **Exchange.ManageAsApp** (AppRole sobre Office 365 Exchange Online)
- **Microsoft Graph AppRoles**:
  - `User.ReadWrite.All`
  - `Group.ReadWrite.All`
  - `GroupMember.ReadWrite.All`
  - `Directory.ReadWrite.All`
  - `Organization.Read.All`
  - `RoleManagement.Read.Directory`
  - `UserAuthenticationMethod.ReadWrite.All`
  - `Policy.Read.All`
- **Roles directorio Entra**:
  - Exchange Administrator
  - User Administrator
  - Groups Administrator

## Estructura

```
GREX365-main_2/
├── Main.ps1                          # bootstrap + loop
├── README.md
├── .gitignore
├── docs/
│   ├── CSV-Schemas.html              # esquemas CSV soportados
│   ├── Set-ExecutionPolicy.html      # guía de ExecutionPolicy
│   └── Certificate-Setup-Steps.csv   # detalle de los 29 pasos del asistente
└── GREX365/
    ├── config/                       # gitignored — datos live
    │   ├── exo-app-params.json
    │   ├── user_preferences.json
    │   └── *.cer
    ├── logs/                         # gitignored — logs por sesión
    └── Modules/
        ├── Logging.ps1               # Write-Log + sesiones de log
        ├── Console.ps1               # cabeceras, paneles, layout
        ├── Validation.ps1            # validación de entradas
        ├── Csv.ps1                   # parser CSV robusto
        ├── Preferences.ps1           # user_preferences.json + cert config
        ├── Connection.ps1            # Graph + EXO + estado servicios
        ├── GroupResolver.ps1         # búsqueda inteligente de grupos
        ├── CertWizard.ps1            # asistente de 29 pasos
        └── Menu.ps1                  # menú principal
    └── Scripts/
        ├── Add-GroupMembers.ps1
        ├── Export-GroupMembers.ps1
        ├── New-GroupsFromCsv.ps1
        └── Convert-SharedToUserMailbox.ps1
```

## Formato de CSVs

Esquemas detallados en `docs/CSV-Schemas.html`. Detección automática de:

- **Encoding**: UTF-8, UTF-8 BOM, UTF-16 LE, UTF-16 BE.
- **Delimitador**: `;`, `,`, TAB.
- **Cabeceras**: alias flexibles para `Email`, `Id`, `GroupName`.

## Logs

Cada operación abre una sesión de log en `GREX365/logs/<operación>_<timestamp>.log`. El log se persiste si la operación genera al menos una línea OK o ERROR (no para flujos puramente informativos).

## Troubleshooting

| Síntoma | Causa probable | Solución |
|---------|----------------|----------|
| `Connect-MgGraph` falla tras crear cert | Propagación de consent | Esperar 30-60s y reintentar |
| `Get-Mailbox` no encuentra usuario | Sesión EXO no activa | Salir y relanzar; cabecera debe mostrar `Exchange: Connected` |
| ExecutionPolicy bloquea | Política restrictiva del sistema | Ver `docs/Set-ExecutionPolicy.html` |
| CSV no se lee correctamente | Encoding o delimitador raros | Ver `docs/CSV-Schemas.html` para esquemas soportados |

## Licencia

Uso interno. Sin licencia pública.
