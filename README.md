<div align="center">

# GREX365

Operaciones masivas en Microsoft 365 desde PowerShell. Input mínimo, resolución automática, ejecución guiada.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%7C%207-1f1f1f?style=flat-square&logo=powershell&logoColor=white) ![Exchange Online](https://img.shields.io/badge/Exchange_Online-1f1f1f?style=flat-square) ![Microsoft Graph](https://img.shields.io/badge/Microsoft_Graph-1f1f1f?style=flat-square) ![Cert auth](https://img.shields.io/badge/auth-certificado-c8a25a?style=flat-square)

</div>

## <img src="https://api.iconify.design/lucide/download.svg?color=%238b949e" width="18" align="top"> Instalar

```powershell
git clone https://github.com/4leX-42/GREX365.git
cd GREX365
```

O **Code → Download ZIP**. Requiere los módulos ExchangeOnlineManagement y Microsoft Graph.

## <img src="https://api.iconify.design/lucide/layers.svg?color=%238b949e" width="18" align="top"> Módulos

| | |
|---|---|
| <img src="https://api.iconify.design/lucide/users.svg?color=%238b949e" width="18" align="top"> **Inyección de usuarios** | Alta masiva en M365 Groups o DLs desde CSV (basta la columna `Email`) |
| <img src="https://api.iconify.design/lucide/file-spreadsheet.svg?color=%238b949e" width="18" align="top"> **Extracción** | Miembros de un grupo a CSV (`Email`, `Object ID`) |
| <img src="https://api.iconify.design/lucide/list-checks.svg?color=%238b949e" width="18" align="top"> **Crear grupos / DL** | Creación masiva desde CSV con propietarios, miembros y detección de duplicados |
| <img src="https://api.iconify.design/lucide/mail.svg?color=%238b949e" width="18" align="top"> **Shared → User mailbox** | Conversión para habilitar Teams, sin pérdida de contenido |
| <img src="https://api.iconify.design/lucide/key-round.svg?color=%238b949e" width="18" align="top"> **Asistente de certificado** | Autenticación app-only en ExO y Graph; la clave privada no sale del equipo |
| <img src="https://api.iconify.design/lucide/settings-2.svg?color=%238b949e" width="18" align="top"> **Preferencias** | Método de conexión (tradicional o certificado), UPN y reset |

<details>
<summary><img src="https://api.iconify.design/lucide/info.svg?color=%238b949e" width="18" align="top"> <b>Detalle de cada módulo</b></summary>

**Inyección de usuarios // 365-DL**
- Resuelve `Object ID` desde `Email` si no está presente.
- Detecta el tipo de grupo (M365 Group o Distribution List) y usa el método de alta que corresponde.
- Admite usuarios sin licencia en Distribution Lists si el objeto existe en el tenant.
- Valida formato y consistencia antes de ejecutar.

**Extracción de usuarios // Email-ID**: salida reutilizable para auditoría, inventario o reinyección.

**Crear grupos / DL desde CSV**: M365 Groups y DLs en la misma ejecución (nombre, alias, tipo, descripción, owners); informe final con creados, omitidos y errores.

**SharedMailbox → UserMailbox (Teams)**: detecta el tipo actual, convierte vía Exchange Online y valida después.

**Asistente de certificado (ExO + Graph)**
- Certificado autofirmado en `CurrentUser\My`; exporta el `.cer` público para la App Registration.
- Guarda `AppId`, `TenantId`, `Organization` y `Thumbprint` en `config/exo-app-params.json`.
- Pasos guiados según `cert_instrunciones/EXO_Cert_Auth_Pasos.csv`; detecta configuración previa.
- Borrado del certificado desde Preferencias, con doble confirmación.

> El `.cer` y la clave privada no se guardan en el repositorio (`.gitignore`).
</details>
