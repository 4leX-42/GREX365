# GREX365

Toolkit PowerShell para operaciones administrativas masivas en Microsoft 365. Input mínimo, resolución automática, ejecución guiada

---

## Módulos

| Módulo | Función |
|--------|---------|
| `INYECCIÓN DE USUARIOS // 365-DL` | Alta masiva en Microsoft 365 Groups o Distribution Lists desde CSV |
| `EXTRACCIÓN DE USUARIOS // EMAIL-ID` | Exportación de miembros a CSV (`Email` + `Object ID`) |
| `CREAR GRUPOS / DL DESDE CSV` | Creación masiva de M365 Groups o Distribution Lists desde CSV |
| `CORREGIR SHAREDMAILBOX → USERMAILBOX (TEAMS)` | Conversión de SharedMailbox a UserMailbox para habilitar Teams |
| `ASISTENTE DE CREACIÓN DE CERTIFICADO (ExO + Graph)` | Wizard guiado para autenticación por certificado en ExO y Graph |
| `PREFERENCIAS / MÉTODO DE CONEXIÓN` | Gestión del método de autenticación (tradicional / certificado) y configuración |

---

### INYECCIÓN DE USUARIOS // 365-DL

Alta masiva de miembros sobre un grupo destino desde CSV

**Lógica interna**
- Resuelve `Object ID` desde `Email` si no está presente
- Detecta el tipo de grupo (M365 Group / Distribution List) y selecciona el método de alta correspondiente
- Acepta CSVs mínimos (una sola columna `Email` es suficiente)
- Soporta usuarios sin licencia en Distribution Lists (si el objeto existe en el tenant)
- Validación de formato y consistencia previa a la ejecución

---

### EXTRACCIÓN DE USUARIOS // EMAIL-ID

Exporta todos los miembros de un grupo origen a CSV estructurado (`Email`, `Object ID`)

La salida es directamente reutilizable para auditoría, inventario o reinyección

---

### CREAR GRUPOS / DL DESDE CSV

Creación masiva de grupos en el tenant desde CSV

**Lógica interna**
- Soporta M365 Groups y Distribution Lists en la misma ejecución
- Lectura de definiciones desde CSV (nombre, alias, tipo, descripción, owners)
- Detección de duplicados previos a la creación
- Asignación de propietarios y miembros iniciales si vienen en el CSV
- Reporte final con creados / omitidos / errores

---

### CORREGIR SHAREDMAILBOX → USERMAILBOX (TEAMS)

Convierte buzones SharedMailbox a UserMailbox para habilitar acceso a Teams

**Lógica interna**
- Detecta el tipo actual del buzón antes de convertir
- Conversión vía Exchange Online sin pérdida de contenido
- Útil cuando un usuario migrado desde SharedMailbox necesita licencia Teams
- Validación post-conversión

---

### ASISTENTE DE CREACIÓN DE CERTIFICADO (ExO + Graph)

Wizard guiado para configurar autenticación por certificado contra Exchange Online y Microsoft Graph

**Lógica interna**
- Genera certificado autofirmado en `CurrentUser\My` (clave privada nunca sale del equipo)
- Exporta `.cer` público para subir a la App Registration
- Registra parámetros (`AppId`, `TenantId`, `Organization`, `Thumbprint`) en `config/exo-app-params.json`
- Pasos guiados según `cert_instrunciones/EXO_Cert_Auth_Pasos.csv`
- Detecta configuración previa y permite rehacerla
- Eliminación destructiva disponible desde el menú de Preferencias (doble confirmación)

> El `.cer` y la clave privada **no se almacenan en el repositorio** (excluidos por `.gitignore`)

---

### PREFERENCIAS / MÉTODO DE CONEXIÓN

Gestión central del método de autenticación y configuración del toolkit

**Opciones**
- Cambio entre método **tradicional** (UPN + login interactivo) y **certificado** (no interactivo)
- Configuración del UPN del administrador tradicional
- Reset de preferencias (vuelve al flujo de primera ejecución)
- Eliminación del certificado configurado (acción destructiva con doble confirmación)

---

## Instalación

```powershell
# Clonar
git clone https://github.com/4leX-42/GREX365.git
cd GREX365
```

O descargar el ZIP desde **Code → Download ZIP** y extraer en local

---

> Probado en PowerShell 5.1+ .7 con los módulos ExchangeOnline y AzureAD/Graph
