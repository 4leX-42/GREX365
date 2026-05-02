# GREX365

Toolkit PowerShell para operaciones administrativas masivas en Microsoft 365. Input mínimo, resolución automática, ejecución guiada

---

## Módulos

| Módulo | Función |
|--------|---------|
| `INYECCIÓN DE USUARIOS // 365-DL` | Alta masiva en Microsoft 365 Groups o Distribution Lists desde CSV |
| `EXTRACCIÓN DE USUARIOS // EMAIL-ID` | Exportación de miembros a CSV (`Email` + `Object ID`) |
| `Coming Soon...` | — |
| `Coming Soon...` | — |

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

## Instalación

```powershell
# Clonar
git clone https://github.com/4leX-42/GREX365.git
cd GREX365
```

O descargar el ZIP desde **Code → Download ZIP** y extraer en local

---

> Probado en PowerShell 5.1+ .7 con los módulos ExchangeOnline y AzureAD/Graph
