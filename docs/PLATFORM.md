# GREX365 como plataforma — arquitectura objetivo y hoja de ruta

Complementa a [`ARCHITECTURE.md`](ARCHITECTURE.md), que describe la aplicación de escritorio, y parte
de la auditoría de [`AUDIT-2026-11.md`](AUDIT-2026-11.md).

## 1. Principios

- **Nativo primero.** Todo funciona sin IA. La IA es un canal más.
- **Una sola implementación por capacidad.** La ventana, la línea de comandos, las tareas
  programadas y los agentes ejecutan la misma operación por el mismo camino.
- **Las reglas viven en el núcleo.** Autorización, comprobaciones previas, confirmación y auditoría
  no dependen del canal, así que ningún canal puede saltárselas.
- **Independiente.** Nada depende de rutas, configuración o identificadores de un entorno concreto.
- **Sin reescrituras.** Se conserva la pila actual y los servicios que ya funcionan; lo nuevo los
  envuelve.

## 2. La pieza central: el catálogo de operaciones

Una **operación** es una capacidad con nombre (`users.search`, `users.license.remove`) que se
describe a sí misma y se ejecuta siempre a través del mismo ejecutor.

```
 Ventana (WPF)   Línea de comandos   Tareas programadas   Agente de IA (MCP)
       │                │                    │                    │
       └────────────────┴─────────┬──────────┴────────────────────┘
                                  ▼
                        OperationExecutor
     validar → política del canal → conectar → autorizar → comprobar el estado real
              → plan confirmado → ejecutar → auditar
                                  │
                                  ▼
                 IOperation (catálogo: OperationRegistry)
                                  │
        ┌─────────────────────────┼──────────────────────────┐
        ▼                         ▼                          ▼
  Servicios de Graph     Servicios de Exchange Online   Otros conectores
  (SDK de .NET)          (pwsh externo)                 (Intune, Azure…)
```

### Qué declara una operación

`OperationDescriptor`: identificador, módulo, título, descripción, **impacto**, parámetros con tipo,
permisos de Microsoft que necesita, qué conexiones requiere y qué comprobaciones previas hace. Con
eso una persona, un script o un agente saben qué existe y cómo llamarlo sin leer el código.

### Impacto y compuertas

| Impacto | Ejemplos | Qué exige |
|---|---|---|
| `Read` | Buscar, consultar, informes | Conexión. No se audita |
| `Write` | Cambios reversibles: permisos de buzón, pertenencia a grupos | Autorización, comprobación previa y plan confirmado. Se audita |
| `Critical` | Destructivas o difíciles de deshacer: retirar licencias, bajas | Lo anterior, y nunca a petición de un agente |

**Plan confirmado.** Una escritura va en dos pasos. La simulación lee el estado real, devuelve el
plan y un `plan_id`; la ejecución exige ese `plan_id`. El identificador es una firma sobre la
operación, sus argumentos y una caducidad de 15 minutos, con una clave propia de cada instalación:
no vale para otros argumentos ni después de caducar. Entre la simulación y la ejecución la
comprobación previa se repite, de modo que si el estado cambió la operación se detiene.

**Política por canal** (`OperationPolicy`):

| Canal | Lectura | Escritura | Crítica |
|---|---|---|---|
| Ventana y línea de comandos | Sí | Con plan | Con plan |
| Agente (MCP) | Sí | Solo si el operador arranca el servidor con `--allow-writes`, y con plan | Nunca |
| Tarea programada | Sí | No | No |

Todo intento de escritura queda en el registro de auditoría con quién lo pidió, por qué canal, con
qué argumentos y cómo acabó, también los denegados y los detenidos.

### Comprobaciones previas como reglas del núcleo

La regla de retirada de licencias es el modelo. `LicenseRemovalGuard` es una función pura que
recibe los datos del buzón y devuelve los motivos para detenerse: más de 50 GB (o tamaño ilegible),
Litigation Hold u otra retención, archivo online. Si el estado del buzón no se puede leer, la
operación falla; no continúa a ciegas. Las reglas equivalentes de otras áreas (autorización de
protección de datos en las bajas, impacto de bloquear un dominio) siguen el mismo patrón.

### Añadir una capacidad

1. Escribir una clase que implemente `IOperation` (las de lectura heredan de `ReadOperation`).
2. Registrarla en el contenedor de dependencias como `IOperation`.

Aparece en `grex365 ops list`, en las herramientas del servidor MCP y queda disponible para la
ventana, sin tocar el núcleo. Un plugin (`IModule.RegisterServices`) puede registrar operaciones
igual que el propio producto.

## 3. Componentes

| Proyecto | Papel |
|---|---|
| `Grex365.Core` | Lógica de negocio, conectores de Graph, reglas, **operaciones**, ejecutor, servidor MCP. Sin interfaz gráfica |
| `Grex365.PowerShell` | Ejecución de PowerShell y de Exchange Online |
| `Grex365.App` | Aplicación de escritorio |
| `Grex365.Cli` | **Nuevo.** Ejecutable `grex365`: catálogo, ejecución y servidor MCP |

### Línea de comandos

```
grex365 ops list [--json]
grex365 ops describe users.license.remove
grex365 ops run users.search --arg query=ana
grex365 ops run users.license.remove --arg user=ana@contoso.com --arg sku=SPE_E3 --dry-run
grex365 ops run users.license.remove --arg user=ana@contoso.com --arg sku=SPE_E3 --plan-id <id>
```

Usa la conexión por certificado configurada en la aplicación (`%LOCALAPPDATA%\Grex365\config`) y
respeta el mismo bloqueo de tenant. `--data-dir` apunta a otra carpeta de datos.

### Agentes de IA

`grex365 mcp` es un servidor del Model Context Protocol por `stdio`. Cada operación del catálogo es
una herramienta, con su esquema de parámetros, sus permisos y sus comprobaciones en la descripción.
Cualquier agente compatible lo usa sin que GREX365 dependa de un proveedor:

```json
{ "mcpServers": { "grex365": { "command": "grex365", "args": ["mcp"] } } }
```

El agente no recibe credenciales ni tokens: pide operaciones y el servidor las ejecuta con la
conexión del operador y con las compuertas de la sección 2.

### Autenticación y permisos

Se mantiene lo que hay: certificado sin interacción (solo aplicación) o código de dispositivo
(delegado), ambos contra el registro de aplicación del propio tenant, que el asistente crea. Cada
operación declara sus permisos; con eso la aplicación puede decir qué falta antes de fallar y el
asistente puede pedir solo los permisos de los módulos que se activen (mínimo privilegio).

### Varios tenants

El modelo pasa de «una instalación, un tenant» a **perfiles**: cada perfil es una carpeta con su
configuración, su certificado, su registro de auditoría y su clave de planes. Nada se comparte
entre perfiles. La opción `--data-dir` de la línea de comandos ya funciona así; falta el selector
en la ventana.

## 4. Decisiones

| Decisión | Motivo |
|---|---|
| Conservar .NET 10 y WPF | Funciona, está probado y el público objetivo usa Windows. Una reescritura no aporta nada que se haya pedido |
| Las operaciones envuelven a los servicios existentes | Se reutilizan 21.000 líneas probadas. Los servicios pasan a ser detalle de implementación |
| Servidor MCP propio, sin dependencias | Un servidor de solo herramientas son cuatro métodos JSON-RPC. Evita una dependencia en evolución y se prueba en memoria |
| MCP por `stdio` y no un servicio en red | No abre puertos ni añade autenticación propia. El proceso lo lanza el agente del propio operador |
| Los agentes no ejecutan operaciones críticas | Un agente que se confirma a sí mismo no es una confirmación. Hasta que exista la cola de aprobaciones, lo crítico lo lanza una persona |
| Licencia MIT | Permisiva y habitual en herramientas de administración. Pendiente de confirmar por el autor |
| Abandonar la ofuscación y la distribución solo de binarios | Incompatibles con un proyecto abierto |

## 5. Hoja de ruta

Cada fase termina con algo usable y probado. El orden es el pedido por el autor; se ha adelantado
la base de operaciones porque todo lo demás se apoya en ella.

| Fase | Contenido | Estado |
|---|---|---|
| **0. Base de la plataforma** | Proyecto compilable; catálogo y ejecutor de operaciones; plan confirmado; política de canales; auditoría en el núcleo; `grex365` (línea de comandos y servidor MCP); primeras operaciones: `users.search`, `users.get`, `licenses.list`, `mailboxes.get`, `users.license.remove` | **Hecha** (ver «Pendiente de validar») |
| **1. La ventana usa el ejecutor** | Las escrituras de Usuarios, Grupos, Buzones y Reglas pasan por `OperationExecutor`; la autorización y la auditoría salen de los ViewModels; se cierra la retirada de licencias sin comprobación (P4 de la auditoría). Una sola composición de servicios para ventana y consola | Siguiente |
| **2. Usuarios, grupos y licencias** | Búsqueda global; licencia directa frente a heredada; inactivos con licencia; consumo; cruce con listados de RR. HH.; asignación; pertenencias transitivas | — |
| **3. Exchange Online** | Dominios y remitentes permitidos y bloqueados con revisión de impacto; traza de mensajes; permisos y conversión de buzones como operaciones; reglas de transporte | — |
| **4. Altas, bajas y cambios** | La baja y el alta como operaciones compuestas; reversión de una baja; autorización de protección de datos adjunta al procedimiento y verificada antes de ejecutar | — |
| **5. Automatización** | Scripts de PowerShell con parámetros declarados como operaciones; historial de ejecuciones; tareas programadas (solo lectura); encadenado | — |
| **6. Agentes** | Cola de aprobaciones para que una persona apruebe en la ventana lo que propone un agente; panel de agentes; recursos MCP con procedimientos y documentación | Base hecha en la fase 0 |
| **7. Intune** | Inventario, cumplimiento, políticas, aplicaciones y errores de instalación, sincronización | — |
| **8. Azure y resto de Microsoft 365** | Perfiles de varios tenants; Azure Resource Manager; SharePoint, Teams, Defender, Purview según demanda | — |
| **9. Publicación** | Limpieza de la historia; licencia; `grex365-2.0` como rama por defecto; guía de contribución, plantillas, política de seguridad; versiones y paquetes | Ver abajo |

### Pendiente de validar

- La lectura de buzones y la simulación de `users.license.remove` contra Exchange Online. El
  registro de aplicación de la instalación de prueba no tiene un rol de directorio admitido por
  Exchange, y la conexión se rechaza. Hace falta asignarle el rol (o usar un registro que lo
  tenga) y repetir la simulación.
- La ejecución real de `users.license.remove`, sobre una cuenta creada para la prueba.

### Antes de publicar

El repositorio ya es público. Estas tareas no se resuelven con commits nuevos:

1. Reescribir la historia para eliminar los datos de la organización (cuenta de administración,
   nombre del tenant, identificador de aplicación, direcciones de cuentas de prueba, editor de los
   paquetes) y forzar la subida. Es irreversible y cambia todos los identificadores de commit.
2. Elegir licencia.
3. Retirar del árbol el toolkit antiguo, la bitácora interna y la configuración de ofuscación.
4. Sustituir los valores de ejemplo por dominios reservados (`contoso.com`).
