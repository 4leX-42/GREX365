# Recomendación de stack y arquitectura

Para una herramienta Windows interna, moderna y escalable lo más aconsejable es una solución **nativa .NET** usando C#. Entre las opciones XAML, WinUI 3 (Windows App SDK) destaca como el framework más **moderno y con mejor rendimiento**, al tiempo que ofrece el diseño Fluent de Windows 11【3†L303-L311】【41†L43-L50】. WinUI 3 funciona sobre .NET 9/10, tiene arranque rápido y un consumo de memoria ~15–20% menor que WPF equivalente【3†L286-L294】【3†L303-L311】, y permite animaciones suaves gracias a su motor de composición nativo【3†L303-L311】. WPF sigue siendo viable (no está deprecado y .NET 9 le agregó temas Fluent y mejoras GPU【1†L127-L135】), pero exige más trabajo manual para lograr el mismo nivel visual moderno. Otras opciones: **Avalonia** permite multiplataforma (Linux/Mac), pero si la meta es puramente Windows añade complejidad y riesgo; es recomendable solo si se prevé soporte en otros OS【23†L125-L133】. MAUI es más enfocada a móvil (requiere reescribir interfaz, falta algunos controles de escritorio)【31†L339-L347】. Blazor Hybrid (UI web dentro de app) o Electron.NET exigen reescribir la UI en HTML/CSS/JS y rinden peor (cada ventana sería un navegador embebido)【32†L585-L594】【31†L410-L419】. En resumen, para este caso Windows-interno, yo optaría por **WinUI 3 con MVVM** (o WPF actualizado si hay limitación en versión de SO), y reservar Avalonia/uno sólo si en el futuro se necesita correr en macOS/Linux【23†L125-L133】【32†L520-L524】.  

* **WinUI 3 (Windows App SDK):** Native Windows UI con Fluent Design, efectos Mica/Acrylic, soporte para .NET 10 y Native AOT (resalta en rendimiento startup)【3†L274-L282】【32†L469-L478】. Controles optimizados y diseño fluido (“Windows look and feel” nativo)【32†L469-L478】. Permite temas claro/oscuro de serie. Ventajas: animaciones en compositor separado (no bloquean hilo UI)【1†L218-L226】【3†L303-L311】. Desventajas: Windows-only (requiere Win 10≥1809/Win 11), escasez de diseñador visual (aún en maduración)【32†L499-L508】【29†L49-L57】. Controladores de UI “a la Windows moderna”. Venta incremental migrando via XAML Islands si ya hay WPF【29†L53-L61】【32†L479-L487】.
* **WPF (Windows Presentation Foundation):** Framework veterano, probado en producción. Continuamente soportado (.NET 9 incluye temas Fluent, GPU mejorado【1†L127-L135】). Ecosistema amplio (refencia de controles, diseño, depuración estable). UI basado en DirectX 11 (pipeline milcore)【1†L125-L134】【3†L303-L311】. Ventajas: gran comunidad y herramientas maduras, binding dinámico (aunque clásico). Desventajas: menos eficiente en animaciones pesadas (hay que virtualizar listados para evitar stutters)【1†L199-L208】【3†L303-L311】. Si se elige WPF, recomendamos usar .NET 9+ con Enable all tweaks (dispatcher async, `BinaryFormatter` eliminado por seguridad, etc.)【1†L125-L134】.
* **Avalonia UI:** XAML multiplataforma (Win/Lin/Mac). Si se requiere ejecutar fuera de Windows, Avalonia es sólido. Emplea render con Skia, buen performance cross-platform. Para un equipo solo Windows, agrega riesgo (nueva stack, menor comunidad)【23†L125-L133】【29†L49-L57】. Avalonia XPF permite migrar pantallas WinUI/XAML más fácilmente, pero no añade nada que WinUI nativo no tenga en Windows (y WPF ya satisface necesidades actuales).
* **.NET MAUI:** Multi-dispositivo (incluye WinUI3 internamente). Fuerte en móviles/tabletas. Para una app puramente de escritorio, implica reescribir UI (XAML MAUI difiere mucho de WPF) y carece de controles de escritorio básicos (p. ej. menú, DockPanel)【31†L323-L333】. Mejor para escenarios que incluyen iOS/Android; no recomendable si solo Windows 10/11 es objetivo.
* **Blazor Hybrid (.NET MAUI Blazor):** UI en HTML/CSS/Razor dentro de WebView2. Permite migración incremental (BlazorWebView en WPF)【31†L372-L381】, pero significa abandonar XAML clásico y reescribir toda la interfaz en componentes web. Rinde peor (WebView2 embebido), sin diseñador XAML, y no aprovecha controles nativos【31†L410-L419】. Útil solo si hay fuerte necesidad de llegar al navegador, lo cual no es el caso aquí.
* **Electron.NET:** Framework web (Chromium+Node). Escalabilidad demostrada (VSCode, Slack, Discord usan Electron nativo)【32†L549-L558】, pero consume 100–200 MB por aplicación y mucha RAM. Invitaría a un cambio completo a JS/TS; rara vez vale la pena para una herramienta admin que ya tiene lógica en C#. Existe Electron.NET para mantener código C#, pero típicamente se evita en soluciones .NET enterprise (sobre todo integrar PowerShell sería complicado). Su ecosistema (NPM) es grande, pero el costo de migración sería altísimo y la integración con M365/Graph vendría por interop con Web APIs.

En síntesis, **yo usaría C# + WinUI 3 (o WPF si hubiese restricciones OS)**, con .NET 9/10. Esto da una app desktop con UI nativa moderna, empaquetada como .exe/.msix, y fácil integración con .NET. WinUI 3 provee rendimiento y UX comparables a las herramientas Microsoft (Admin Center, Windows Admin Center, etc.)【3†L303-L311】【32†L469-L478】. WPF solo si se prefiere estabilidad probada y compatibilidad con .NET Framework heredado; de lo contrario WinUI es más futuro.

## 1. Framework y arquitectura recomendada 

- **C# / .NET 9+**: Lenguaje base de la aplicación. Aprovecha últimos avances de .NET (C# 13/14, AOT nativo, mejora ARM64, GC, etc.)【1†L166-L175】【3†L314-L323】. Con el runtime .NET 10/11 previsto, la aplicación podrá beneficiarse de mejoras en rendimiento y despliegue (AOT, single file, mejoras de inicio)【1†L166-L175】【3†L314-L323】.
- **Patrón MVVM**: Separación estricta Vista/Vista-Modelo/Modelo. Facilita testing y mantenimiento. MVVMToolkit (CommunityToolkit.Mvvm) o frameworks como Prism/Caliburn.Micro pueden usarse para inyectar dependencias, manejar navegación y plugins. El MVVM es estándar en WPF/WinUI (data binding, INotifyPropertyChanged, Commands).
- **Arquitectura modular**: Usar Prism o MEF para **módulos/plugins**. Cada funcionalidad (por ejemplo, importación CSV, auditoría, etc.) se implementa como módulo independiente cargable dinámicamente. Esto permite extender la app sin recompilar el núcleo. Por ejemplo, Prism suporta inyección de dependencias y regiónes en la UI para cargar módulos por demanda.
- **Frontend/Backend desacoplados**: La capa UI (WinUI) debe comunicarse con el “motor” de automatización por interfaces claras. Podría implementarse como **servicio interno** o librería dedicada (“core engine”) que ejecuta los scripts PowerShell. Esta capa de motor encapsula runspaces y abstracciones de comandos. De esta forma la lógica de negocios (powershell+Graph) queda independiente de los eventos UI.
- **Motor de automatización**: Una clase o servicio (singleton) que administra la ejecución de PowerShell. Usa `System.Management.Automation.PowerShell` con RunspacePool para crear runspaces asincrónicos. Un **RunspacePool** permite ejecutar múltiples scripts concurrentemente, especificando un máximo de concurrencia para evitar sobrecargar CPU/memoria【13†L141-L149】.  Cada script se envía mediante `PowerShell.BeginInvoke()` y se esperan resultados con `EndInvoke()`.
- **Registro centralizado (logging)**: Utilizar un framework como Serilog o NLog. Registrar eventos de acción (ejecución de scripts, errores) en ficheros rotativos y/o destino remoto (DB o SIEM). Configurar distintos niveles (INFO, ERROR) y formatos estructurados (JSON) para integración con monitorización. El motor PowerShell debe capturar streams Verbose/Warning/Error de cada ejecución y volcarlos al log central del app.
- **Manejo de errores**: Capturar excepciones en todas las capas. Usar un manejador global (DispatcherUnhandledException) para errores de UI, y en la capa backend envolver cada invocación PowerShell en try/catch logueando el error completo. Proveer retroalimentación al usuario (diálogo o panel de errores) en lugar de caer. Implementar reintentos configurables para tareas críticas con fallos transitorios.
- **Configuración dinámica**: Configurar la app via JSON o XML externo (p. ej. AppSettings.json). Permitir editar parámetros de conexión Graph, credenciales cifradas, rutas, etc. Tal vez exponer UI de administración para cambiar ajustes sin redeploy. Cargar la configuración en tiempo de ejecución (p.ej. usando `IOptions<T>` o similar).
- **Control de permisos/roles**: Integrarse con Windows AD o un servicio de identidad (Azure AD) para controlar quién puede ejecutar qué. La app puede requerir inicio de sesión del usuario Windows, y luego validar grupos/roles en AD. Por ejemplo, solo ciertos grupos corporativos podrían ver funciones de auditoría. Alternativamente, usar un mecanismo de autorización propio (roles definidos en configuración) con políticas en la capa de servicios.
- **Auditoría**: Registrar acciones del usuario (quien hizo qué y cuándo), guardando eventos relevantes en logs o base de datos. Cada vez que se ejecuta un workflow administrativo (creación de usuarios, asignación de licencias, etc.), grabar en un log de auditoría separado (p.ej. EventLog o tabla SQL) para cumplimiento. Esto va más allá del logging técnico: es seguimiento de negocio.
- **Telemetría**: Integrar una plataforma de monitoreo (Application Insights, OpenTelemetry) para enviar métricas de uso y performance. Por ejemplo, cuántos scripts se ejecutan por día, tiempos de respuesta, errores frecuentes. Usar Application Insights para recoger excepciones no capturadas, métricas de latencia de tareas y trazas de eventos. Esto ayuda a detectar cuellos de botella reales en entorno corporativo.

## 2. Integración con PowerShell 7

Ya que la lógica de automatización está en PowerShell 7, debemos incrustar su ejecución de forma asíncrona y sin ventanas de consola. Las mejores prácticas son:

- **Runspaces asincrónicos**: Usar `System.Management.Automation.PowerShell` con runspaces en segundo plano. Crear un **RunspacePool** compartido y usar `PowerShell.BeginInvoke()` en lugar de `Invoke()`, de modo que cada script corra en paralelo sin bloquear el hilo de UI【36†L51-L60】. Por ejemplo:  
  ```csharp
  using RunspacePool pool = RunspaceFactory.CreateRunspacePool(minRunspaces:1, maxRunspaces:5);
  pool.Open();
  PowerShell ps = PowerShell.Create();
  ps.RunspacePool = pool;
  ps.AddScript("...script.ps1...");
  IAsyncResult result = ps.BeginInvoke();
  // Registrar eventos o mostrar "en progreso"
  ps.EndInvoke(result);
  ```
  Esto evita que la interfaz se congele mientras el script se ejecuta【36†L51-L60】. 

- **Threads y tareas**: Dado que WinUI/WPF usa un único hilo UI, conviene lanzar las ejecuciones en **Tasks** (TPL) o **BackgroundWorker**, actualizando el UI mediante el dispatcher al completarse. Cada PowerShell async puede correrse en un `Task.Run` y luego notificar progreso con `DispatcherQueue.TryEnqueue` (WinUI) o `Dispatcher.Invoke` (WPF). Nunca interactuar directamente con controles desde el runspace【36†L58-L66】.

- **Gestión de concurrencia**: Controlar el número de scripts simultáneos (p.ej. un máximo de 5-10 tareas a la vez) usando un `SemaphoreSlim` o el límite de runspace pool. Esto evita sobrecargar la CPU o agotar la memoria (cada script de PS load de módulos, GC, etc.). Se puede implementar una cola simple: cada petición de tarea entra en una `BlockingCollection` y se procesan en background hasta cierto límite.

- **Comunicación interna (IPC)**: Si se opta por desacoplar el motor de la UI en procesos separados, usar canales IPC formales. Por ejemplo, un **servicio Windows** o proceso hilo (*backend*) que expone un API local (WCF, gRPC o Named Pipes) para recibir peticiones de ejecución. La UI llamaría a ese servicio para lanzar scripts y recibir resultados/estado. Esto permite aislar fallos; si un script crashea el proceso motor, la UI puede reiniciarlo sin caer. Sin embargo, para simplificar, no es estrictamente necesario: podemos mantener todo dentro de la misma app en runspaces.

- **Sessions de Exchange Online/Graph**: Las conexiones a M365 (MS Graph, EXO) deben gestionarse centralizadamente. Por ejemplo, crear un módulo de conexión que maneje tokens y refrescos. Si varios scripts usan la misma sesión, reusar credenciales para evitar múltiples logins. Se puede mantener un objeto `GraphServiceClient` persistente o usar el módulo `MSAL.PS` en el runspace con credenciales secure-string pasadas. Es clave cerrar sesion (`Disconnect-ExchangeOnline`, `Stop-MgGraphSession`) al terminar o al salir del programa.

- **Actualización de estado en tiempo real**: Implementar eventos o callbacks en el código C# que se disparen desde los runspaces. La clase `PowerShell` tiene la propiedad `Streams.Progress` que envía actualizaciones de progreso (ej. progreso de comandos). Además, usar `IObservable` o eventos personalizados: cada tarea puede reportar su estado (“en cola”, “ejecutando”, “exitoso”, “error”) que la UI muestra en tiempo real (p.ej. en un panel de logs o indicadores de carga).

- **Logging de ejecución**: Redirigir los flujos de salida de PowerShell (`Output`, `Error`, `Verbose`) hacia el sistema de logs. En C#, suscribirse a `ps.Streams.Error.DataAdded` para capturar errores, y registrar su mensaje. Los Verbose/Debug se pueden capturar igual y mostrar si el usuario habilita modo detallado. Es importante congelar el log a disco o base de datos incluso de mensajes stdout (por ejemplo, resultados de exportaciones).

- **Jobs vs Runspaces**: Preferir runspaces por su eficiencia. Un **PowerShell Job** crea un proceso aparte y serializa salida, lo que implica 2–5 segundos de overhead por job【13†L109-L117】【13†L141-L149】. En el blog técnico citado, un Start-Job tardó ~2.5 s en arrancar y ~5 s en terminar, mientras que un runspace (PowerShell.BeginInvoke) tardó ~2.5 s en arrancar pero prácticamente **0 ms** en concluir【13†L109-L117】【13†L141-L149】. Es decir, runspaces no crean nuevos procesos y comparten contexto, por lo que son mucho más rápidos y escalan mejor.

- **Evitar consola visible**: Al ejecutar PowerShell desde C#, no lanzar “pwsh.exe” con CreateNoWindow (aunque es una opción). Es mejor usar la API embebida (`PowerShell.Create()`), así no aparecen ventanas de consola. Para casos puntuales donde se desee un proceso aparte, usar `ProcessStartInfo` con `RedirectStandardOutput = true` y `CreateNoWindow = true`.

- **Seguridad**: Como se ejecutan scripts potentes (modifican usuarios, licencias, etc.), validar cuidadosamente cualquier entrada y usar siempre credenciales seguras (no hardcode, usar Windows Credential Manager o Azure Key Vault para secrets). Además, ejecutar los runspaces con las credenciales necesarias; por defecto corren con el usuario actual, lo cual es deseable si es admin IT con permisos M365.

## 3. Arquitectura profesional

La solución resultante debe seguir **principios de diseño empresarial**:  

- **MVVM limpio**: Cada módulo (por ejemplo, administración de grupos, auditoría, offboarding, etc.) tendría su Vista XAML, ViewModel y Model separados. Los ViewModels no deben contener lógica UI; se comunica con el motor PowerShell vía interfaces (por ejemplo, un servicio `IPowerShellRunner`). Así la lógica se testea fuera de UI. Utilizar librerías MVVM (CommunityToolkit, Prism) para facilitar binding y navegación.
- **Separación de capas**: La interfaz (WinUI) sólo se preocupa por presentar datos y comandos. Toda la lógica de ejecución de scripts, manejo de datos e interacción con M365 se aísla en la capa de *servicios* o motor. Por ejemplo, una clase `IdentityEngine` expone métodos C# (async) que internamente ejecutan los scripts PowerShell y retornan un resultado tipado (o lanza excepción). El frontend invoca estos servicios sin usar `Process` directamente.
- **Modularidad/Plugins**: Diseñar un sistema donde nuevas funcionalidades se puedan añadir como ensamblados plug-and-play. Esto permite que, por ejemplo, haya un plugin de “Informes de identidad” que se carga dinámicamente. Se puede usar la extensión *PluginContract/Dynamic Plugin Loader* o incluso MEF para descubrir y cargar módulos por convención. Cada plugin registra sus comandos en la interfaz, y no provoca recompilación del core.
- **Logging centralizado**: Implementar un servicio de log (p.ej. Serilog) que escriba en ficheros y, opcionalmente, envíe logs críticos a un servidor (Syslog, Elasticsearch, Azure Log Analytics). Incluir ID de sesión/usuario en cada entrada. Las acciones del usuario (botones clickeados, opciones elegidas) también deben loguearse como eventos de auditoría.
- **Gestión de errores enterprise**: Adoptar un patrón de manejo de errores uniforme. Por ejemplo, toda llamada al backend se envuelve en un try/catch general que captura excepciones no anticipadas, las registra y muestra un error genérico al usuario. Los errores esperados (como un script devolviendo fallo por permisos insuficientes) deben manejarse para mostrar mensaje legible. Se pueden usar librerías de resiliencia como Polly para reintentos automáticos en fallos transitorios (p. ej. conexión Graph perdida).
- **Permisos/Roles**: Integrar con Active Directory/Azure AD. Opciones: validar que el usuario de Windows pertenece a ciertos grupos antes de habilitar acciones en la UI; o implementar un login interno (p.ej. OAuth a Azure AD) para determinar roles. Esto permite funciones granulares (por ejemplo, “solo usuarios del grupo `IT-Admins` pueden usar la función X”).
- **Auditoría**: Además del logging técnico, registrar *audit trail* de cambios importantes. Ejemplo: «El usuario admin1 deshabilitó el buzón userX@contoso». Esto puede ir a un log específico o base de datos. Compliance corporativo a menudo exige saber quién hizo qué cambio administrativo.
- **Telemetría y monitoreo**: Incorporar telemetría de uso (Application Insights u otro). Enviar métricas clave (tiempos de ejecución de scripts, número de operaciones/día) y trazas de eventos. También monitorear el estado de la aplicación (crashes, memoria usada). Este telemetría puede ayudar a anticipar problemas en rollout empresarial.
- **CI/CD y control de versiones**: Usar un repositorio tipo Git y pipeline (Azure Pipelines o GitHub Actions) para construir la app. En cada commit se compila y se genera el paquete .msix/.exe firmado. Se puede automatizar pruebas unitarias y de integración (p.ej. tests de sanity para scripts críticos). Tener versionamiento semántico para facilitar despliegues controlados en ambiente corporate.

## 4. Interfaz de usuario (UI/UX)

La UI debe ser **polished y profesional**. Algunas pautas:

- **Fluent Design / FluentWPF**: Aprovechar los estilos Fluent. WinUI 3 lo integra nativamente, con soporte para materia Mica, Acrylic, esquinas redondeadas, tema oscuro/claro sincronizado con el sistema【32†L469-L478】【41†L43-L50】. Si fuera WPF, usar el tema Fluent oficial o bibliotecas de terceros (MahApps.Metro, FluentWPF).
- **Dashboard principal**: Pantalla inicial con métricas clave (cards con número de usuarios activos, grupos, salud del tenant, etc.), similar a Admin Center. Puede incluir gráficos sencillos (usar librerías de charts de .NET o Microsoft Toolkit).
- **Navegación lateral**: Un menú tipo sidebar (p.ej. `NavigationView` de WinUI) con secciones (“Identidad”, “Licencias”, “Auditoría”, “Configuración”). Esto proporciona experiencia de app moderna (Windows Admin Center, VS Code, Intune UX).
- **Animaciones suaves**: Transiciones de pantallas, realce de botones, loading spinners. WinUI gestiona animations en compositor (no bloquean). En WPF, usar `Storyboard` o la biblioteca WinUI Toolkit TransitionCollection para animar cambios de página con fade/slide.
- **Modo claro/oscuro**: Tema dinámico según preferencia de Windows o manual por usuario. WinUI permite usar `ThemeMode="System"` para sincronizar con sistema, o cambiar en runtime. Cada control debe respetar el tema (colores acordes).
- **Tablas y listados avanzados**: Para ver datos (miembros de grupo, logs, etc.), usar el `DataGrid` con capacidad de ordenar/filtrar/paginación virtualizada. Controladores comerciales (Telerik, DevExpress) ofrecen grids ricas con filtrado en UI, pero el `DataGrid` de WinUI/WPF con `VirtualizingStackPanel` suele bastar para miles de filas sin problema【3†L303-L311】.
- **Carga asíncrona y progresos**: Siempre que se carguen datos de la red o scripts, mostrar indicadores (spinner o barra de progreso). Por ejemplo, usar un `ProgressBar` o `ProgressRing` de WinUI enlazado a un valor IProgress desde el backend, para no bloquear la UI.
- **Notificaciones y consola integrada**: Incluir una sección tipo *log* o ventana de eventos donde se vean mensajes en tiempo real (por ej. conexión iniciada, Script finalizado). Además, se puede embebir un terminal PowerShell interactivo opcional. Existen controles como **EasyWindowsTerminalControl** (WPF/WinUI) que usan el backend de Windows Terminal para mostrar un consolas embebido【48†L299-L307】. Con esto, un administrador podría ver/ejecutar comandos directamente desde la app. 
- **Animaciones y gráficos**: Para “sensación enterprise”, animar transiciones discretamente (p. ej. ItemAppear de lista) y usar gráficas de librerías (Syncfusion, SciChart) para tendencias. Pero no abusar de efectos que distraigan; mantener UX sobrio.
- **Experiencia responsiva**: Aunque es app desktop fija, permitir resizing de ventanas y que los layouts se adapten (grids fluidos, etc.). Optimizar DPI (“per-monitor DPI awareness” incorporado en .NET 9 mejora WPF). Soporte multi-ventana (WinUI puede abrir nuevas ventanas, útil para herramientas de monitoreo en paralelo).
- **Comparación visual**: UI parecida a *Microsoft Admin Center* e *Intune*, con acentos azules, tarjetas informativas, paneles expandibles. Por ejemplo, Admin Center usa `Pivot` o `NavigationView` y tarjetas de estadísticas, así que usar esos controles nativos para que se sienta familiar.

## 5. Packaging y despliegue

- **Generación de `.exe` / `.msix`**: Con .NET 9/10, la aplicación puede publicarse en un solo ejecutable. Usando `<PublishSingleFile>true</PublishSingleFile>` en el proyecto (auto-sell-contained) se crea un EXE único con runtime incluido【46†L75-L83】. Alternativamente, lo empacaremos como **MSIX** para facilitar despliegue corporativo. Visual Studio soporta “Single-project MSIX” que toma el exe y genera paquete instalable. MSIX garantiza integridad firmada y fácil actualización.
- **Firma digital**: Firmar binarios (.exe/.msix) con certificado de código (preferiblemente EV). Esto evita advertencias de seguridad y cumple políticas corporativas. MSIX exige firma.
- **Instaladores**: Aunque un .exe self-contained basta, MSIX permite actualizaciones automáticas integradas. Considerar usar la herramienta *App Installer* (`.appinstaller`) que puede comprobar actualizaciones de un feed y aplicar nuevos MSIX automáticamente. Para actualizaciones dentro de la propia app, .NET también soporta clickonce, pero MSIX es más moderno.
- **Auto-update**: Configurar AppInstaller con `UpdateBlocksActivation` o elegir que busque actualizaciones al iniciar. Esto permite que IT distribuya nuevas versiones sin reinstalar manualmente. También se puede integrar dentro de la app un mecanismo que invoque `DesktopBridge.AppInstaller` APIs para forzar check de actualización a voluntad.
- **Despliegue corporativo**: Documentar que el .msix se puede distribuir por Intune/ConfigMgr (SCCM)【44†L42-L50】. Como MSIX aparece en el catálogo de aplicaciones internas de Intune, se puede propagar automáticamente. También es posible instalar vía `Add-AppxPackage` en PowerShell remotamente. Además, AppLocker/GPO pueden restringir apps no autorizadas (MSIX permite definir reglas por Publisher)【44†L131-L139】.
- **Compatibilidad Windows 10/11**: WinUI 3 requiere Win 10 1809+; si la organización tiene instalaciones anteriores, habría que evaluar WPF o proveer builds separados. Windows 11 nativo soporta WinUI completo. .NET 9/10 es compatible con Win 10+.
- **Requisitos `.NET`**: Decidir entre framework-dependent (requiere runtime .NET preinstalado) o self-contained. Para facilidad, self-contained libera al usuario de instalar .NET aparte, pero aumenta tamaño (~100 MB). Dado que es herramienta interna, se podría optar por dependencia (asumiendo .NET 9+ ya está en estaciones) para reducir tamaño.
- **Consideraciones de seguridad**: MSIX se integra con AppLocker/GPO (por ejemplo, restringir apps no firmadas)【44†L131-L139】. Indicar a TI que puede bloquear ejecuciones no deseadas. Además, usar TLS actual (no dejar de lado forcing updated TLS/crypto) al comunicarse con servicios online.
- **CI/CD y canal de distribución**: Implementar un pipeline que, al hacer build en Git, genere el msix firmado y lo suba a un repositorio interno (Azure Artifacts, InTune repo, etc.). Mantener control de versiones y notas de release. Los administradores IT pueden otorgar el paquete vía herramientas de despliegue corporativo o un portal interno de software.

## 6. Rendimiento y estabilidad

- **Escalabilidad y bloqueos**: La arquitectura actual (scripts en consola + XAMPP) falla bajo carga. Los procesos de consola bloquean secuencialmente y no gestionan bien la concurrencia. Además, un servidor web local (XAMPP) es innecesario para una app desktop y añade complejidad (Punto único de fallo, dependencias adicionales). Para evitar congelamientos, todo trabajo pesado *no debe correr en el hilo UI*. Se ha de usar asincronía: runspaces en background, `await Task`, etc., de manera que la UI permanezca responsiva siempre.
- **Uso de recursos**: Planificar la memoria y CPU. Cada instancia de PowerShell carga librerías (Graph, módulos M365). Por ello, limitar concurrencia (ver arriba) y liberar runspaces tras usarlos (`Dispose()`). Revisar fugas: por ejemplo, las conexiones de .NET Graph deben ser .Dispose() adecuadamente. También vigilar leaks de COM/WinAPI (si se usa XAML Islands, pero aquí no es el caso).
- **Ejecución de largas tareas**: Para tareas muy largas (horas), implementar un patrón de “jobs” internos con estado. P.ej. una operación de auditoría masiva podría ejecutar múltiples sub-scripts y reportar progreso intermedio, para no causar timeout. La UI puede mostrar el progreso y dar opción de cancelar la tarea (cancelar el runspace usando `CancellationToken`).
- **Evitar congelamientos**: Como best practice, usar `await` en tareas async. Si se necesita bloquear por breve momento, se puede usar `DispatcherQueue.TryEnqueue(DispatcherQueuePriority.Lowest, ...)` para dar tiempo al hilo UI a procesar input entre chucks. Sin embargo, ideal es que el flujo sea totalmente asíncrono. WinUI/WPF permiten usar `async void` en comandos si es correcto.
- **Base de datos/servicios remotos**: Si la app consulta servidores (Graph, AD LDS, etc.), implementar timeouts razonables y reintentos. Por ejemplo, para peticiones Graph usar `HttpClient` con políticas de retry. Evitar deadlocks al capturar tareas sin `await` adecuadamente.
- **Monitorización en runtime**: Incluir en la app una sección de “Estado” que muestre tiempo de actividad, uso de CPU/memoria. Esto ayuda a detectar si la app va creciendo en recursos anormalmente (por ejemplo, por memory leak). También se pueden exponer contadores de rendimiento .NET vía Performance Counters personalizados.

En conclusión, hay que **romper el modelo actual de “script-consola interconectados”** y migrar a un diseño no bloqueante, multicapa y con gestión profesional de tareas. Un frontend moderno + backend robusto manejará cientos de operaciones simultáneas mejor que una serie de scripts secuenciales en consola.

## 7. Roadmap técnico (fases estimadas)

1. **Fase 1 – Refactor del backend y preparación de la plataforma**  
   - **Objetivo:** Organizar la lógica existente de scripts en una librería o servicio `.NET`. Crear la capa de abstracción PowerShell (clases de ejecución, logging). Configurar proyecto .NET base, inyección de dependencias, MVVM skeleton.  
   - **Tareas:** Diseñar interfaces para la ejecución de scripts; inicializar RunspacePool; primera versión de registro/log. Migrar lógica de “Main.ps1”, modularizar funciones en métodos C#. Preparar proyecto de WinUI/Vista vacía.  
   - **Dificultad:** Media. Se requiere entender todos los scripts actuales para no romper la lógica.  
   - **Tiempo estimado:** ~4–6 semanas.  
   - **Dependencias:** Ninguna externa; requiere colaboración con equipo actual de scripting.  
   - **Riesgos:** Fallos en entender scripts complejos; posponer UI nuevo; posibles incompatibilidades si los scripts usan características especiales.  
   - **Prioridad:** Alta – sin este refactor no puede empezar la UI moderna.

2. **Fase 2 – Motor PowerShell y ejecución asincrónica**  
   - **Objetivo:** Implementar la ejecución de cada función/flujo de trabajo en background sin bloqueos. Manejo de estados y logs.  
   - **Tareas:**  
     - Construir un servicio `PowerShellRunner` que ejecute scripts desde C# de forma asincrónica (Runspaces, captura de streams).  
     - Integrar colas/tareas concurrentes (por ejemplo, un `SemaphoreSlim`).  
     - Registrar salida completa de scripts en sistema de logs y exponer eventos/progress al ViewModel.  
     - Gestión de sesiones Graph/EXO: inicializar conexión, refrescar tokens, reusar sesión entre invocaciones.  
     - Tests unitarios de funcionalidad (por ejemplo, simular un script sencillo).  
   - **Dificultad:** Alta (multithreading + abstracciones).  
   - **Tiempo estimado:** ~4–6 semanas.  
   - **Dependencias:** Completación de Fase1.  
   - **Riesgos:** Errores de concurrencia (deadlocks), problemas al mezclar contextos de ejecución. Probables ajustes iterativos.  
   - **Prioridad:** Alta – necesaria para UI siga trabajando.  
     
3. **Fase 3 – Nueva UI moderna (WinUI3/WPF)**  
   - **Objetivo:** Desarrollar la interfaz gráfica profesional.  
   - **Tareas:**  
     - Crear vistas XAML para cada módulo (dashboard, formularios, panel logs, etc.) siguiendo diseño Fluent.  
     - Implementar navegación (por ejemplo con `NavigationView`) y enlaces (bindings) a ViewModels.  
     - Añadir controles avanzados: tablas (DataGrid), gráficos (chart control), etc.  
     - Incorporar UI de terminal integrada opcional (usando EasyWindowsTerminalControl u otro).  
     - Tema claro/oscuro y estilos consistentes.  
     - Integrar notificaciones (usando `MessageDialog` o `ToastNotification` WinUI).  
   - **Dificultad:** Alta (UI/UX + binding).  
   - **Tiempo estimado:** ~6–8 semanas.  
   - **Dependencias:** Fase2 estable; feedback de UX.  
   - **Riesgos:** Cambios de requisitos visuales; necesitar iterar con usuario final para ajustes. Posibles ajustes a lógica backend para datos enlazados.  
   - **Prioridad:** Alta. Visual y usabilidad importantes.

4. **Fase 4 – Arquitectura modular y plugins**  
   - **Objetivo:** Implementar sistema de plugins/módulos para añadir funcionalidades futuras.  
   - **Tareas:**  
     - Integrar Prism (o similar) para carga dinámica de módulos.  
     - Convertir algunas funcionalidades (p.ej. “Offboarding Wizard”) en módulos independientes cargables.  
     - Crear interfaz de descubrimiento (p. ej. carpeta de plugins desde donde se cargan .dll).  
     - Asegurar que cada plugin puede añadirse sin modificar el core.  
   - **Dificultad:** Media/alta (dependencia de framework, testing).  
   - **Tiempo estimado:** ~3–4 semanas.  
   - **Dependencias:** Fases 1–3 completas.  
   - **Riesgos:** Integración de Prism puede romper referencias. Buen aislamiento de módulos es clave.  
   - **Prioridad:** Media. Importante para extensibilidad a largo plazo.

5. **Fase 5 – Packaging y despliegue**  
   - **Objetivo:** Preparar instalador y plan de distribución empresarial.  
   - **Tareas:**  
     - Configurar publicación MSIX (.msixbundle) con auto-update.  
     - Implementar firma de binarios en pipeline de CI.  
     - Crear documento guía de instalación por Intune/SCCM.  
     - Pruebas de instalación en entornos Windows 10/11.  
   - **Dificultad:** Media. Con MSIX es relativamente directo, pero puede haber detalles.  
   - **Tiempo estimado:** ~2–3 semanas.  
   - **Dependencias:** Aplicación funcional en .exe. Herramientas de firma disponibles.  
   - **Riesgos:** Problemas de compatibilidad de manifiesto MSIX, errores de permisos AppLocker.  
   - **Prioridad:** Alta (sin despliegue no hay entrega).

6. **Fase 6 – Telemetría y features enterprise**  
   - **Objetivo:** Integrar monitoreo, auditoría y ajustes finales.  
   - **Tareas:**  
     - Incluir Application Insights (o similar) para excepciones y métricas de performance.  
     - Ajustar niveles de logging (DEBUG, INFO) en producción.  
     - Crear documentación técnica interna (arquitectura, manual de operación).  
     - QA final con escenarios reales (simular 100 conexiones simultáneas, etc.).  
   - **Dificultad:** Media. Configuración de telemetría y análisis.  
   - **Tiempo estimado:** ~2–3 semanas.  
   - **Dependencias:** Todo lo anterior implementado y estable.  
   - **Riesgos:** Revelación de datos sensibles en telemetría si no se filtra.  
   - **Prioridad:** Media/alta. Mejora calidad a largo plazo.

Cada fase incluiría revisiones y posibles retroalimentaciones de stakeholders. Las estimaciones suponen un equipo pequeño dedicado; ajustarlas según recursos reales.

## 8. Recomendación final

**Stack seleccionado:** Basándome en los requisitos enterprise y Windows-only, uso **C# + WinUI 3 (.NET 9/10)** como principal. WinUI 3 me garantiza UI moderna (Fluent) y rendimiento óptimo para Windows 10/11【3†L303-L311】【32†L469-L478】. Como UI components nativos, las animaciones y temas serán consistentes con el ecosistema Microsoft. Backend en .NET usando `System.Management.Automation` para PowerShell 7, con un servicio centralizado de ejecución (runspaces) que interactúe con la UI via MVVM.

**Por qué no…**:
- **WPF puro**: Aunque probado, exigiría mucho diseño manual para igualar estética Fluent actual. Es menos "futurible" (aunque se mantendrá, no genera innovación gráfica tan fácil).  
- **Avalonia/Uno**: A menos que se prevea port a Mac/Linux, no justifican su complejidad y aprendizaje. En caso de necesidad cross-platform, valdrían (Avalonia incluso publicó este año análisis honesto), pero aquí la prioridad es Windows.  
- **Electron/Blazor**: Requieren reescribir UI en web, rompen MVVM nativo. Conducen a apps muy pesadas (Electron) o a UX web dentro de escritorio (Blazor). Pierden la integración natural con tecnología de Windows (los controles no son nativos, aunque funcionan, la sensación no es la misma). Además la lógica existente en PowerShell/Graph se vuelve más complicada de invocar desde JS.

**Errores actuales**: La arquitectura actual basada en scripts de consola y XAMPP está muy acoplada y bloqueante. No hay verdaderos estados; cada script lanza su propio `pwsh`. El UI web con Bootstrap es un *parche* improvisado, vulnerable a bloqueos de proceso y mal manejo de errores. La falta de separación (todo ocurre en un solo servidor local) complica depuración. La experiencia usuario es pobre (no hay cargas intermedias, todo aparece de golpe). Para convertirlo en producto serio, debemos **romper con esa estructura** y reconstruir capas limpias.

**Camino profesional**: Empezar por encapsular la lógica en .NET (como hemos descrito) y luego desarrollar una app de escritorio robusta. Mantener la lógica de negocio de los scripts (ya validada) pero ejecutar desde C#. En concreto, Yo propondría:
- Construir una librería .NET que exponga todas las acciones (p.ej. `AddGroupMembers`, `ConvertSharedToUser`) como métodos asíncronos. Esta librería usaría internamente `PowerShell.Create()` o runspaces. 
- UI en WinUI consume esa librería. 
- Usar Prism para módulos “flujo de trabajo” (por ejemplo módulo Offboarding, Auditoría). 
- Centralizar configuración y credenciales. 
- Desplegar vía MSIX auto-update en Intune. 

Esto cumple los principios clave: estabilidad (no hay consolas expuestas, tasks en background), mantenibilidad (MVVM, código organizado, logs estructurados), excelente UX (Fluent y moderno) y capacidad de escalado (módulos, telemetría, servicio central). 

Finalmente, si fuera necesario un enfoque híbrido, podría separar aún más: por ejemplo, un **servicio Windows** (en .NET Core) como motor PowerShell, y la app WinUI solo de frontend conectado por gRPC. Esto protegería contra crashes (el servicio se podría reiniciar independientemente). Pero tal separación solo vale la pena si esperamos escalar muchísimo o distribuir la carga. En la mayoría de escenarios, un app combinada en .NET es suficiente.

En resumen: **stack elegido: C# + WinUI 3 (.NET 9/10) con MVVM + PowerShell 7 integrado vía runspaces**. Esta combinación ofrece un producto Windows de nivel enterprise, con interfaz premium y robustez empresarial【3†L303-L311】【32†L520-L524】. Evitaríamos arquitecturas JS/Electron complejas o hacks web locales. Este camino, aunque exige inversión inicial (reposicionar el proyecto en .NET), garantizará un software escalable y mantenible para el futuro. 

**Fuentes:** Comparativas .NET Desktop (CTCO, Avalonia, Uno) y documentación oficial confirman que WinUI 3 es óptimo para aplicaciones empresariales Windows【3†L303-L311】【32†L520-L524】. La experiencia de integrar PowerShell en GUIs (blogs especializados) muestra que runspaces asincrónicas con MVVM es la práctica recomendada【36†L51-L60】【13†L109-L117】. La gestión de paquetes MSIX e intune en entornos corporativos está documentada en Microsoft Learn【44†L42-L50】【46†L75-L83】, garantizando una distribución profesional. Todos estos elementos apuntan hacia una arquitectura moderna y sólida conforme a estándares IT empresariales.

