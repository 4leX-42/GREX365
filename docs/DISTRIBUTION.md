# GREX365 — Estrategia de distribución y despliegue

> Deep-research 2026-06-05. Objetivo: distribuir SOLO binarios (sin exponer código fuente).
> Contexto: .NET 10 WPF self-contained, audiencia IT admins/MSPs, desarrollador individual/pequeña empresa (España).

## Hallazgos clave (cambian el plan)

1. **Microsoft Store ya no cuesta nada**: cuota eliminada para cuentas individuales (sep 2025) y company (may 2026). Sigue habiendo KYC (individual: DNI + selfie; company: verificación de negocio D-U-N-S/docs), pero **gratis**. La premisa de descartarla por KYC merece revisión: es la única vía con **cero avisos SmartScreen**.
2. **Azure Trusted Signing (ahora "Azure Artifact Signing") NO disponible para individuos en España** — solo individuos USA/Canadá; en la UE solo **organizaciones**. ~10 $/mes (Basic). Decisivo: firma barata requiere entidad legal.
3. **Los certificados EV ya NO dan reputación SmartScreen instantánea** (confirmado por Microsoft). No pagar premium EV solo por SmartScreen.
4. **NativeAOT y trimming inviables con WPF** (el SDK deshabilita trimming para WPF; AOT exige trimming). Mitigación: **ReadyToRun** (arranque, no tamaño ni protección).
5. **GitHub Releases en repo privado NO permite descarga pública de assets** (requiere token). Para binarios públicos sin fuente: repo público solo-releases, o web propia.

## Canales — comparativa

| Canal | Requisitos | Coste | KYC | Hosting propio | Mantenimiento | Notas |
|---|---|---|---|---|---|---|
| MS Store (individual) | Verificación ID + selfie | Gratis | Sí (ligero) | No | Medio | Cero SmartScreen; MS re-firma |
| MS Store (company) | Verificación negocio | Gratis | Sí (negocio) | No | Medio | |
| **WinGet** (winget-pkgs) | PR YAML + instalador en URL pública | Gratis | No | **Sí** | Bajo (`wingetcreate` en CI) | Estándar de facto IT admins |
| Chocolatey community | Moderación humana por versión | Gratis | No | Normalmente | Alto | Latencia; descartar |
| GitHub Releases (repo público solo-releases) | — | Gratis | No | No | Bajo | Sin fuente en el árbol |
| Web propia | Dominio + CDN | Bajo | No | Sí | Bajo | Control total |
| Scoop | Manifiesto JSON | Gratis | No | Recomendable | Bajo | Público dev, no corporativo |

## Empaquetado

- **Velopack** — recomendado: instalador + auto-update (full + delta), MIT, activo (v1.2.0 jun-2026), sucesor de Squirrel/Clowd.Squirrel, soporta WPF. CLI `vpk`.
- **MSIX** — firma **obligatoria** (self-signed solo dev). Bueno para Store / `.appinstaller`. (Scaffold ya existe — ver PACKAGING.md.)
- **WiX v6 (MSI)** — para despliegue empresarial Intune/GPO en clientes grandes.
- **Inno Setup** — alternativa simple a Velopack sin auto-update.
- ClickOnce / Squirrel.Windows: evitar (legacy).
- `dotnet publish`: self-contained ✓, single-file viable (probar), **trimming NO**, **NativeAOT NO**, **ReadyToRun SÍ**.

## Code signing

| Opción | Coste | Elegibilidad (individual ES) | Notas |
|---|---|---|---|
| Azure Artifact Signing | ~10 $/mes + sub Azure de pago | **NO individual ES** — UE solo org | Sin token, integra CI/CD |
| OV tradicional (Sectigo/SSL.com/Certum) | ~65–300 $/año | Sí | Token HSM obligatorio desde jun-2023 (o cloud-HSM) |
| EV | ~250–560 $/año | Sí | Ya no da reputación instantánea — no justifica premium |
| SignPath.io | Gratis solo OSS | Comercial de pago | No aplica (closed source) |
| Sin firmar | 0 | — | SmartScreen bloquea, Smart App Control (Win11) bloquea, AV false positives. Inviable en corporativo |

Reputación SmartScreen: se acumula por **identidad de firma** estable entre versiones; semanas + cientos de instalaciones limpias por release nueva.

## Protección IP (.NET IL decompilable)

- ILSpy/dnSpy recuperan ~100% del código sin medidas. NativeAOT no es opción (WPF).
- **Recomendado: .NET Reactor (249 $ perpetua)** — virtualización de código + control-flow + string encryption + anti-tamper + licensing; buena compatibilidad WPF.
- Alternativas: ArmDot (499 $, virtualización), Eazfuscator, Obfuscar (gratis, issues con bindings XAML), ConfuserEx2 (apunta a .NET Framework viejo).
- Ofuscación eleva coste de reversing, no lo impide. Testear runtime (bindings XAML + reflexión WPF rompen con renombrado agresivo — usar exclusiones).
- ReadyToRun NO protege (los metadatos IL siguen).

## Updates

**Velopack** (delta + full, ruta exe fija) > MSIX `.appinstaller` (solo si MSIX firmado) > custom (evitar). Squirrel/Clowd.Squirrel: migrar.

## Entra app — publisher verification (crítico para esta app)

- Scopes de alto privilegio (admin M365) + **multitenant** ⇒ **Publisher Verification** casi obligatoria: apps multitenant post-nov-2020 con permisos más allá de sign-in no reciben consentimiento de usuarios normales sin verified publisher (step-up consent).
- Publisher Verification: **gratis**, requiere cuenta **Microsoft AI Cloud Partner Program (CPP/ex-MPN)** verificada + MFA + publisher domain coincidente. El trámite tarda — iniciar pronto.
- **Single-tenant** (app por cliente): más fricción de onboarding, sin requisito de verification, menor blast radius. Apropiado para pilotos.

## Recomendación por fases

**Corto plazo (pilotos):**
- `dotnet publish` self-contained + ReadyToRun (sin trimming/AOT) → **Velopack**.
- Canal: GitHub Releases (repo público solo-releases) o web propia.
- Firma: con entidad legal ES → **Azure Artifact Signing** (~10 $/mes); individual → **OV** Sectigo/Certum (~65–150 €/año). No EV.
- IP: **.NET Reactor** desde el día 1.
- Entra: single-tenant o multitenant + admin consent; iniciar Publisher Verification.
- Avisar a pilotos del aviso SmartScreen inicial.

**Medio plazo (distribución amplia):**
- Canal primario: **WinGet** (`winget install GREX365`), PR automatizado con `wingetcreate` en CI.
- Secundario: **Microsoft Store** (gratis, cero SmartScreen) con MSIX.
- Empresarial: MSI (WiX) para Intune/GPO.
- Firma consolidada en Artifact Signing (company UE verificada).
- Entra: multitenant + publisher verified.

## Decisiones tomadas (usuario, 2026-06-05)

1. **Firma: coste CERO** — no se paga certificado ni servicio.
2. **App Entra: single-tenant** (cada cliente registra su app; la herramienta ya auto-crea el App Reg). Publisher Verification **no necesaria** con single-tenant (y de todas formas exigiría MFA en la cuenta Partner).
3. **Microsoft Store: descartada** (evitar KYC).
4. Cuenta Microsoft Partner: se crearía gratis solo si algún día hace falta (multitenant futuro).

### Estrategia resultante (coste 0 €)

| Ámbito | Solución | Coste | Avisos |
|---|---|---|---|
| **Dentro de Andersen** (máquinas del dominio) | Cert de code-signing emitido por la **CA interna (AD CS)** o self-signed desplegado por **GPO/Intune** a Trusted Publishers | 0 € | **Ninguno** en máquinas gestionadas |
| **Fuera** (pilotos externos) | Binario **sin firmar** — avisar del SmartScreen ("Más información → Ejecutar de todas formas") | 0 € | SmartScreen sí (asumido) |
| Canal | **GitHub Releases en repo público solo-releases** (sin árbol de fuente) — Velopack puede usarlo como feed de updates nativo | 0 € | — |
| Empaquetado + auto-update | **Velopack** (MIT) | 0 € | — |
| IP | **Obfuscar** (OSS gratis) con exclusiones para bindings XAML — protección limitada (renombrado básico); .NET Reactor queda como upgrade futuro si algún día hay presupuesto | 0 € | — |
| Canal futuro opcional | **WinGet** (gratis, sin KYC) — ojo: instalador sin firma puede fallar la validación de malware del PR; mejor cuando haya firma interna o reputación | 0 € | — |

**Limitaciones aceptadas del plan a coste cero**: SmartScreen/AV avisarán fuera de máquinas gestionadas; Smart App Control (Win11 estricto) puede bloquear sin firma; Obfuscar protege menos que un ofuscador comercial con virtualización.

### Próximos pasos concretos
1. Integrar **Velopack** (paquete NuGet + `VelopackApp.Build().Run()` en startup + script `vpk pack` en packaging/) con feed GitHub Releases.
2. Emitir cert de code-signing en la **CA interna de Andersen** (AD CS) y firmar el instalador en CI/local (`signtool`); desplegar confianza por GPO si hiciera falta.
3. Añadir **Obfuscar** al pipeline de publish con exclusiones XAML + smoke test runtime de la app ofuscada.
4. Repo público `grex365-releases` (solo binarios) cuando haya primer piloto externo.

## Fuentes

- Store individual gratis: https://blogs.windows.com/windowsdeveloper/2025/09/10/free-developer-registration-for-individual-developers-on-microsoft-store/
- Store company gratis: https://blogs.windows.com/windowsdeveloper/2026/05/07/publish-to-microsoft-store-as-a-company-now-with-free-registration-and-faster-onboarding/
- Verificación company: https://learn.microsoft.com/en-us/windows/apps/publish/store-business-verification-reqs
- WinGet submit: https://learn.microsoft.com/en-us/windows/package-manager/package/repository · https://github.com/microsoft/winget-create
- Chocolatey moderation: https://docs.chocolatey.org/en-us/community-repository/moderation/
- GitHub private assets: https://github.com/orgs/community/discussions/47453
- Artifact Signing pricing/FAQ: https://azure.microsoft.com/en-us/pricing/details/artifact-signing/ · https://learn.microsoft.com/en-us/azure/artifact-signing/faq
- SmartScreen reputation: https://learn.microsoft.com/en-us/windows/apps/package-and-deploy/smartscreen-reputation
- Code signing options: https://learn.microsoft.com/en-us/windows/apps/package-and-deploy/code-signing-options
- MSIX signing: https://learn.microsoft.com/en-us/windows/msix/package/sign-msix-package-guide
- Velopack: https://github.com/velopack/velopack · https://docs.velopack.io/
- Trimming WPF incompat: https://learn.microsoft.com/en-us/dotnet/core/deploying/trimming/incompatibilities
- NativeAOT: https://learn.microsoft.com/en-us/dotnet/core/deploying/native-aot/
- .NET Reactor: https://www.eziriz.com/order.htm · ArmDot: https://www.armdot.com/
- Obfuscadores: https://www.softanics.com/net-obfuscation/tools · https://github.com/NotPrab/.NET-Obfuscator
- Publisher verification: https://learn.microsoft.com/en-us/entra/identity-platform/publisher-verification-overview
