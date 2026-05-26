# GREX365 ${TAG_NAME}

> Reemplaza este preámbulo por 2-3 líneas describiendo el foco del release (feature mayor / fix crítico / hardening).

## Highlights

- 🚀 **Feature** — _resume en 1 línea_
- 🛠️ **Mejora** — _resume en 1 línea_
- 🐛 **Fix** — _resume en 1 línea_

## Detalle completo

Ver [CHANGELOG.md](https://github.com/4leX-42/GREX365/blob/${TAG_NAME}/CHANGELOG.md#${TAG_NAME_ANCHOR}).

## Artefactos

| Asset | Uso |
|-------|-----|
| `Grex365-${TAG_NAME}-portable/Grex365.App.exe` | EXE single-file portable. Sin instalación, requiere .NET 10 runtime. |
| `Grex365-${TAG_NAME}-msix/Grex365.msix` | Paquete MSIX firmado (si secret `SIGN_CERT_PFX_B64` configurado). Instalable via Intune/SCCM. |
| `Grex365-${TAG_NAME}-msix/Grex365.appinstaller` | Manifest auto-update. Apuntar al feed `MSIX_FEED_BASE_URI`. |

## Instalación

**Portable** (sin admin):
```pwsh
# Descarga Grex365.App.exe, ponlo en una carpeta y ejecuta. Necesita .NET 10 runtime instalado.
winget install Microsoft.DotNet.Runtime.10
.\Grex365.App.exe
```

**MSIX** (recomendado para enterprise):
```pwsh
# Doble-clic en .appinstaller para registrar el feed de auto-update, o:
Add-AppxPackage -Path Grex365.msix
```

## Verificación

- Build artifacts: ver pestaña _Actions_ → workflow CI run del tag
- Test report: artifact `test-results` (`test-results.trx`) en el run de CI
- Coverage: tests pasan en commit base — sin tests deshabilitados

## Tenants soportados

- Microsoft Graph + Exchange Online (Microsoft 365 / Office 365)
- Auth: cert (app-only) o device code (delegado)

## Breaking changes

_Nota explícita si alguna llave de preferencias / esquema de cert / nav titles cambió. Caso típico: rename de modulo en navegación (cubierto por `NavTitleMigrator.RenameMap`)._

## Próximos pasos

- _link a Issues / Discussions / ROADMAP.md_

---
🤖 Notas generadas a partir de `.github/release-template.md` + sección del CHANGELOG correspondiente.
