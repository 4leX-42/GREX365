# Changelog · GREX365 v2.0

Formato: [Keep a Changelog](https://keepachangelog.com/es/1.1.0/) · SemVer.

> **Scope**: rewrite .NET 10 / WPF de la app (`Grex365.App` + `Grex365.Core` + `Grex365.PowerShell`). Para el toolkit legacy PowerShell, ver `GREX365/Seguimiento Claude/CHANGELOG.md`.

## [Unreleased]

### Added
- Internacionalización ES + EN — 20/20 superficies con `L10nExtension` + `Settings.LanguageLabel` switcher. Tests parity ES↔EN para todas las keys (~557 L10n tests). Sprints AA–AG.
- `BoolToOnOffConverter` admite L10n keys (parámetros dotted) — sin strings hardcoded ES en converters.
- Fail-soft en config corrupto: `JsonPreferencesStore` y `JsonCertConfigStore` cuarentenan el archivo dañado a `*.corrupted-yyyyMMddHHmmss.bak` y arrancan con defaults en lugar de crashear. Sprint AH (H1.5.5).
- Window state restore con `WindowPlacementGuard` (off-screen safety en RDP / multi-monitor disconnect).
- Tray icon + minimize-to-tray + reconexión silenciosa (Sprint X+Y).

### Changed
- Paleta corporativa azul `#1E40AF → #2563EB → #3B82F6` (anteriormente púrpura). Theme-aware contrast pass en 17 XAMLs (Sprint V+W).
- `DashboardViewModel.GoTo` match por stable `NavKey` primero, fallback a `Title` localizado — fix silent breakage tras nav rename `Salud tenant`→`Licencias`.
- WFO0003 (WinForms DPI advisory) suppressed en `Grex365.App.csproj` — app.manifest sigue canónico para WPF.

### Removed
- 13 PNG screenshots de root — capturas dev ad-hoc sin referencias (Sprint AA).

## How to cut a release

```pwsh
# 1. Update [Unreleased] → [X.Y.Z] - YYYY-MM-DD en este archivo
# 2. Bump <Version> en Directory.Build.props
# 3. Commit: chore(release): bump to vX.Y.Z
# 4. Tag + push: git tag vX.Y.Z && git push origin vX.Y.Z
# 5. CI dispara: build → publish portable EXE → msix → (opcional) signtool
# 6. Job `release` crea el GH Release usando .github/release-template.md como body
```

---

## [0.2.0-alpha] - 2026-05-23

### Added
- 13 security audits con report HTML/JSON/CSV + baseline diff
- RBAC vía Entra group membership en VMs destructivos
- Auto AppRegistration vía Graph con cert upload + admin-consent URL
- Onboarding wizard (UPN/password/usageLocation/SKUs/groups)
- Offboarding wizard (disable + remove licenses + convert mailbox to shared)
- Plugin system (`IModule` + AssemblyLoadContext) + sample plugin
- MSIX scaffold + CI release job + `.appinstaller` template

### Changed
- Rewrite legacy PowerShell GUI → WPF + Fluent (WPF-UI 4.3.0)
- Tenant lock enforced post-auth (cert + device-code)
- App-only auth bypassa RBAC by design (sin contexto `/me/checkMemberGroups`)
