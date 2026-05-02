#requires -Version 5.1
[CmdletBinding()]
param()

# --- BOOTSTRAP ---

try {
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction Stop
} catch {}

$Script:LauncherPath = Split-Path -Parent $MyInvocation.MyCommand.Path

if (Test-Path (Join-Path $Script:LauncherPath 'GREX365')) {
    $Script:BasePath = Join-Path $Script:LauncherPath 'GREX365'
}
else {
    $Script:BasePath = $Script:LauncherPath
}

$Script:ModulesPath  = Join-Path $Script:BasePath 'Modules'
$Script:ScriptsPath  = Join-Path $Script:BasePath 'Scripts'
$Script:ConfigPath   = Join-Path $Script:BasePath 'config'
$Script:CertCsvPath  = Join-Path $Script:BasePath 'cert_instrunciones\EXO_Cert_Auth_Pasos.csv'
$Script:CsvDocPath   = Join-Path $Script:LauncherPath 'Instrucciones_CSV.html'

foreach ($p in @($Script:ConfigPath)) {
    if (-not (Test-Path -LiteralPath $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null }
}

if (-not (Test-Path -LiteralPath $Script:CsvDocPath)) {
    Write-Host "  [aviso] No se encontró Instrucciones_CSV.html en: $Script:CsvDocPath" -ForegroundColor DarkYellow
}

Get-ChildItem -Path $Script:BasePath -Recurse -Filter *.ps1 -File -ErrorAction SilentlyContinue |
    Unblock-File -ErrorAction SilentlyContinue

$modulesToLoad = @(
    Join-Path $Script:ModulesPath 'Common.ps1'
    Join-Path $Script:ModulesPath 'Preferences.ps1'
    Join-Path $Script:ModulesPath 'Ui.ps1'
    Join-Path $Script:ModulesPath 'Connect-Services.ps1'
    Join-Path $Script:ModulesPath 'Cert-Wizard.ps1'
)

foreach ($m in $modulesToLoad) {
    if (-not (Test-Path -LiteralPath $m)) {
        Write-Host ''
        Write-Host "No se encontró el módulo requerido: $m" -ForegroundColor Red
        Write-Host ''
        exit 1
    }
    . $m
}

# --- DEFINICIÓN DEL MENÚ ---

$Script:MenuItems = @(
    @{
        Label       = 'INYECCIÓN DE USUARIOS // 365-DL'
        Script      = 'Add-MembersToGroup_Fixed.ps1'
        NeedsGraph  = $true
        NeedsExo    = $true
        Scopes      = @('User.Read.All','Group.Read.All','GroupMember.ReadWrite.All')
    }
    @{
        Label       = 'EXTRACCIÓN DE USUARIOS // EMAIL-ID'
        Script      = 'extraccion_user.ps1'
        NeedsGraph  = $true
        NeedsExo    = $true
        Scopes      = @('User.Read.All','Group.Read.All')
    }
    @{
        Label       = 'CREAR GRUPOS / DL DESDE CSV'
        Script      = 'New-GroupsFromCsv.ps1'
        NeedsGraph  = $true
        NeedsExo    = $true
        Scopes      = @('User.Read.All','Group.ReadWrite.All','GroupMember.ReadWrite.All')
    }
    @{
        Label       = 'CORREGIR SHAREDMAILBOX → USERMAILBOX (TEAMS)'
        Script      = 'Fix-SharedMailbox.ps1'
        NeedsGraph  = $false
        NeedsExo    = $true
        Scopes      = @()
    }
    @{
        Label       = 'ASISTENTE DE CREACIÓN DE CERTIFICADO (ExO + Graph)'
        Action      = 'cert-wizard'
        NeedsGraph  = $false
        NeedsExo    = $false
    }
    @{
        Label       = 'PREFERENCIAS / MÉTODO DE CONEXIÓN'
        Action      = 'preferences'
        NeedsGraph  = $false
        NeedsExo    = $false
    }
)

# --- LANZADORES DE SCRIPTS ---

function Invoke-ToolkitScript {
    param(
        [Parameter(Mandatory = $true)][hashtable]$MenuItem
    )

    $scriptPath = Join-Path $Script:ScriptsPath $MenuItem.Script

    if (-not (Test-Path -LiteralPath $scriptPath)) {
        Show-Header -Title $MenuItem.Label -Subtitle 'SCRIPT NO DISPONIBLE'
        Write-Log "No se encontró el script: $scriptPath" 'WARN'
        Wait-ForMenuReturn
        return
    }

    try {
        $needsGraph = [bool]$MenuItem.NeedsGraph
        $needsExo   = [bool]$MenuItem.NeedsExo

        if ($needsGraph -or $needsExo) {
            Connect-RequiredServices -MgGraph:$needsGraph -ExchangeOnline:$needsExo -GraphScopes $MenuItem.Scopes
        }
        else {
            Write-Log "Este script no requiere conexión a servicios M365." 'INFO'
        }

        & $scriptPath
    }
    catch {
        $msg = $_.Exception.Message
        if ($msg -in @('INPUT_EMPTY','INPUT_INVALID','INPUT_NOTFOUND')) {
            # Ya mostrado por el script con Show-ErrorPanel.
        }
        elseif (Get-Command Show-ErrorPanel -ErrorAction SilentlyContinue) {
            Show-Header -Title $MenuItem.Label -Subtitle 'ERROR'
            Show-ErrorPanel -Title $MenuItem.Label -Reason $msg
        }
        else {
            Show-Header -Title $MenuItem.Label -Subtitle 'ERROR'
            Write-Log $msg 'ERROR'
        }
    }

    Wait-ForMenuReturn
}

function Invoke-PreferencesMenu {
    while ($true) {
        Show-Header -Title 'PREFERENCIAS' -Subtitle 'Método de conexión y configuración'

        $prefs = Get-UserPreferences
        Write-Centered -Text ("Método actual:    {0}" -f (Format-MethodLabel -Method $prefs.ConnectionMethod)) -Color 'Cyan'
        if ($prefs.TraditionalAdminUpn) {
            Write-Centered -Text ("Admin tradicional: {0}" -f $prefs.TraditionalAdminUpn) -Color 'Gray'
        }
        $certStatus = if (Test-CertConfigExists) { 'SÍ' } else { 'NO' }
        Write-Centered -Text ("Cert configurado: {0}" -f $certStatus) -Color 'Gray'

        if (Get-Command Show-StatusPanel -ErrorAction SilentlyContinue) {
            Show-StatusPanel
        }

        Write-Centered -Text '[1] Cambiar método de conexión' -Color 'White'
        Write-Centered -Text '[2] Cambiar UPN de admin tradicional' -Color 'White'
        Write-Centered -Text '[3] Resetear preferencias (vuelve a primera ejecución)' -Color 'White'
        if ($certStatus -eq 'SÍ') {
            Write-Centered -Text '[4] Eliminar certificado configurado (DESTRUCTIVO)' -Color 'Red'
        }
        Write-Centered -Text '[0] Volver' -Color 'DarkGray'
        Write-Host ''

        $opt = Read-Host 'Opción'
        switch ($opt) {
            '1' {
                $newMethod = Show-MethodSelector
                if ($newMethod) {
                    Set-PreferenceValue -Key 'ConnectionMethod' -Value $newMethod
                    Write-Log ("Método cambiado a: {0}" -f $newMethod) 'OK'
                    Start-Sleep -Milliseconds 800
                }
            }
            '2' {
                $upn = Read-Host 'Nuevo UPN de administrador'
                if (-not [string]::IsNullOrWhiteSpace($upn)) {
                    Set-PreferenceValue -Key 'TraditionalAdminUpn' -Value $upn.Trim()
                    Write-Log "UPN actualizado." 'OK'
                    Start-Sleep -Milliseconds 800
                }
            }
            '3' {
                $confirm = Read-Host '¿Seguro que quieres resetear preferencias? (S/N)'
                if ($confirm -match '^[Ss]') {
                    $defaults = New-DefaultPreferences
                    Save-UserPreferences -Preferences $defaults
                    Write-Log "Preferencias reseteadas." 'OK'
                    Start-Sleep -Milliseconds 800
                }
            }
            '4' {
                if ($certStatus -ne 'SÍ') { continue }
                Invoke-DeleteCertificateFlow
            }
            '0' { return }
            default { }
        }
    }
}

function Invoke-DeleteCertificateFlow {
    Write-Host ''
    Write-Host '  ╔════════════════════════════════════════════════════════════════════╗' -ForegroundColor Red
    Write-Host '  ║   ATENCIÓN — ACCIÓN DESTRUCTIVA                                    ║' -ForegroundColor Red
    Write-Host '  ╚════════════════════════════════════════════════════════════════════╝' -ForegroundColor Red
    Write-Host ''
    Write-Host '  Se va a eliminar:' -ForegroundColor Yellow
    Write-Host '    - El certificado del almacén CurrentUser\My (con su clave privada)' -ForegroundColor Yellow
    Write-Host '    - El archivo .cer público exportado' -ForegroundColor Yellow
    Write-Host '    - El JSON exo-app-params.json (parámetros de la App)' -ForegroundColor Yellow
    Write-Host ''
    Write-Host '  La App Registration en Entra ID NO se elimina automáticamente.' -ForegroundColor DarkYellow
    Write-Host '  Tendrás que borrarla a mano si quieres limpiar el tenant.' -ForegroundColor DarkYellow
    Write-Host ''

    $first = Read-Host '  ¿Continuar con la eliminación? (S/N)'
    if ($first -notmatch '^[Ss]') {
        Write-Log 'Eliminación cancelada.' 'WARN'
        Start-Sleep -Milliseconds 800
        return
    }

    Write-Host ''
    Write-Host '  Confirmación final. Para proceder, escribe literalmente la palabra: ' -ForegroundColor Red -NoNewline
    Write-Host 'CONFIRMAR' -ForegroundColor White
    $second = Read-Host '  >'
    if ($second -ne 'CONFIRMAR') {
        Write-Log 'Texto de confirmación incorrecto. Eliminación abortada.' 'WARN'
        Start-Sleep -Milliseconds 1200
        return
    }

    try {
        $ok = Remove-CertConfig
        if ($ok) {
            Write-Log 'Certificado y configuración eliminados correctamente.' 'OK'
        }
        else {
            Write-Log 'Eliminación parcial. Revisa los mensajes anteriores.' 'WARN'
        }
    }
    catch {
        Write-Log ("Error eliminando certificado: {0}" -f $_.Exception.Message) 'ERROR'
    }
    Start-Sleep -Seconds 2
}

function Invoke-CertWizardLauncher {
    if (Test-CertConfigExists) {
        Show-Header -Title 'ASISTENTE DE CERTIFICADO' -Subtitle 'Configuración existente detectada'
        $cfg = Get-CertConfig
        Write-Centered -Text "Ya tienes un certificado válido configurado:" -Color 'Green'
        Write-Host ''
        Write-Centered -Text ("AppId:        {0}" -f $cfg.AppId) -Color 'Gray'
        Write-Centered -Text ("Thumbprint:   {0}" -f $cfg.CertThumbprint) -Color 'Gray'
        Write-Centered -Text ("Tenant:       {0}" -f $cfg.TenantId) -Color 'Gray'
        Write-Centered -Text ("Organization: {0}" -f $cfg.Organization) -Color 'Gray'
        Write-Host ''
        $r = Read-Host '¿Quieres rehacer el asistente de todas formas? (S/N)'
        if ($r -notmatch '^[Ss]') { return }
    }

    try {
        Start-CertificateWizard -CsvStepsPath $Script:CertCsvPath -ConfigPath (Join-Path $Script:ConfigPath 'exo-app-params.json')
    }
    catch {
        Write-Log $_.Exception.Message 'ERROR'
    }

    Wait-ForMenuReturn
}

# --- FLUJO PRINCIPAL ---

function Initialize-FirstRun {
    $prefs = Get-UserPreferences

    if ($prefs.FirstRunCompleted -and $prefs.ConnectionMethod) {
        return
    }

    $method = Show-MethodSelector
    if (-not $method) {
        Write-Log "Selección cancelada. Saliendo." 'WARN'
        exit 0
    }

    Set-PreferenceValue -Key 'ConnectionMethod' -Value $method
    Set-PreferenceValue -Key 'FirstRunCompleted' -Value $true

    if ($method -eq 'cert' -and -not (Test-CertConfigExists)) {
        Show-Header -Title 'PRIMERA EJECUCIÓN' -Subtitle 'No hay certificado configurado'
        Write-Centered -Text 'Has elegido método CERTIFICADO pero no hay configuración válida.' -Color 'Yellow'
        Write-Centered -Text '¿Quieres lanzar ahora el asistente de creación de certificado?' -Color 'White'
        Write-Host ''
        $r = Read-Host '(S/N)'
        if ($r -match '^[Ss]') {
            Invoke-CertWizardLauncher
        }
        else {
            Write-Centered -Text 'Puedes lanzarlo más tarde desde el Main menu' -Color 'DarkGray'
            Start-Sleep -Seconds 2
        }
    }
}

function Get-MenuItemsForUi {
    return @($Script:MenuItems | ForEach-Object { @{ Label = $_.Label; ComingSoon = $false } })
}

try {
    Initialize-FirstRun

    while ($true) {
        $prefs = Get-UserPreferences
        $items = Get-MenuItemsForUi
        $result = Show-MainMenu -Items $items -ActiveMethod $prefs.ConnectionMethod

        switch ($result.Action) {
            'exit' {
                Show-Header -Title 'GREX365' -Subtitle 'Hasta luego.'
                exit 0
            }
            'toggle' {
                if ($result.Value -ne $prefs.ConnectionMethod) {
                    Set-PreferenceValue -Key 'ConnectionMethod' -Value $result.Value
                    Write-Log ("Método cambiado a: {0}" -f $result.Value) 'OK'
                    Start-Sleep -Milliseconds 700
                }
            }
            'option' {
                $idx = [int]$result.Value - 1
                if ($idx -lt 0 -or $idx -ge $Script:MenuItems.Count) { continue }
                $item = $Script:MenuItems[$idx]

                if ($item.ContainsKey('Action')) {
                    switch ($item.Action) {
                        'cert-wizard' { Invoke-CertWizardLauncher }
                        'preferences' { Invoke-PreferencesMenu }
                        default { }
                    }
                }
                else {
                    Invoke-ToolkitScript -MenuItem $item
                }
            }
        }
    }
}
catch {
    Write-Host ''
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ''
    exit 1
}
