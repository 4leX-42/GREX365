#requires -Version 7.4
[CmdletBinding()]
param()

# --- BOOTSTRAP ---

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host ''
    Write-Host '  GREX365 requiere PowerShell 7.4 o superior.' -ForegroundColor Red
    Write-Host ('  Versión detectada: {0}' -f $PSVersionTable.PSVersion) -ForegroundColor Yellow
    Write-Host ''
    Write-Host '  Instala PowerShell 7+ desde:' -ForegroundColor Yellow
    Write-Host '    winget install --id Microsoft.PowerShell' -ForegroundColor DarkGray
    Write-Host '    o https://aka.ms/powershell' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host '  Después ejecuta el toolkit con: pwsh .\Main.ps1' -ForegroundColor Yellow
    Write-Host ''
    exit 1
}

try {
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction Stop
} catch {}

# Minimal progress bar style (cleaner Write-Progress lines on PS7+).
try { $PSStyle.Progress.View = 'Minimal' } catch {}

# Force UTF-8 output (renders box-drawing + acentos correctamente en consola Windows).
try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = [System.Text.Encoding]::UTF8
} catch {}

$Script:LauncherPath = Split-Path -Parent $MyInvocation.MyCommand.Path

if (Test-Path (Join-Path $Script:LauncherPath 'GREX365')) {
    $Script:BasePath = Join-Path $Script:LauncherPath 'GREX365'
} else {
    $Script:BasePath = $Script:LauncherPath
}

# Exposed as $global: so child scripts invoked via & inherit the toolkit context.
$global:GREX365_BasePath = $Script:BasePath

$Script:ModulesPath = Join-Path $Script:BasePath  'Modules'
$Script:ScriptsPath = Join-Path $Script:BasePath  'Scripts'
$Script:ConfigPath  = Join-Path $Script:BasePath  'config'
$Script:LogsPath    = Join-Path $Script:BasePath  'logs'
$Script:DocsPath    = Join-Path $Script:LauncherPath 'docs'
$Script:CertCsvPath = Join-Path $Script:DocsPath  'Certificate-Setup-Steps.csv'
$Script:CsvDocPath  = Join-Path $Script:DocsPath  'CSV-Schemas.html'

foreach ($p in @($Script:ConfigPath, $Script:LogsPath)) {
    if (-not (Test-Path -LiteralPath $p)) {
        New-Item -ItemType Directory -Path $p -Force | Out-Null
    }
}

Get-ChildItem -Path $Script:BasePath -Recurse -Filter *.ps1 -File -ErrorAction SilentlyContinue |
    Unblock-File -ErrorAction SilentlyContinue

$Script:ModulesToLoad = @(
    'Logging.ps1'
    'Console.ps1'
    'Validation.ps1'
    'Csv.ps1'
    'Preferences.ps1'
    'Retry.ps1'
    'Audit.ps1'
    'Report.ps1'
    'Roles.ps1'
    'Templates.ps1'
    'Jobs.ps1'
    'Connection.ps1'
    'GroupResolver.ps1'
    'CertWizard.ps1'
    'Menu.ps1'
)

foreach ($name in $Script:ModulesToLoad) {
    $modulePath = Join-Path $Script:ModulesPath $name
    if (-not (Test-Path -LiteralPath $modulePath)) {
        Write-Host ''
        Write-Host ('No se encontró el módulo requerido: ' + $modulePath) -ForegroundColor Red
        Write-Host ''
        exit 1
    }
    . $modulePath
}

# --- MENU DEFINITION ---

$Script:MenuItems = @(
    @{
        Section     = 'Operaciones'
        Label       = 'Workflow de grupos (crear + miembros + permisos)'
        Tag         = 'Graph + EXO'
        Script      = 'Invoke-GroupsWorkflow.ps1'
        NeedsGraph  = $true
        NeedsExo    = $true
        Scopes      = @('User.Read.All','Group.ReadWrite.All','GroupMember.ReadWrite.All')
    }
    @{
        Section     = 'Operaciones'
        Label       = 'Exportar miembros de grupo/DL'
        Tag         = 'Graph + EXO'
        Script      = 'Export-GroupMembers.ps1'
        NeedsGraph  = $true
        NeedsExo    = $true
        Scopes      = @('User.Read.All','Group.Read.All')
    }
    @{
        Section     = 'Operaciones'
        Label       = 'Convertir SharedMailbox a UserMailbox'
        Tag         = 'EXO'
        Script      = 'Convert-SharedToUserMailbox.ps1'
        NeedsGraph  = $false
        NeedsExo    = $true
        Scopes      = @()
    }
    @{
        Section     = 'Operaciones'
        Label       = 'Permisos sobre buzón (bulk CSV)'
        Tag         = 'EXO'
        Script      = 'Set-SharedMailboxPermissions.ps1'
        NeedsGraph  = $false
        NeedsExo    = $true
        Scopes      = @()
    }
    @{
        Section     = 'Legacy'
        Label       = 'Agregar miembros (solo)'
        Tag         = 'usar el Workflow en su lugar'
        Script      = 'Add-GroupMembers.ps1'
        NeedsGraph  = $true
        NeedsExo    = $true
        Scopes      = @('User.Read.All','Group.Read.All','GroupMember.ReadWrite.All')
    }
    @{
        Section     = 'Legacy'
        Label       = 'Crear grupos/DL (solo)'
        Tag         = 'usar el Workflow en su lugar'
        Script      = 'New-GroupsFromCsv.ps1'
        NeedsGraph  = $true
        NeedsExo    = $true
        Scopes      = @('User.Read.All','Group.ReadWrite.All','GroupMember.ReadWrite.All')
    }
    @{
        Section     = 'Workflows'
        Label       = 'Offboarding wizard (14 pasos)'
        Tag         = 'Graph + EXO · admin'
        Script      = 'Invoke-OffboardingWizard.ps1'
        NeedsGraph  = $true
        NeedsExo    = $true
        Scopes      = @('User.ReadWrite.All','Group.ReadWrite.All','Directory.AccessAsUser.All','UserAuthenticationMethod.ReadWrite.All','Mail.ReadWrite')
    }
    @{
        Section     = 'Auditoría'
        Label       = 'Dashboard de salud del tenant'
        Tag         = 'Graph + EXO · read-only'
        Script      = 'Show-TenantHealth.ps1'
        NeedsGraph  = $true
        NeedsExo    = $true
        Scopes      = @('Directory.Read.All','Reports.Read.All','Application.Read.All','AuditLog.Read.All','ServiceHealth.Read.All')
    }
    @{
        Section     = 'Auditoría'
        Label       = 'Identity audit (stale + huérfanos)'
        Tag         = 'Graph · read-only'
        Script      = 'Invoke-IdentityAudit.ps1'
        NeedsGraph  = $true
        NeedsExo    = $false
        Scopes      = @('User.Read.All','Group.Read.All','AuditLog.Read.All')
    }
    @{
        Section     = 'Auditoría'
        Label       = 'Self-test sobre objetos testeo*'
        Tag         = 'Graph + EXO · crea y borra DL temporal'
        Script      = 'Invoke-SelfTest.ps1'
        NeedsGraph  = $true
        NeedsExo    = $true
        Scopes      = @('User.Read.All','Group.ReadWrite.All','GroupMember.ReadWrite.All')
    }
    @{
        Section     = 'Configuración'
        Label       = 'Asistente de certificado (ExO + Graph)'
        Tag         = ''
        Action      = 'cert-wizard'
        NeedsGraph  = $false
        NeedsExo    = $false
    }
    @{
        Section     = 'Configuración'
        Label       = 'Preferencias'
        Tag         = ''
        Action      = 'preferences'
        NeedsGraph  = $false
        NeedsExo    = $false
    }
    @{
        Section     = 'Configuración'
        Label       = 'Jobs en background'
        Tag         = ''
        Action      = 'jobs'
        NeedsGraph  = $false
        NeedsExo    = $false
    }
)

# --- LAUNCHERS ---

function Invoke-ToolkitScript {
    param([Parameter(Mandatory = $true)][hashtable]$MenuItem)

    $scriptPath = Join-Path $Script:ScriptsPath $MenuItem.Script

    if (-not (Test-Path -LiteralPath $scriptPath)) {
        Show-Header -Title 'GREX365' -Subtitle $MenuItem.Label
        Show-ErrorBlock -Title 'Script no disponible' -Detail $scriptPath
        Wait-ForKey
        return
    }

    $auditMeta = @{
        label  = $MenuItem.Label
        script = $MenuItem.Script
        needsGraph = [bool]$MenuItem.NeedsGraph
        needsExo   = [bool]$MenuItem.NeedsExo
    }
    $auditCtx = $null
    $auditResult = 'OK'
    try { $auditCtx = Start-AuditOperation -Operation $MenuItem.Label -Metadata $auditMeta } catch {}

    try {
        $needsGraph = [bool]$MenuItem.NeedsGraph
        $needsExo   = [bool]$MenuItem.NeedsExo

        if ($needsGraph -or $needsExo) {
            Show-Header -Title 'GREX365' -Subtitle ('Conectando · ' + $MenuItem.Label) -ActiveMethod (Get-UserPreferences).ConnectionMethod
            Connect-RequiredServices -MgGraph:$needsGraph -ExchangeOnline:$needsExo -GraphScopes $MenuItem.Scopes
        } else {
            Write-Log 'Este script no requiere conexión a servicios.' -Source 'Main'
        }

        & $scriptPath
    } catch {
        $msg = $_.Exception.Message
        if ($msg -in @('INPUT_EMPTY','INPUT_INVALID','INPUT_NOTFOUND')) {
            $auditResult = 'CANCELLED'
        } else {
            $auditResult = 'ERROR'
            Show-Header -Title 'GREX365' -Subtitle $MenuItem.Label
            Show-ErrorBlock -Title $MenuItem.Label -Detail $msg
            try { Write-AuditEvent -EventType 'Exception' -Properties @{ message = $msg } } catch {}
        }
    } finally {
        try { Stop-AuditOperation -Result $auditResult } catch {}
    }

    Wait-ForKey
}

function Invoke-PreferencesMenu {
    while ($true) {
        $prefs = Get-UserPreferences

        Show-Header -Title 'GREX365' -Subtitle 'Preferencias' -ActiveMethod $prefs.ConnectionMethod

        $lockEnabled = $false
        $expectedTid = $null
        if ($prefs.PSObject.Properties.Name -contains 'EnforceTenantLock') { $lockEnabled = [bool]$prefs.EnforceTenantLock }
        if ($prefs.PSObject.Properties.Name -contains 'ExpectedTenantId')  { $expectedTid = [string]$prefs.ExpectedTenantId }
        $lockStatus = if ($lockEnabled -and $expectedTid) { 'Activo' } elseif ($expectedTid) { 'Configurado (desactivado)' } else { 'No configurado' }

        $role   = Get-CurrentRole
        $uiMode = Get-CurrentUIMode

        Show-Section -Title 'Configuración actual'
        Write-KeyValue -Key 'Método activo'  -Value (Format-MethodLabel -Method $prefs.ConnectionMethod)
        Write-KeyValue -Key 'Admin UPN'      -Value ($prefs.TraditionalAdminUpn ? $prefs.TraditionalAdminUpn : '—')
        $certStatus = if (Test-CertConfigExists) { 'Configurado' } else { 'No configurado' }
        Write-KeyValue -Key 'Certificado'    -Value $certStatus
        Write-KeyValue -Key 'Rol activo'     -Value $role
        Write-KeyValue -Key 'UI Mode'        -Value $uiMode
        Write-KeyValue -Key 'Tenant lock'    -Value $lockStatus
        if ($expectedTid) {
            Write-KeyValue -Key 'Tenant esperado' -Value $expectedTid
        }

        if (Get-Command Show-StatusPanel -ErrorAction SilentlyContinue) {
            Show-StatusPanel
        }

        Show-Section -Title 'Acciones'
        Write-Indent -Level 2; Write-Host '1   Cambiar método de conexión' -ForegroundColor Gray
        Write-Indent -Level 2; Write-Host '2   Cambiar UPN admin tradicional' -ForegroundColor Gray
        Write-Indent -Level 2; Write-Host '3   Resetear preferencias' -ForegroundColor Gray
        Write-Indent -Level 2; Write-Host '5   Tenant lock (anclar a tenant actual)' -ForegroundColor Gray
        if ($expectedTid) {
            Write-Indent -Level 2; Write-Host '6   Tenant lock — activar/desactivar' -ForegroundColor Gray
            Write-Indent -Level 2; Write-Host '7   Tenant lock — limpiar' -ForegroundColor Gray
        }
        Write-Indent -Level 2; Write-Host '8   Cambiar rol (viewer / operator / admin)' -ForegroundColor Gray
        Write-Indent -Level 2; Write-Host '9   Cambiar UI Mode (support / advanced)' -ForegroundColor Gray
        if ($certStatus -eq 'Configurado') {
            Write-Indent -Level 2; Write-Host '4   Eliminar certificado (DESTRUCTIVO)' -ForegroundColor Red
        }
        Write-Indent -Level 2; Write-Host '0   Volver' -ForegroundColor DarkGray
        Write-Host ''

        $opt = Read-Input -Prompt 'Opción'
        switch ($opt) {
            '1' {
                $newMethod = Show-MethodSelector
                if ($newMethod) {
                    Set-PreferenceValue -Key 'ConnectionMethod' -Value $newMethod
                    Write-Log ('Método cambiado a: ' + $newMethod) -Level OK -Source 'Prefs'
                    Start-Sleep -Milliseconds 700
                }
            }
            '2' {
                $upn = Read-Input -Prompt 'Nuevo UPN de administrador'
                if (-not [string]::IsNullOrWhiteSpace($upn)) {
                    Set-PreferenceValue -Key 'TraditionalAdminUpn' -Value (Normalize-Input -Value $upn)
                    Write-Log 'UPN actualizado.' -Level OK -Source 'Prefs'
                    Start-Sleep -Milliseconds 700
                }
            }
            '3' {
                $confirm = Read-Input -Prompt '¿Resetear preferencias? (S/N)'
                if ($confirm -match '^[Ss]') {
                    $defaults = New-DefaultPreferences
                    Save-UserPreferences -Preferences $defaults
                    Write-Log 'Preferencias reseteadas.' -Level OK -Source 'Prefs'
                    Start-Sleep -Milliseconds 700
                }
            }
            '4' {
                if ($certStatus -ne 'Configurado') { continue }
                Invoke-DeleteCertificateFlow
            }
            '5' {
                $currentTid = Get-CurrentConnectedTenantId
                if (-not $currentTid) {
                    Show-WarningBlock -Title 'No hay tenant conectado' -Detail 'Conecta antes a Graph / EXO para fijar el tenant esperado.'
                    Start-Sleep -Seconds 2
                } else {
                    $state = Get-SessionState
                    Set-PreferenceValue -Key 'ExpectedTenantId'     -Value $currentTid
                    Set-PreferenceValue -Key 'ExpectedTenantDomain' -Value ([string]$state.TenantDomain)
                    Set-PreferenceValue -Key 'EnforceTenantLock'    -Value $true
                    Write-Log ('Tenant lock activado y anclado a ' + $currentTid) -Level OK -Source 'Prefs'
                    Start-Sleep -Milliseconds 900
                }
            }
            '6' {
                if (-not $expectedTid) { continue }
                $newVal = -not $lockEnabled
                Set-PreferenceValue -Key 'EnforceTenantLock' -Value $newVal
                Write-Log ('Tenant lock = ' + $newVal) -Level OK -Source 'Prefs'
                Start-Sleep -Milliseconds 700
            }
            '7' {
                if (-not $expectedTid) { continue }
                Set-PreferenceValue -Key 'ExpectedTenantId'     -Value $null
                Set-PreferenceValue -Key 'ExpectedTenantDomain' -Value $null
                Set-PreferenceValue -Key 'EnforceTenantLock'    -Value $false
                Write-Log 'Tenant lock limpiado.' -Level OK -Source 'Prefs'
                Start-Sleep -Milliseconds 700
            }
            '8' {
                Write-Host ''
                Write-Indent -Level 2; Write-Host '1) viewer   (solo lectura, no destructivas)' -ForegroundColor Gray
                Write-Indent -Level 2; Write-Host '2) operator (operaciones diarias)'           -ForegroundColor Gray
                Write-Indent -Level 2; Write-Host '3) admin    (offboarding, bulk destructivo)' -ForegroundColor Gray
                $r = Read-Input -Prompt 'Rol nuevo'
                $map = @{ '1'='viewer'; '2'='operator'; '3'='admin' }
                if ($map.ContainsKey($r)) {
                    Set-CurrentRole -Role $map[$r]
                    Start-Sleep -Milliseconds 700
                }
            }
            '9' {
                Write-Host ''
                Write-Indent -Level 2; Write-Host '1) support  (wizards, dry-run forzado, confirmaciones extra)' -ForegroundColor Gray
                Write-Indent -Level 2; Write-Host '2) advanced (cmdlets visibles, menos confirmaciones)'         -ForegroundColor Gray
                $m = Read-Input -Prompt 'Modo nuevo'
                $map = @{ '1'='support'; '2'='advanced' }
                if ($map.ContainsKey($m)) {
                    Set-CurrentUIMode -Mode $map[$m]
                    Start-Sleep -Milliseconds 700
                }
            }
            '0' { return }
            default { }
        }
    }
}

function Invoke-JobsMenu {
    Show-Header -Title 'GREX365' -Subtitle 'Jobs en background' -ActiveMethod (Get-UserPreferences).ConnectionMethod

    $jobs = Get-ToolkitJobs
    Show-Section -Title 'Jobs registrados'

    if (-not $jobs -or $jobs.Count -eq 0) {
        Write-Indent -Level 2; Write-Host 'No hay jobs registrados.' -ForegroundColor DarkGray
        Write-Host ''
        Wait-ForKey
        return
    }

    $i = 0
    foreach ($j in $jobs) {
        $i++
        $color = switch ($j.State) {
            'Completed' { 'Green' }
            'Failed'    { 'Red' }
            'Stopped'   { 'Yellow' }
            'Running'   { 'Cyan' }
            default     { 'Gray' }
        }
        Write-Indent -Level 2
        Write-Host ('{0,2}  ' -f $i)        -NoNewline -ForegroundColor DarkGray
        Write-Host ($j.Name.PadRight(40))   -NoNewline -ForegroundColor White
        Write-Host ($j.State.PadRight(12))  -NoNewline -ForegroundColor $color
        Write-Host ('id=' + $j.Id)          -ForegroundColor DarkGray
    }
    Write-Host ''
    Write-Indent -Level 2; Write-Host '1   Limpiar jobs finalizados' -ForegroundColor Gray
    Write-Indent -Level 2; Write-Host '0   Volver' -ForegroundColor DarkGray
    Write-Host ''

    $opt = Read-Input -Prompt 'Opción'
    if ($opt -eq '1') {
        Remove-FinishedJobs
        Write-Log 'Jobs finalizados eliminados.' -Level OK -Source 'Jobs'
        Start-Sleep -Milliseconds 700
    }
}

function Invoke-DeleteCertificateFlow {
    Show-Header -Title 'GREX365' -Subtitle 'Eliminar certificado' -ActiveMethod (Get-UserPreferences).ConnectionMethod

    Show-ErrorBlock -Title 'Acción destructiva' -Detail @"
Se va a eliminar:
  · El certificado de CurrentUser\My (con clave privada)
  · El archivo .cer público exportado
  · El JSON exo-app-params.json

La App Registration en Entra ID NO se elimina automáticamente.
"@

    $first = Read-Input -Prompt '¿Continuar? (S/N)'
    if ($first -notmatch '^[Ss]') {
        Write-Log 'Cancelado.' -Level WARN -Source 'CertDelete'
        Start-Sleep -Milliseconds 700
        return
    }

    Write-Indent
    Write-Host 'Confirmación final. Escribe literalmente ' -NoNewline -ForegroundColor Red
    Write-Host 'CONFIRMAR' -ForegroundColor White
    $second = Read-Input -Prompt '>'
    if ($second -ne 'CONFIRMAR') {
        Show-WarningBlock -Title 'Texto incorrecto. Eliminación abortada.'
        Start-Sleep -Milliseconds 1000
        return
    }

    try {
        $ok = Remove-CertConfig
        if ($ok) { Write-Log 'Certificado eliminado.' -Level OK -Source 'CertDelete' }
        else     { Write-Log 'Eliminación parcial. Revisa mensajes anteriores.' -Level WARN -Source 'CertDelete' }
    } catch {
        Show-ErrorBlock -Title 'Error eliminando certificado' -Detail $_.Exception.Message
    }
    Start-Sleep -Seconds 2
}

function Invoke-CertWizardLauncher {
    if (Test-CertConfigExists) {
        Show-Header -Title 'GREX365' -Subtitle 'Asistente de certificado' -ActiveMethod (Get-UserPreferences).ConnectionMethod
        $cfg = Get-CertConfig
        Show-Section -Title 'Configuración existente detectada'
        Write-KeyValue -Key 'AppId'        -Value $cfg.AppId
        Write-KeyValue -Key 'Thumbprint'   -Value $cfg.CertThumbprint
        Write-KeyValue -Key 'Tenant'       -Value $cfg.TenantId
        Write-KeyValue -Key 'Organization' -Value $cfg.Organization

        Write-Host ''
        $r = Read-Input -Prompt '¿Rehacer el asistente de todas formas? (S/N)'
        if ($r -notmatch '^[Ss]') { return }
    }

    try {
        Start-CertificateWizard -CsvStepsPath $Script:CertCsvPath -ConfigPath (Join-Path $Script:ConfigPath 'exo-app-params.json')
    } catch {
        Show-ErrorBlock -Title 'Asistente fallido' -Detail $_.Exception.Message
    }
    Wait-ForKey
}

function Initialize-FirstRun {
    $prefs = Get-UserPreferences
    if ($prefs.FirstRunCompleted -and $prefs.ConnectionMethod) { return }

    $method = Show-MethodSelector
    if (-not $method) {
        Write-Log 'Selección cancelada. Saliendo.' -Level WARN -Source 'Main'
        exit 0
    }

    Set-PreferenceValue -Key 'ConnectionMethod'  -Value $method
    Set-PreferenceValue -Key 'FirstRunCompleted' -Value $true

    if ($method -eq 'cert' -and -not (Test-CertConfigExists)) {
        Show-Header -Title 'GREX365' -Subtitle 'Primera ejecución'
        Show-WarningBlock -Title 'Certificado no configurado' -Detail 'Has elegido método CERTIFICADO pero no hay configuración válida.'

        $r = Read-Input -Prompt '¿Lanzar ahora el asistente de creación? (S/N)'
        if ($r -match '^[Ss]') { Invoke-CertWizardLauncher }
        else {
            Write-Indent; Write-Host 'Puedes lanzarlo más tarde desde el menú principal.' -ForegroundColor DarkGray
            Start-Sleep -Seconds 2
        }
    }
}

# --- MAIN LOOP ---

try {
    Initialize-FirstRun

    while ($true) {
        $prefs = Get-UserPreferences
        $result = Show-MainMenu -Items $Script:MenuItems -ActiveMethod $prefs.ConnectionMethod

        switch ($result.Action) {
            'exit' {
                Show-Header -Title 'GREX365' -Subtitle 'Hasta luego'
                exit 0
            }
            'toggle' {
                if ($result.Value -ne $prefs.ConnectionMethod) {
                    Set-PreferenceValue -Key 'ConnectionMethod' -Value $result.Value
                    Write-Log ('Método cambiado a: ' + $result.Value) -Level OK -Source 'Main'
                    Start-Sleep -Milliseconds 600
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
                        'jobs'        { Invoke-JobsMenu }
                        default       { }
                    }
                } else {
                    Invoke-ToolkitScript -MenuItem $item
                }
            }
        }
    }
} catch {
    Write-Host ''
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ''
    exit 1
}
