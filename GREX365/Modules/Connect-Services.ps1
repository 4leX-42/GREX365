# --- CONEXIÓN INTELIGENTE A SERVICIOS M365 ---
# Conecta solo los servicios que cada script necesita.
# Soporta dos métodos: certificado (App-only) y tradicional (delegado / device code).

function Test-ToolkitModuleImported {
    param([Parameter(Mandatory = $true)][string]$ModuleName)
    return [bool](Get-Module -Name $ModuleName -ErrorAction SilentlyContinue)
}

function Test-ToolkitModuleInstalled {
    param([Parameter(Mandatory = $true)][string]$ModuleName)
    return [bool](Get-Module -ListAvailable -Name $ModuleName -ErrorAction SilentlyContinue)
}

function Ensure-NuGetProvider {
    if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
        Write-Log "Instalando proveedor NuGet..." 'INFO'
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force -ErrorAction Stop | Out-Null
    }
}

function Ensure-ToolkitModule {
    param([Parameter(Mandatory = $true)][string]$ModuleName)

    if (Test-ToolkitModuleImported -ModuleName $ModuleName) { return }

    if (-not (Test-ToolkitModuleInstalled -ModuleName $ModuleName)) {
        Ensure-NuGetProvider
        Write-Log "Instalando módulo '$ModuleName'..." 'INFO'
        Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }

    # Importación silenciosa: solo reporta si falla.
    Import-Module $ModuleName -ErrorAction Stop -Verbose:$false -WarningAction SilentlyContinue 3>$null 4>$null
}

# --- ESTADO DE SESIÓN ---

function Test-ExchangeOnlineConnected {
    try {
        if (-not (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue)) { return $false }
        $sessions = Get-ConnectionInformation -ErrorAction Stop
        if (-not $sessions) { return $false }
        $active = $sessions | Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1
        return [bool]$active
    }
    catch { return $false }
}

function Test-GraphConnected {
    param([string[]]$RequiredScopes = @())

    try {
        if (-not (Get-Command Get-MgContext -ErrorAction SilentlyContinue)) { return $false }
        $ctx = Get-MgContext -ErrorAction SilentlyContinue
        if (-not $ctx) { return $false }

        # AuthType 'AppOnly' (cert) no tiene Account — basta con ClientId + TenantId.
        # AuthType 'Delegated' (interactivo / device code) sí necesita Account.
        $isAppOnly = ($ctx.AuthType -eq 'AppOnly') -or ([string]$ctx.AuthType -match 'AppOnly')
        $hasIdentity = if ($isAppOnly) {
            [bool]($ctx.ClientId -and $ctx.TenantId)
        } else {
            [bool]$ctx.Account
        }
        if (-not $hasIdentity) { return $false }

        # Para app-only los scopes son AppRoles concedidos en el tenant; no se validan aquí.
        if ($isAppOnly) { return $true }

        if (-not $RequiredScopes -or $RequiredScopes.Count -eq 0) { return $true }

        $current = @($ctx.Scopes)
        $missing = $RequiredScopes | Where-Object { $_ -notin $current }
        return [bool](-not $missing -or $missing.Count -eq 0)
    }
    catch { return $false }
}

# --- CONEXIÓN POR CERTIFICADO ---

function Connect-ByCertificate {
    param(
        [switch]$IncludeMgGraph,
        [switch]$IncludeExchangeOnline,
        [string[]]$GraphScopes = @()
    )

    if (-not (Test-CertConfigExists)) {
        throw "Método 'cert' seleccionado pero no hay certificado configurado. Ejecuta el asistente desde el menú principal."
    }

    $certParams = Get-CertConfig

    if ($IncludeExchangeOnline) {
        Ensure-ToolkitModule -ModuleName 'ExchangeOnlineManagement'

        if (-not (Test-ExchangeOnlineConnected)) {
            Write-Log "Conectando a Exchange Online (cert)..." 'INFO'
            Connect-ExchangeOnline `
                -AppId                 $certParams.AppId `
                -CertificateThumbprint $certParams.CertThumbprint `
                -Organization          $certParams.Organization `
                -ShowBanner:$false `
                -ErrorAction Stop
            Write-Log "Exchange Online conectado." 'OK'
        }
        else {
            Write-Log "Exchange Online ya conectado." 'OK'
        }
    }

    if ($IncludeMgGraph) {
        Ensure-ToolkitModule -ModuleName 'Microsoft.Graph.Authentication'
        Ensure-ToolkitModule -ModuleName 'Microsoft.Graph.Users'
        Ensure-ToolkitModule -ModuleName 'Microsoft.Graph.Groups'

        if (-not (Test-GraphConnected)) {
            Write-Log "Conectando a Microsoft Graph (cert)..." 'INFO'
            Connect-MgGraph `
                -ClientId              $certParams.AppId `
                -CertificateThumbprint $certParams.CertThumbprint `
                -TenantId              $certParams.TenantId `
                -NoWelcome `
                -ErrorAction Stop
            Write-Log "Microsoft Graph conectado." 'OK'
        }
        else {
            Write-Log "Microsoft Graph ya conectado." 'OK'
        }
    }
}

# --- CONEXIÓN TRADICIONAL ---

function Connect-Traditional {
    param(
        [switch]$IncludeMgGraph,
        [switch]$IncludeExchangeOnline,
        [string[]]$GraphScopes = @('User.Read.All','Group.Read.All')
    )

    $prefs = Get-UserPreferences

    if ($IncludeExchangeOnline) {
        Ensure-ToolkitModule -ModuleName 'ExchangeOnlineManagement'

        if (-not (Test-ExchangeOnlineConnected)) {
            $upn = $prefs.TraditionalAdminUpn
            if ([string]::IsNullOrWhiteSpace($upn)) {
                $upn = Read-Host 'UPN de administrador para Exchange Online'
                if ([string]::IsNullOrWhiteSpace($upn)) { throw "UPN requerido para conexión tradicional." }
                Set-PreferenceValue -Key 'TraditionalAdminUpn' -Value $upn
            }

            Write-Log "Conectando a Exchange Online como '$upn' (device code, abrirá login en navegador)..." 'INFO'
            Connect-ExchangeOnline -UserPrincipalName $upn -Device -ShowBanner:$false -ErrorAction Stop
            Write-Log "Exchange Online conectado." 'OK'
        }
        else {
            Write-Log "Exchange Online ya conectado." 'OK'
        }
    }

    if ($IncludeMgGraph) {
        Ensure-ToolkitModule -ModuleName 'Microsoft.Graph.Authentication'
        Ensure-ToolkitModule -ModuleName 'Microsoft.Graph.Users'
        Ensure-ToolkitModule -ModuleName 'Microsoft.Graph.Groups'

        $scopes = if ($GraphScopes -and $GraphScopes.Count -gt 0) { $GraphScopes } else { @('User.Read.All','Group.Read.All') }

        if (-not (Test-GraphConnected -RequiredScopes $scopes)) {
            Write-Log ("Conectando a Microsoft Graph con scopes [{0}] (login interactivo en navegador)..." -f ($scopes -join ', ')) 'INFO'
            Write-Host ''
            Write-Host '  >> Se abrirá una ventana del navegador para iniciar sesión en Microsoft Graph.' -ForegroundColor Yellow
            Write-Host '     Tras el login, la sesión queda activa hasta cerrar esta consola.' -ForegroundColor DarkGray
            Write-Host ''
            Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop
            Write-Log "Microsoft Graph conectado." 'OK'
        }
        else {
            Write-Log "Microsoft Graph ya conectado con scopes suficientes." 'OK'
        }
    }
}

# --- DISPATCHER PRINCIPAL ---

function Connect-RequiredServices {
    param(
        [switch]$MgGraph,
        [switch]$ExchangeOnline,
        [string[]]$GraphScopes = @('User.Read.All','Group.Read.All'),
        [ValidateSet('cert','traditional','auto')]
        [string]$Method = 'auto'
    )

    if (-not $MgGraph -and -not $ExchangeOnline) {
        Write-Log "Connect-RequiredServices invocado sin servicios. Nada que conectar." 'WARN'
        return
    }

    $prefs = Get-UserPreferences
    $effectiveMethod = $Method
    if ($effectiveMethod -eq 'auto') {
        $effectiveMethod = if ($prefs.ConnectionMethod) { $prefs.ConnectionMethod } else { 'cert' }
    }

    $services = @()
    if ($MgGraph)         { $services += 'Microsoft Graph' }
    if ($ExchangeOnline)  { $services += 'Exchange Online' }
    Write-Log ("Servicios a conectar: {0} | Método: {1}" -f ($services -join ', '), $effectiveMethod) 'INFO'

    switch ($effectiveMethod) {
        'cert' {
            if (-not (Test-CertConfigExists)) {
                throw "El método activo es 'certificado' pero no hay configuración válida. Lanza el asistente de creación de certificado desde el menú principal."
            }
            Connect-ByCertificate -IncludeMgGraph:$MgGraph -IncludeExchangeOnline:$ExchangeOnline -GraphScopes $GraphScopes
        }
        'traditional' {
            Connect-Traditional -IncludeMgGraph:$MgGraph -IncludeExchangeOnline:$ExchangeOnline -GraphScopes $GraphScopes
        }
        default { throw "Método desconocido: $effectiveMethod" }
    }
}

function Disconnect-AllServices {
    try {
        if (Test-ExchangeOnlineConnected) {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            Write-Log "Exchange Online desconectado." 'INFO'
        }
    } catch {}

    try {
        if (Test-GraphConnected) {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            Write-Log "Microsoft Graph desconectado." 'INFO'
        }
    } catch {}
}
