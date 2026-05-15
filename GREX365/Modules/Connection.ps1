# Connection module
# Authenticated session orchestration for Microsoft Graph and Exchange Online.
# Supports certificate (app-only) and device-code (delegated) flows.

function Get-RequiredToolkitModules {
    return @(
        'ExchangeOnlineManagement'
        'Microsoft.Graph.Authentication'
        'Microsoft.Graph.Users'
        'Microsoft.Graph.Groups'
        'Microsoft.Graph.Applications'
        'Microsoft.Graph.Identity.DirectoryManagement'
        'Microsoft.Graph.Identity.SignIns'
    )
}

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
        Write-Log 'Instalando proveedor NuGet...' -Source 'Connection'
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force -Confirm:$false -ErrorAction Stop | Out-Null
    }
}

function Ensure-PSGalleryTrusted {
    try {
        $repo = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
        if ($repo -and $repo.InstallationPolicy -ne 'Trusted') {
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
        }
    } catch {}
}

function Ensure-ToolkitModule {
    param([Parameter(Mandatory = $true)][string]$ModuleName)

    if (Test-ToolkitModuleImported -ModuleName $ModuleName) { return }

    if (-not (Test-ToolkitModuleInstalled -ModuleName $ModuleName)) {
        Ensure-NuGetProvider
        Ensure-PSGalleryTrusted
        Write-Log ("Instalando módulo '{0}'..." -f $ModuleName) -Source 'Connection'
        Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber -Confirm:$false -ErrorAction Stop
    }

    Import-Module $ModuleName -ErrorAction Stop -Verbose:$false -WarningAction SilentlyContinue 3>$null 4>$null
}

function Get-ToolkitModuleStatus {
    $list = Get-RequiredToolkitModules
    $result = New-Object System.Collections.Generic.List[object]
    foreach ($name in $list) {
        $available = Get-Module -ListAvailable -Name $name -ErrorAction SilentlyContinue |
                     Sort-Object Version -Descending | Select-Object -First 1
        $result.Add([PSCustomObject]@{
            Name      = $name
            Installed = [bool]$available
            Version   = if ($available) { [string]$available.Version } else { '' }
            Path      = if ($available) { [string]$available.ModuleBase } else { '' }
        })
    }
    return $result
}

function Get-ToolkitConfigFiles {
    $result = New-Object System.Collections.Generic.List[object]
    $prefsPath = $null; $certPath = $null
    try { $prefsPath = Get-PreferencesPath } catch {}
    try { $certPath  = Get-CertParamsPath } catch {}

    foreach ($entry in @(
        @{ Name = 'user_preferences.json'; Path = $prefsPath; Description = 'Método de conexión + UPN admin' }
        @{ Name = 'exo-app-params.json';   Path = $certPath;  Description = 'Parámetros App Registration + cert' }
    )) {
        $exists = $false
        if ($entry.Path) { $exists = Test-Path -LiteralPath $entry.Path }
        $result.Add([PSCustomObject]@{
            Name        = $entry.Name
            Path        = [string]$entry.Path
            Exists      = $exists
            Description = $entry.Description
        })
    }
    return $result
}

function Get-ToolkitConnectionState {
    $tenant = $null; $account = $null; $exoOrg = $null; $domain = $null
    $graphConnected = $false; $exoConnected = $false

    try {
        if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
            $ctx = Get-MgContext -ErrorAction SilentlyContinue
            if ($ctx) {
                if ($ctx.TenantId) { $tenant = [string]$ctx.TenantId }
                $isAppOnly = ([string]$ctx.AuthType -match 'AppOnly')
                if ($isAppOnly -and $ctx.ClientId -and $ctx.TenantId) {
                    $account = "App-only ($($ctx.ClientId))"
                    $graphConnected = $true
                }
                elseif ($ctx.Account) {
                    $account = [string]$ctx.Account
                    $graphConnected = $true
                }
            }
        }
    } catch {}

    try {
        if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
            $info = Get-ConnectionInformation -ErrorAction SilentlyContinue
            if ($info | Where-Object { $_.State -eq 'Connected' }) { $exoConnected = $true }
        }
    } catch {}

    if ($exoConnected) {
        try {
            if (Get-Command Get-OrganizationConfig -ErrorAction SilentlyContinue) {
                $org = Get-OrganizationConfig -ErrorAction SilentlyContinue
                if ($org -and $org.DisplayName) { $exoOrg = [string]$org.DisplayName }
            }
        } catch {}
        try {
            if (Get-Command Get-AcceptedDomain -ErrorAction SilentlyContinue) {
                $dom = Get-AcceptedDomain -ErrorAction SilentlyContinue | Where-Object { $_.Default } | Select-Object -First 1
                if ($dom) { $domain = [string]$dom.DomainName }
            }
        } catch {}
    }

    return [PSCustomObject]@{
        TenantId        = $tenant
        Account         = $account
        ExchangeOrgName = $exoOrg
        DefaultDomain   = $domain
        ExoConnected    = $exoConnected
        GraphConnected  = $graphConnected
    }
}

# --- SESSION CHECKS ---

function Test-ExchangeOnlineConnected {
    try {
        if (-not (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue)) { return $false }
        $sessions = Get-ConnectionInformation -ErrorAction Stop
        if (-not $sessions) { return $false }
        $active = $sessions | Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1
        return [bool]$active
    } catch { return $false }
}

function Test-GraphConnected {
    param([string[]]$RequiredScopes = @())

    try {
        if (-not (Get-Command Get-MgContext -ErrorAction SilentlyContinue)) { return $false }
        $ctx = Get-MgContext -ErrorAction SilentlyContinue
        if (-not $ctx) { return $false }

        $isAppOnly = ($ctx.AuthType -eq 'AppOnly') -or ([string]$ctx.AuthType -match 'AppOnly')
        $hasIdentity = if ($isAppOnly) { [bool]($ctx.ClientId -and $ctx.TenantId) } else { [bool]$ctx.Account }
        if (-not $hasIdentity) { return $false }
        if ($isAppOnly) { return $true }
        if (-not $RequiredScopes -or $RequiredScopes.Count -eq 0) { return $true }

        $current = @($ctx.Scopes)
        $missing = $RequiredScopes | Where-Object { $_ -notin $current }
        return [bool](-not $missing -or $missing.Count -eq 0)
    } catch { return $false }
}

function Assert-RequiredServicesReady {
    $exoOk = Test-ExchangeOnlineConnected
    $mgOk  = Test-GraphConnected
    if (-not $exoOk -or -not $mgOk) {
        throw ("Faltan servicios M365 conectados (EXO={0}, Graph={1}). Ejecuta este script desde Main.ps1." -f $exoOk, $mgOk)
    }
}

# --- TENANT LOCK ---
# When EnforceTenantLock is enabled in preferences, every successful connection is
# validated against the expected TenantId. Mismatch -> disconnect + throw.
# Designed to prevent accidental operation against the wrong tenant when admins
# manage several environments (dev / prod / customer A / customer B).

function Get-CurrentConnectedTenantId {
    try {
        if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
            $ctx = Get-MgContext -ErrorAction SilentlyContinue
            if ($ctx -and $ctx.TenantId) { return [string]$ctx.TenantId }
        }
    } catch {}
    try {
        if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
            $info = Get-ConnectionInformation -ErrorAction SilentlyContinue |
                    Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1
            if ($info -and $info.TenantId) { return [string]$info.TenantId }
        }
    } catch {}
    return $null
}

function Assert-TenantLock {
    param([string]$Source = 'Connection')

    $prefs = Get-UserPreferences
    $enforce = $false
    if ($prefs.PSObject.Properties.Name -contains 'EnforceTenantLock') {
        $enforce = [bool]$prefs.EnforceTenantLock
    }
    if (-not $enforce) { return }

    $expected = $null
    if ($prefs.PSObject.Properties.Name -contains 'ExpectedTenantId') {
        $expected = [string]$prefs.ExpectedTenantId
    }
    if ([string]::IsNullOrWhiteSpace($expected)) { return }

    $actual = Get-CurrentConnectedTenantId
    if (-not $actual) { return }

    if ($actual -ne $expected) {
        Write-Log ("TENANT LOCK violado: esperado={0} actual={1}. Desconectando." -f $expected, $actual) -Level ERROR -Source $Source
        try { Disconnect-AllServices } catch {}
        throw ("Tenant lock violado. Esperado={0} actual={1}. Conexión abortada." -f $expected, $actual)
    }
}

# --- CERTIFICATE FLOW ---

function Connect-ByCertificate {
    param(
        [switch]$IncludeMgGraph,
        [switch]$IncludeExchangeOnline,
        [string[]]$GraphScopes = @()
    )

    Write-Log 'cert: validando configuración local...' -Source 'Connection'
    if (-not (Test-CertConfigExists)) {
        throw "Método 'cert' seleccionado pero no hay certificado configurado. Ejecuta el asistente desde el menú principal."
    }

    $certParams = Get-CertConfig
    Write-Log ("cert: config OK · AppId={0} · Org={1} · TenantId={2} · Thumbprint={3}" -f $certParams.AppId, $certParams.Organization, $certParams.TenantId, $certParams.CertThumbprint) -Source 'Connection'

    if ($IncludeExchangeOnline) {
        Write-Log 'cert/EXO: asegurando módulo ExchangeOnlineManagement...' -Source 'Connection'
        Ensure-ToolkitModule -ModuleName 'ExchangeOnlineManagement'
        Write-Log 'cert/EXO: módulo listo.' -Source 'Connection'
        if (-not (Test-ExchangeOnlineConnected)) {
            Write-Log 'cert/EXO: llamando Connect-ExchangeOnline (puede tardar 10-30s)...' -Source 'Connection'
            Connect-ExchangeOnline `
                -AppId                 $certParams.AppId `
                -CertificateThumbprint $certParams.CertThumbprint `
                -Organization          $certParams.Organization `
                -ShowBanner:$false `
                -ErrorAction Stop
            Write-Log 'cert/EXO: Connect-ExchangeOnline retornó.' -Source 'Connection'
            if (Get-Command Reset-SessionStateCache -ErrorAction SilentlyContinue) { Reset-SessionStateCache }
            Assert-TenantLock -Source 'ConnectByCertificate-EXO'
            Write-Log 'Exchange Online conectado.' -Level OK -Source 'Connection'
        } else {
            Write-Log 'Exchange Online ya conectado.' -Level OK -Source 'Connection'
        }
    }

    if ($IncludeMgGraph) {
        Write-Log 'cert/Graph: asegurando módulos Microsoft.Graph.* ...' -Source 'Connection'
        Ensure-ToolkitModule -ModuleName 'Microsoft.Graph.Authentication'
        Ensure-ToolkitModule -ModuleName 'Microsoft.Graph.Users'
        Ensure-ToolkitModule -ModuleName 'Microsoft.Graph.Groups'
        Write-Log 'cert/Graph: módulos listos.' -Source 'Connection'

        if (-not (Test-GraphConnected)) {
            Write-Log 'cert/Graph: llamando Connect-MgGraph (puede tardar 5-20s)...' -Source 'Connection'
            Connect-MgGraph `
                -ClientId              $certParams.AppId `
                -CertificateThumbprint $certParams.CertThumbprint `
                -TenantId              $certParams.TenantId `
                -NoWelcome `
                -ErrorAction Stop
            Write-Log 'cert/Graph: Connect-MgGraph retornó.' -Source 'Connection'
            if (Get-Command Reset-SessionStateCache -ErrorAction SilentlyContinue) { Reset-SessionStateCache }
            Assert-TenantLock -Source 'ConnectByCertificate-Graph'
            Write-Log 'Microsoft Graph conectado.' -Level OK -Source 'Connection'
        } else {
            Write-Log 'Microsoft Graph ya conectado.' -Level OK -Source 'Connection'
        }
    }
}

# --- TRADITIONAL FLOW ---

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
                $upn = Read-Input -Prompt 'UPN de administrador para Exchange Online'
                if ([string]::IsNullOrWhiteSpace($upn)) { throw 'UPN requerido para conexión tradicional.' }
                Set-PreferenceValue -Key 'TraditionalAdminUpn' -Value $upn
            }
            Write-Log ("Conectando a Exchange Online como '{0}' (device code)..." -f $upn) -Source 'Connection'
            Connect-ExchangeOnline -UserPrincipalName $upn -Device -ShowBanner:$false -ErrorAction Stop
            if (Get-Command Reset-SessionStateCache -ErrorAction SilentlyContinue) { Reset-SessionStateCache }
            Assert-TenantLock -Source 'ConnectTraditional-EXO'
            Write-Log 'Exchange Online conectado.' -Level OK -Source 'Connection'
        } else {
            Write-Log 'Exchange Online ya conectado.' -Level OK -Source 'Connection'
        }
    }

    if ($IncludeMgGraph) {
        Ensure-ToolkitModule -ModuleName 'Microsoft.Graph.Authentication'
        Ensure-ToolkitModule -ModuleName 'Microsoft.Graph.Users'
        Ensure-ToolkitModule -ModuleName 'Microsoft.Graph.Groups'

        $scopes = if ($GraphScopes -and $GraphScopes.Count -gt 0) { $GraphScopes } else { @('User.Read.All','Group.Read.All') }

        if (-not (Test-GraphConnected -RequiredScopes $scopes)) {
            Write-Log ('Conectando a Microsoft Graph con scopes: ' + ($scopes -join ', ')) -Source 'Connection'
            Show-WarningBlock -Title 'Login interactivo requerido' -Detail 'Se abrirá una ventana del navegador para autenticar.'
            Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop
            if (Get-Command Reset-SessionStateCache -ErrorAction SilentlyContinue) { Reset-SessionStateCache }
            Assert-TenantLock -Source 'ConnectTraditional-Graph'
            Write-Log 'Microsoft Graph conectado.' -Level OK -Source 'Connection'
        } else {
            Write-Log 'Microsoft Graph ya conectado.' -Level OK -Source 'Connection'
        }
    }
}

# --- DISPATCHER ---

function Connect-RequiredServices {
    param(
        [switch]$MgGraph,
        [switch]$ExchangeOnline,
        [string[]]$GraphScopes = @('User.Read.All','Group.Read.All'),
        [ValidateSet('cert','traditional','auto')]
        [string]$Method = 'auto'
    )

    if (-not $MgGraph -and -not $ExchangeOnline) {
        Write-Log 'Connect-RequiredServices invocado sin servicios.' -Level WARN -Source 'Connection'
        return
    }

    $prefs = Get-UserPreferences
    $effective = $Method
    if ($effective -eq 'auto') {
        $effective = if ($prefs.ConnectionMethod) { $prefs.ConnectionMethod } else { 'cert' }
    }

    $services = @()
    if ($MgGraph)         { $services += 'Graph' }
    if ($ExchangeOnline)  { $services += 'Exchange' }
    Write-Log ("Conectando: {0} | Método: {1}" -f ($services -join ', '), $effective) -Source 'Connection'

    switch ($effective) {
        'cert' {
            if (-not (Test-CertConfigExists)) {
                throw "El método activo es 'certificado' pero no hay configuración válida. Lanza el asistente desde el menú principal."
            }
            Connect-ByCertificate -IncludeMgGraph:$MgGraph -IncludeExchangeOnline:$ExchangeOnline -GraphScopes $GraphScopes
        }
        'traditional' {
            Connect-Traditional -IncludeMgGraph:$MgGraph -IncludeExchangeOnline:$ExchangeOnline -GraphScopes $GraphScopes
        }
        default { throw "Método desconocido: $effective" }
    }
}

function Disconnect-AllServices {
    try {
        if (Test-ExchangeOnlineConnected) {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            Write-Log 'Exchange Online desconectado.' -Source 'Connection'
        }
    } catch {}
    try {
        if (Test-GraphConnected) {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            Write-Log 'Microsoft Graph desconectado.' -Source 'Connection'
        }
    } catch {}
    if (Get-Command Reset-SessionStateCache -ErrorAction SilentlyContinue) { Reset-SessionStateCache }
}
