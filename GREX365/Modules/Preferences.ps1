# --- PREFERENCIAS DE USUARIO ---
# Lectura/escritura de GREX365/config/user_preferences.json
# No mantiene estado en variables: cada llamada lee el JSON desde disco.

function Get-PreferencesPath {
    if (-not $script:BasePath) {
        throw "BasePath no inicializado. Llama desde Main.ps1."
    }
    $configDir = Join-Path $script:BasePath 'config'
    if (-not (Test-Path -LiteralPath $configDir)) {
        New-Item -ItemType Directory -Path $configDir -Force | Out-Null
    }
    return (Join-Path $configDir 'user_preferences.json')
}

function Get-CertParamsPath {
    if (-not $script:BasePath) {
        throw "BasePath no inicializado."
    }
    $configDir = Join-Path $script:BasePath 'config'
    if (-not (Test-Path -LiteralPath $configDir)) {
        New-Item -ItemType Directory -Path $configDir -Force | Out-Null
    }
    return (Join-Path $configDir 'exo-app-params.json')
}

function New-DefaultPreferences {
    return [PSCustomObject]@{
        ConnectionMethod = $null
        TraditionalAdminUpn = $null
        Organization = $null
        FirstRunCompleted = $false
        LastUpdated = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss')
    }
}

function Get-UserPreferences {
    $path = Get-PreferencesPath
    if (-not (Test-Path -LiteralPath $path)) {
        return New-DefaultPreferences
    }

    try {
        $raw = Get-Content -LiteralPath $path -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) {
            return New-DefaultPreferences
        }
        $loaded = $raw | ConvertFrom-Json -ErrorAction Stop

        $defaults = New-DefaultPreferences
        foreach ($prop in $defaults.PSObject.Properties) {
            if (-not ($loaded.PSObject.Properties.Name -contains $prop.Name)) {
                $loaded | Add-Member -NotePropertyName $prop.Name -NotePropertyValue $prop.Value -Force
            }
        }
        return $loaded
    }
    catch {
        Write-Log "Preferencias corruptas, regenerando: $($_.Exception.Message)" 'WARN'
        return New-DefaultPreferences
    }
}

function Save-UserPreferences {
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Preferences
    )

    $Preferences.LastUpdated = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss')
    $path = Get-PreferencesPath
    ($Preferences | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $path -Encoding UTF8
}

function Set-PreferenceValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Key,

        [object]$Value
    )

    $prefs = Get-UserPreferences
    if ($prefs.PSObject.Properties.Name -contains $Key) {
        $prefs.$Key = $Value
    }
    else {
        $prefs | Add-Member -NotePropertyName $Key -NotePropertyValue $Value -Force
    }
    Save-UserPreferences -Preferences $prefs
}

function Test-CertConfigExists {
    $path = Get-CertParamsPath
    if (-not (Test-Path -LiteralPath $path)) { return $false }

    try {
        $params = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
        if (-not $params.AppId)          { return $false }
        if (-not $params.CertThumbprint) { return $false }
        if (-not $params.TenantId)       { return $false }
        if (-not $params.Organization)   { return $false }

        $cert = Get-Item -LiteralPath ("Cert:\CurrentUser\My\{0}" -f $params.CertThumbprint) -ErrorAction SilentlyContinue
        if (-not $cert) { return $false }
        if ($cert.NotAfter -lt (Get-Date)) { return $false }

        return $true
    }
    catch {
        return $false
    }
}

function Get-CertConfig {
    $path = Get-CertParamsPath
    if (-not (Test-Path -LiteralPath $path)) {
        throw "No existe configuración de certificado en: $path"
    }
    return (Get-Content -LiteralPath $path -Raw | ConvertFrom-Json)
}

function Remove-CertConfig {
    $configPath = Get-CertParamsPath
    if (-not (Test-Path -LiteralPath $configPath)) {
        Write-Log 'No hay configuración de certificado que eliminar.' 'WARN'
        return $false
    }

    $cfg = $null
    try { $cfg = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json } catch {
        Write-Log "Configuración corrupta, se eliminará igualmente: $($_.Exception.Message)" 'WARN'
    }

    $thumbprint = ''
    if ($cfg -and $cfg.PSObject.Properties.Name -contains 'CertThumbprint') {
        $thumbprint = [string]$cfg.CertThumbprint
    }

    $allOk = $true

    if ($thumbprint) {
        $certStorePath = "Cert:\CurrentUser\My\$thumbprint"
        if (Test-Path -LiteralPath $certStorePath) {
            try {
                Remove-Item -LiteralPath $certStorePath -DeleteKey -Force -ErrorAction Stop
                Write-Log "Certificado eliminado del almacén CurrentUser\My (thumbprint=$thumbprint)." 'OK'
            }
            catch {
                Write-Log "No se pudo eliminar cert del almacén: $($_.Exception.Message)" 'ERROR'
                $allOk = $false
            }
        }
        else {
            Write-Log "El cert ya no estaba en el almacén ($thumbprint)." 'WARN'
        }
    }
    else {
        Write-Log 'No se encontró thumbprint en el JSON. Solo se eliminará el archivo de configuración.' 'WARN'
    }

    if ($cfg -and $cfg.PSObject.Properties.Name -contains 'CerPath' -and $cfg.CerPath) {
        $cerFile = [string]$cfg.CerPath
        if (Test-Path -LiteralPath $cerFile) {
            try {
                Remove-Item -LiteralPath $cerFile -Force -ErrorAction Stop
                Write-Log "Archivo .cer eliminado: $cerFile" 'OK'
            }
            catch {
                Write-Log "No se pudo eliminar .cer: $($_.Exception.Message)" 'WARN'
            }
        }
    }

    try {
        Remove-Item -LiteralPath $configPath -Force -ErrorAction Stop
        Write-Log "JSON de configuración eliminado: $configPath" 'OK'
    }
    catch {
        Write-Log "No se pudo eliminar JSON: $($_.Exception.Message)" 'ERROR'
        $allOk = $false
    }

    if ($cfg -and $cfg.PSObject.Properties.Name -contains 'AppId' -and $cfg.AppId) {
        Write-Log ("La App Registration en Entra (AppId={0}) sigue existiendo. Elimínala manualmente desde el portal si quieres limpiar el tenant." -f $cfg.AppId) 'WARN'
    }

    return $allOk
}
