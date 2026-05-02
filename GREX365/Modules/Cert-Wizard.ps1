# --- ASISTENTE AUTOMATIZADO DE CERTIFICADO ExO + Graph ---
# Implementa los 29 pasos del CSV cert_instrunciones/EXO_Cert_Auth_Pasos.csv
# de forma desatendida. Solo pide al usuario:
#   - TenantId
#   - Nombre de la App Registration
#   - Validez del certificado (años)
#   - Dominio de la organización (para Connect-ExchangeOnline)

# IDs constantes (públicos, fijos en Microsoft)
$script:CW_ExoAppId   = '00000002-0000-0ff1-ce00-000000000000'
$script:CW_GraphAppId = '00000003-0000-0000-c000-000000000000'

# Roles directorio (templateIds públicos)
$script:CW_RoleTemplate_ExchangeAdmin = '29232cdf-9323-42fd-ade2-1d097af3e4de'
$script:CW_RoleTemplate_UserAdmin     = 'fe930be7-5e62-47db-91af-98c3a49a38b1'
$script:CW_RoleTemplate_GroupsAdmin   = 'fdd7a751-b60b-444a-984c-02652fe8fa1c'

# Permisos Graph a otorgar (8 AppRoles)
$script:CW_GraphAppRoles = @(
    'User.ReadWrite.All'
    'Group.ReadWrite.All'
    'GroupMember.ReadWrite.All'
    'Directory.ReadWrite.All'
    'Organization.Read.All'
    'RoleManagement.Read.Directory'
    'UserAuthenticationMethod.ReadWrite.All'
    'Policy.Read.All'
)

# --- HELPERS ---

function Show-WizardStep {
    param(
        [int]$Number,
        [string]$Title,
        [string]$Description = ''
    )
    Write-Host ''
    Write-Host ("  [{0}/29] {1}" -f $Number, $Title) -ForegroundColor Cyan
    if ($Description) {
        Write-Host ("        {0}" -f $Description) -ForegroundColor DarkGray
    }
}

function Wait-WithProgress {
    param(
        [int]$Seconds = 60,
        [string]$Activity = 'Esperando propagación de permisos...'
    )

    for ($i = 1; $i -le $Seconds; $i++) {
        $pct = [int](($i / $Seconds) * 100)
        Write-Progress -Activity $Activity -Status ("$i / $Seconds segundos") -PercentComplete $pct
        Start-Sleep -Seconds 1
    }
    Write-Progress -Activity $Activity -Completed
}

function Read-NonEmptyInput {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [string]$Default = ''
    )

    while ($true) {
        $msg = if ($Default) { "$Prompt [$Default]" } else { $Prompt }
        $value = Read-Host $msg
        if ([string]::IsNullOrWhiteSpace($value) -and $Default) { return $Default }
        if (-not [string]::IsNullOrWhiteSpace($value)) { return $value.Trim() }
        Write-Host '  Valor requerido.' -ForegroundColor Yellow
    }
}

function Show-CertWizardFieldHint {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('TenantId','AppName','Years','Organization')]
        [string]$Field
    )

    switch ($Field) {
        'TenantId' {
            Write-Host '    └─ Qué es      : GUID único del tenant Entra ID (Azure AD).' -ForegroundColor DarkGray
            Write-Host '       Para qué    : dirige autenticación y consent al tenant correcto.' -ForegroundColor DarkGray
            Write-Host '       Dónde mirar : Entra Portal → Microsoft Entra ID → Overview → Tenant ID' -ForegroundColor DarkGray
            Write-Host '       URL directa : https://entra.microsoft.com/#view/Microsoft_AAD_IAM/TenantOverview.ReactView' -ForegroundColor DarkCyan
        }
        'AppName' {
            Write-Host '    └─ Qué es      : nombre visible de la App Registration que se creará.' -ForegroundColor DarkGray
            Write-Host '       Para qué    : la verás con este nombre en Entra y en los logs.' -ForegroundColor DarkGray
            Write-Host '       Dónde mirar : Entra Portal → Applications → App registrations' -ForegroundColor DarkGray
            Write-Host '       URL directa : https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade' -ForegroundColor DarkCyan
        }
        'Years' {
            Write-Host '    └─ Qué es      : validez del certificado X.509 self-signed.' -ForegroundColor DarkGray
            Write-Host '       Rango       : 1-5 años. Renovación manual al expirar.' -ForegroundColor DarkGray
        }
        'Organization' {
            Write-Host '    └─ Qué es      : dominio principal del tenant (parámetro -Organization de Connect-ExchangeOnline).' -ForegroundColor DarkGray
            Write-Host '       Formato     : <tenant>.onmicrosoft.com (o tu dominio verificado).' -ForegroundColor DarkGray
            Write-Host '       Dónde mirar : Entra Portal → Settings → Domain names (marca el Default/Primary)' -ForegroundColor DarkGray
            Write-Host '       URL directa : https://entra.microsoft.com/#view/Microsoft_AAD_IAM/DomainsList.ReactView' -ForegroundColor DarkCyan
        }
    }
    Write-Host ''
}

function Test-GuidString {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    $g = [guid]::Empty
    return [guid]::TryParse($Value.Trim(), [ref]$g)
}

# --- PASO 1: módulos requeridos ---

function Step-EnsureModules {
    Show-WizardStep -Number 1 -Title 'Verificar módulos PowerShell' -Description 'ExchangeOnlineManagement + Microsoft.Graph (Authentication, Applications, Identity.DirectoryManagement, Identity.SignIns)'

    $required = @(
        'ExchangeOnlineManagement'
        'Microsoft.Graph.Authentication'
        'Microsoft.Graph.Applications'
        'Microsoft.Graph.Identity.DirectoryManagement'
        'Microsoft.Graph.Identity.SignIns'
    )

    foreach ($m in $required) {
        Ensure-ToolkitModule -ModuleName $m
    }
    Write-Log 'Módulos OK.' 'OK'
}

# --- PASOS 2-6: creación + persistencia + export del certificado ---

function Step-CreateCertificate {
    param(
        [Parameter(Mandatory = $true)][string]$AppName,
        [Parameter(Mandatory = $true)][int]$ValidityYears,
        [Parameter(Mandatory = $true)][string]$ExportFolder
    )

    Show-WizardStep -Number 2 -Title 'Generar clave RSA 2048 en memoria'
    $rsa = [System.Security.Cryptography.RSA]::Create(2048)
    Write-Log "Clave RSA 2048 creada." 'OK'

    Show-WizardStep -Number 3 -Title 'Crear CertificateRequest self-signed SHA256'
    $subject = "CN=$AppName"
    $req = New-Object System.Security.Cryptography.X509Certificates.CertificateRequest(
        $subject, $rsa,
        [System.Security.Cryptography.HashAlgorithmName]::SHA256,
        [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
    )

    $basicConstraints = New-Object System.Security.Cryptography.X509Certificates.X509BasicConstraintsExtension($false, $false, 0, $false)
    $req.CertificateExtensions.Add($basicConstraints)

    $keyUsage = New-Object System.Security.Cryptography.X509Certificates.X509KeyUsageExtension(
        [System.Security.Cryptography.X509Certificates.X509KeyUsageFlags]::DigitalSignature, $false
    )
    $req.CertificateExtensions.Add($keyUsage)
    Write-Log "Request preparado con BasicConstraints + KeyUsage." 'OK'

    Show-WizardStep -Number 4 -Title "Firmar self-signed (validez $ValidityYears años)"
    $notBefore = (Get-Date).ToUniversalTime().AddMinutes(-5)
    $notAfter  = (Get-Date).ToUniversalTime().AddYears($ValidityYears)
    $tempCert = $req.CreateSelfSigned([System.DateTimeOffset]$notBefore, [System.DateTimeOffset]$notAfter)
    Write-Log "Cert en memoria generado. Thumbprint=$($tempCert.Thumbprint)" 'OK'

    Show-WizardStep -Number 5 -Title 'Importar a CurrentUser\My (PersistKeySet + UserKeySet)'
    $pfxBytes = $tempCert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx, '')

    $flags = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::PersistKeySet -bor `
             [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::UserKeySet

    $persistedCert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($pfxBytes, '', $flags)

    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store('My', 'CurrentUser')
    $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    $store.Add($persistedCert)
    $store.Close()
    Write-Log "Cert persistido. Thumbprint=$($persistedCert.Thumbprint)" 'OK'

    Show-WizardStep -Number 6 -Title 'Exportar parte pública (.cer) sin clave privada'
    if (-not (Test-Path -LiteralPath $ExportFolder)) {
        New-Item -ItemType Directory -Path $ExportFolder -Force | Out-Null
    }
    $cerPath = Join-Path $ExportFolder ("{0}.cer" -f $AppName)
    [System.IO.File]::WriteAllBytes($cerPath, $persistedCert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))
    Write-Log "Cert público exportado: $cerPath" 'OK'

    try { $rsa.Dispose() } catch {}
    try { $tempCert.Dispose() } catch {}

    return [PSCustomObject]@{
        Thumbprint = $persistedCert.Thumbprint
        Subject    = $persistedCert.Subject
        NotAfter   = $persistedCert.NotAfter
        NotBefore  = $persistedCert.NotBefore
        CerPath    = $cerPath
        RawPublic  = $persistedCert.GetRawCertData()
    }
}

# --- PASO 7: ACL en clave privada (best-effort) ---

function Step-RestrictPrivateKeyAcl {
    param([Parameter(Mandatory = $true)][string]$Thumbprint)

    Show-WizardStep -Number 7 -Title 'Restringir ACL de la clave privada (usuario actual + SYSTEM)'

    try {
        $cert = Get-Item "Cert:\CurrentUser\My\$Thumbprint" -ErrorAction Stop
        $keyName = $null

        if ($cert.PrivateKey -and $cert.PrivateKey.CspKeyContainerInfo) {
            $keyName = $cert.PrivateKey.CspKeyContainerInfo.UniqueKeyContainerName
        }
        if (-not $keyName) {
            $rsaCng = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
            if ($rsaCng -and $rsaCng.Key) { $keyName = $rsaCng.Key.UniqueName }
        }

        if (-not $keyName) {
            Write-Log "No se pudo localizar el contenedor de clave privada. Se omite ACL hardening." 'WARN'
            return
        }

        $candidates = @(
            (Join-Path $env:APPDATA 'Microsoft\Crypto\Keys')
            (Join-Path $env:APPDATA 'Microsoft\Crypto\RSA')
            (Join-Path $env:ALLUSERSPROFILE 'Application Data\Microsoft\Crypto\Keys')
            (Join-Path $env:ALLUSERSPROFILE 'Microsoft\Crypto\Keys')
        )

        $keyFile = $null
        foreach ($folder in $candidates) {
            if (-not (Test-Path -LiteralPath $folder)) { continue }
            $candidate = Join-Path $folder $keyName
            if (Test-Path -LiteralPath $candidate) { $keyFile = $candidate; break }
        }

        if (-not $keyFile) {
            Write-Log "No se localizó el archivo físico de la clave. Se omite ACL hardening." 'WARN'
            return
        }

        $acl = Get-Acl -LiteralPath $keyFile
        $acl.SetAccessRuleProtection($true, $false)
        $rules = @($acl.Access)
        foreach ($r in $rules) { [void]$acl.RemoveAccessRule($r) }

        $userSid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
        $userRule = New-Object System.Security.AccessControl.FileSystemAccessRule($userSid, 'FullControl', 'Allow')
        [void]$acl.AddAccessRule($userRule)

        $systemSid = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-18')
        $systemRule = New-Object System.Security.AccessControl.FileSystemAccessRule($systemSid, 'FullControl', 'Allow')
        [void]$acl.AddAccessRule($systemRule)

        Set-Acl -LiteralPath $keyFile -AclObject $acl
        Write-Log "ACL aplicada en: $keyFile" 'OK'
    }
    catch {
        Write-Log "ACL hardening omitido: $($_.Exception.Message)" 'WARN'
    }
}

# --- PASO 8: conexión Graph (delegated, device code) ---

function Step-ConnectGraphAdmin {
    param([Parameter(Mandatory = $true)][string]$TenantId)

    Show-WizardStep -Number 8 -Title 'Conectar a Microsoft Graph (login interactivo en navegador)' -Description 'Necesitas un Global Admin del tenant'

    Write-Host ''
    Write-Host '  >> Se abrirá el navegador. Inicia sesión con la cuenta Global Admin del tenant.' -ForegroundColor Yellow
    Write-Host ''

    Connect-MgGraph `
        -TenantId $TenantId `
        -Scopes @(
            'Application.ReadWrite.All'
            'AppRoleAssignment.ReadWrite.All'
            'RoleManagement.ReadWrite.Directory'
            'Directory.ReadWrite.All'
        ) `
        -ContextScope Process `
        -NoWelcome `
        -ErrorAction Stop

    $ctx = Get-MgContext
    if (-not $ctx -or -not $ctx.Account) { throw "Login Graph fallido." }
    Write-Log ("Login Graph OK. Cuenta: {0} | Tenant: {1}" -f $ctx.Account, $ctx.TenantId) 'OK'
    return $ctx
}

# --- PASOS 9-15: App Reg, SP, permisos Exchange, rol Exchange Admin ---

function Step-CreateAppRegistration {
    param(
        [Parameter(Mandatory = $true)][string]$DisplayName,
        [Parameter(Mandatory = $true)][byte[]]$CertPublicBytes
    )

    Show-WizardStep -Number 9 -Title "Crear App Registration single-tenant: $DisplayName"
    $app = New-MgApplication -DisplayName $DisplayName -SignInAudience 'AzureADMyOrg' -ErrorAction Stop
    Write-Log ("App creada. AppId={0} | ObjectId={1}" -f $app.AppId, $app.Id) 'OK'

    Show-WizardStep -Number 10 -Title 'Adjuntar certificado público a la App'
    $keyCred = @{
        Type = 'AsymmetricX509Cert'
        Usage = 'Verify'
        Key = $CertPublicBytes
        DisplayName = "$DisplayName-cert"
    }
    Update-MgApplication -ApplicationId $app.Id -KeyCredentials @($keyCred) -ErrorAction Stop
    Write-Log "Certificado asociado a la App." 'OK'

    Show-WizardStep -Number 11 -Title 'Crear Service Principal'
    $sp = New-MgServicePrincipal -AppId $app.AppId -ErrorAction Stop
    Write-Log ("SP creado. ObjectId={0}" -f $sp.Id) 'OK'

    return [PSCustomObject]@{ App = $app; ServicePrincipal = $sp }
}

function Get-ExoManageAsAppRoleId {
    $exoSp = Get-MgServicePrincipal -Filter "appId eq '$($script:CW_ExoAppId)'" -ErrorAction Stop
    if (-not $exoSp) { throw "No se encontró el SP de Office 365 Exchange Online en el tenant." }

    $role = $exoSp.AppRoles | Where-Object { $_.Value -eq 'Exchange.ManageAsApp' } | Select-Object -First 1
    if (-not $role) { throw "No se encontró el AppRole 'Exchange.ManageAsApp'." }

    return [PSCustomObject]@{ Sp = $exoSp; RoleId = $role.Id }
}

function Step-AssignExchangePermission {
    param(
        [Parameter(Mandatory = $true)]$AppObject,
        [Parameter(Mandatory = $true)]$SpObject
    )

    Show-WizardStep -Number 12 -Title 'Declarar RequiredResourceAccess Exchange.ManageAsApp'
    $exoInfo = Get-ExoManageAsAppRoleId

    $rra = @{
        ResourceAppId = $script:CW_ExoAppId
        ResourceAccess = @(@{ Id = $exoInfo.RoleId; Type = 'Role' })
    }
    Update-MgApplication -ApplicationId $AppObject.Id -RequiredResourceAccess @($rra) -ErrorAction Stop
    Write-Log "Permiso Exchange.ManageAsApp declarado." 'OK'

    Show-WizardStep -Number 13 -Title 'Admin consent (AppRoleAssignment)'
    $body = @{
        PrincipalId = $SpObject.Id
        ResourceId  = $exoInfo.Sp.Id
        AppRoleId   = $exoInfo.RoleId
    }
    try {
        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $SpObject.Id -BodyParameter $body -ErrorAction Stop | Out-Null
        Write-Log "Consent Exchange.ManageAsApp aplicado." 'OK'
    }
    catch {
        if ($_.Exception.Message -match 'Permission_Conflict|already exists') {
            Write-Log "Consent Exchange ya estaba aplicado." 'WARN'
        }
        else { throw }
    }

    return $exoInfo
}

function Get-OrActivate-DirectoryRole {
    param([Parameter(Mandatory = $true)][string]$RoleTemplateId)

    $role = Get-MgDirectoryRole -Filter "roleTemplateId eq '$RoleTemplateId'" -ErrorAction SilentlyContinue
    if ($role) { return $role }

    Write-Log "Rol no activo. Activando templateId=$RoleTemplateId..." 'INFO'
    $role = New-MgDirectoryRole -RoleTemplateId $RoleTemplateId -ErrorAction Stop
    return $role
}

function Add-RoleMember {
    param(
        [Parameter(Mandatory = $true)][string]$RoleId,
        [Parameter(Mandatory = $true)][string]$PrincipalObjectId
    )

    $body = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$PrincipalObjectId" }
    try {
        New-MgDirectoryRoleMemberByRef -DirectoryRoleId $RoleId -BodyParameter $body -ErrorAction Stop
    }
    catch {
        if ($_.Exception.Message -match 'already exist|already a member|conflicting object') {
            Write-Log "  El SP ya era miembro de este rol." 'WARN'
        }
        else { throw }
    }
}

function Step-AssignExchangeAdminRole {
    param([Parameter(Mandatory = $true)]$SpObject)

    Show-WizardStep -Number 14 -Title 'Activar rol Exchange Administrator (si no existe)'
    $role = Get-OrActivate-DirectoryRole -RoleTemplateId $script:CW_RoleTemplate_ExchangeAdmin
    Write-Log ("Rol Exchange Administrator OK. Id={0}" -f $role.Id) 'OK'

    Show-WizardStep -Number 15 -Title 'Asignar SP al rol Exchange Administrator'
    Add-RoleMember -RoleId $role.Id -PrincipalObjectId $SpObject.Id
    Write-Log "SP asignado a Exchange Administrator." 'OK'
}

# --- PASO 16: persistir parámetros ---

function Step-SaveConfig {
    param(
        [Parameter(Mandatory = $true)]$AppObject,
        [Parameter(Mandatory = $true)]$SpObject,
        [Parameter(Mandatory = $true)]$CertInfo,
        [Parameter(Mandatory = $true)][string]$TenantId,
        [Parameter(Mandatory = $true)][string]$Organization,
        [Parameter(Mandatory = $true)][string]$ConfigPath
    )

    Show-WizardStep -Number 16 -Title 'Guardar parámetros (no secretos) en JSON'
    $payload = [PSCustomObject]@{
        TenantId         = $TenantId
        Organization     = $Organization
        AppId            = $AppObject.AppId
        AppObjectId      = $AppObject.Id
        SpObjectId       = $SpObject.Id
        CertThumbprint   = $CertInfo.Thumbprint
        CertSubject      = $CertInfo.Subject
        CertNotAfter     = $CertInfo.NotAfter.ToString('o')
        CerPath          = $CertInfo.CerPath
        CreatedAt        = (Get-Date).ToString('o')
    }

    $folder = Split-Path -Parent $ConfigPath
    if (-not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }

    ($payload | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $ConfigPath -Encoding UTF8
    Write-Log "Configuración guardada en: $ConfigPath" 'OK'
}

# --- PASOS 17-19: espera + test EXO ---

function Step-WaitAndTestExo {
    param(
        [Parameter(Mandatory = $true)]$AppObject,
        [Parameter(Mandatory = $true)]$CertInfo,
        [Parameter(Mandatory = $true)][string]$Organization
    )

    Show-WizardStep -Number 17 -Title 'Esperar 60s para propagación de consent + rol'
    Wait-WithProgress -Seconds 60

    Show-WizardStep -Number 18 -Title 'Conectar Exchange Online con certificado (test)'
    Connect-ExchangeOnline `
        -AppId                 $AppObject.AppId `
        -CertificateThumbprint $CertInfo.Thumbprint `
        -Organization          $Organization `
        -ShowBanner:$false `
        -ErrorAction Stop
    Write-Log "Connect-ExchangeOnline OK (cert)." 'OK'

    Show-WizardStep -Number 19 -Title 'Validar operación real (Get-OrganizationConfig)'
    $orgConf = Get-OrganizationConfig -ErrorAction Stop
    Write-Log ("Tenant Display: {0}" -f $orgConf.DisplayName) 'OK'
}

# --- PASOS 22-25: añadir permisos Graph ---

function Step-AddGraphPermissions {
    param(
        [Parameter(Mandatory = $true)]$AppObject,
        [Parameter(Mandatory = $true)]$SpObject,
        [Parameter(Mandatory = $true)][PSCustomObject]$ExoInfo
    )

    Show-WizardStep -Number 22 -Title 'Reusar contexto Graph (sigue activa la sesión delegated)'
    $ctx = Get-MgContext
    if (-not $ctx -or -not $ctx.Account) {
        Write-Log "Contexto Graph perdido, reconectando..." 'WARN'
        Connect-MgGraph -TenantId $ctx.TenantId -Scopes @(
            'Application.ReadWrite.All','AppRoleAssignment.ReadWrite.All',
            'RoleManagement.ReadWrite.Directory','Directory.ReadWrite.All'
        ) -ContextScope Process -NoWelcome -ErrorAction Stop
    }

    Show-WizardStep -Number 23 -Title 'Resolver IDs de los 8 AppRoles de Microsoft Graph'
    $graphSp = Get-MgServicePrincipal -Filter "appId eq '$($script:CW_GraphAppId)'" -ErrorAction Stop
    if (-not $graphSp) { throw "No se encontró el SP de Microsoft Graph." }

    $resolvedRoles = New-Object System.Collections.Generic.List[object]
    foreach ($value in $script:CW_GraphAppRoles) {
        $r = $graphSp.AppRoles | Where-Object { $_.Value -eq $value -and $_.AllowedMemberTypes -contains 'Application' } | Select-Object -First 1
        if (-not $r) {
            Write-Log "  AppRole no encontrado: $value" 'ERROR'
            continue
        }
        $resolvedRoles.Add([PSCustomObject]@{ Value = $value; Id = $r.Id })
        Write-Log "  $value → $($r.Id)" 'OK'
    }
    if ($resolvedRoles.Count -eq 0) { throw "No se pudo resolver ningún AppRole de Graph." }

    Show-WizardStep -Number 24 -Title 'Actualizar manifest preservando Exchange + añadiendo Graph'
    $rraExchange = @{
        ResourceAppId  = $script:CW_ExoAppId
        ResourceAccess = @(@{ Id = $ExoInfo.RoleId; Type = 'Role' })
    }
    $rraGraph = @{
        ResourceAppId  = $script:CW_GraphAppId
        ResourceAccess = @($resolvedRoles | ForEach-Object { @{ Id = $_.Id; Type = 'Role' } })
    }
    Update-MgApplication -ApplicationId $AppObject.Id -RequiredResourceAccess @($rraExchange, $rraGraph) -ErrorAction Stop
    Write-Log "Manifest actualizado con Exchange + Graph." 'OK'

    Show-WizardStep -Number 25 -Title 'Admin consent para cada AppRole de Graph'
    foreach ($r in $resolvedRoles) {
        $body = @{ PrincipalId = $SpObject.Id; ResourceId = $graphSp.Id; AppRoleId = $r.Id }
        try {
            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $SpObject.Id -BodyParameter $body -ErrorAction Stop | Out-Null
            Write-Log "  Consent OK: $($r.Value)" 'OK'
        }
        catch {
            if ($_.Exception.Message -match 'Permission_Conflict|already exists') {
                Write-Log "  Consent ya aplicado: $($r.Value)" 'WARN'
            }
            else {
                Write-Log "  Consent fallido para $($r.Value): $($_.Exception.Message)" 'ERROR'
            }
        }
    }
}

# --- PASOS 26-27: roles directorio User Admin + Groups Admin ---

function Step-AssignDirectoryRoles {
    param([Parameter(Mandatory = $true)]$SpObject)

    Show-WizardStep -Number 26 -Title 'Añadir rol User Administrator al SP'
    $userAdmin = Get-OrActivate-DirectoryRole -RoleTemplateId $script:CW_RoleTemplate_UserAdmin
    Add-RoleMember -RoleId $userAdmin.Id -PrincipalObjectId $SpObject.Id
    Write-Log "User Administrator asignado." 'OK'

    Show-WizardStep -Number 27 -Title 'Añadir rol Groups Administrator al SP'
    $groupsAdmin = Get-OrActivate-DirectoryRole -RoleTemplateId $script:CW_RoleTemplate_GroupsAdmin
    Add-RoleMember -RoleId $groupsAdmin.Id -PrincipalObjectId $SpObject.Id
    Write-Log "Groups Administrator asignado." 'OK'
}

# --- PASO 29: test final dual ---

function Step-FinalTest {
    Show-WizardStep -Number 29 -Title 'Test dual: Graph (Get-MgUser) + EXO (Get-DistributionGroup)'

    try {
        $u = Get-MgUser -Top 3 -ErrorAction Stop
        Write-Log ("Get-MgUser OK ({0} usuarios listados)." -f @($u).Count) 'OK'
    }
    catch {
        Write-Log "Get-MgUser falló: $($_.Exception.Message)" 'WARN'
    }

    try {
        $g = Get-MgGroup -Top 3 -ErrorAction Stop
        Write-Log ("Get-MgGroup OK ({0} grupos listados)." -f @($g).Count) 'OK'
    }
    catch {
        Write-Log "Get-MgGroup falló: $($_.Exception.Message)" 'WARN'
    }

    try {
        $dl = Get-DistributionGroup -ResultSize 3 -ErrorAction Stop
        Write-Log ("Get-DistributionGroup OK ({0} listados)." -f @($dl).Count) 'OK'
    }
    catch {
        Write-Log "Get-DistributionGroup falló: $($_.Exception.Message)" 'WARN'
    }
}

# --- ENTRYPOINT ---

function Start-CertificateWizard {
    param(
        [Parameter(Mandatory = $true)][string]$CsvStepsPath,
        [Parameter(Mandatory = $true)][string]$ConfigPath
    )

    Show-Header -Title 'ASISTENTE DE CERTIFICADO' -Subtitle 'Automatización de los 29 pasos (ExO + Graph)'

    Write-Centered -Text 'Este asistente creará un certificado y una App Registration con todos' -Color White
    Write-Centered -Text 'los permisos necesarios para conexión desatendida a Exchange Online' -Color White
    Write-Centered -Text 'y Microsoft Graph desde GREX365.' -Color White
    Write-Host ''
    Write-Centered -Text 'Necesitas: cuenta de Global Admin del tenant + acceso a internet.' -Color Yellow
    Write-Host ''

    if (Test-Path -LiteralPath $CsvStepsPath) {
        Write-Centered -Text ("Pasos basados en: {0}" -f (Split-Path $CsvStepsPath -Leaf)) -Color DarkGray
        Write-Host ''
    }

    $proceed = Read-Host '¿Continuar? (S/N)'
    if ($proceed -notmatch '^[Ss]') {
        Write-Log 'Asistente cancelado.' 'WARN'
        return
    }

    # --- INPUTS ---

    Write-Host ''
    Write-Host '  Necesito 4 datos. Junto a cada campo verás dónde encontrarlo en Entra Portal.' -ForegroundColor Yellow
    Write-Host ''

    Write-Host '  1) TenantId' -ForegroundColor Cyan
    Show-CertWizardFieldHint -Field 'TenantId'
    $tenantId = $null
    while (-not (Test-GuidString -Value $tenantId)) {
        $tenantId = Read-NonEmptyInput -Prompt '     TenantId (GUID)'
        if (-not (Test-GuidString -Value $tenantId)) {
            Write-Host '     Formato inválido. Esperaba GUID.' -ForegroundColor Yellow
        }
    }

    Write-Host ''
    Write-Host '  2) Nombre de la App Registration' -ForegroundColor Cyan
    Show-CertWizardFieldHint -Field 'AppName'
    $appName = Read-NonEmptyInput -Prompt '     Nombre de la App' -Default 'GREX365-Unattended'
    $appName = $appName -replace '[^a-zA-Z0-9 _\-]', ''
    if ([string]::IsNullOrWhiteSpace($appName)) { $appName = 'GREX365-Unattended' }

    Write-Host ''
    Write-Host '  3) Validez del certificado' -ForegroundColor Cyan
    Show-CertWizardFieldHint -Field 'Years'
    $yearsRaw = Read-NonEmptyInput -Prompt '     Años (1-5)' -Default '2'
    $years = 0
    if (-not [int]::TryParse($yearsRaw, [ref]$years) -or $years -lt 1 -or $years -gt 5) {
        $years = 2
        Write-Host '     Valor fuera de rango, usando 2.' -ForegroundColor Yellow
    }

    Write-Host ''
    Write-Host '  4) Dominio de la organización (Connect-ExchangeOnline)' -ForegroundColor Cyan
    Show-CertWizardFieldHint -Field 'Organization'
    $organization = Read-NonEmptyInput -Prompt '     Dominio (ej: contoso.onmicrosoft.com)'

    $exportFolder = Split-Path -Parent $ConfigPath

    # --- EJECUCIÓN ---

    $state = @{}

    try {
        Step-EnsureModules

        $certInfo = Step-CreateCertificate -AppName $appName -ValidityYears $years -ExportFolder $exportFolder
        $state.Cert = $certInfo

        Step-RestrictPrivateKeyAcl -Thumbprint $certInfo.Thumbprint

        Step-ConnectGraphAdmin -TenantId $tenantId | Out-Null

        $appResult = Step-CreateAppRegistration -DisplayName $appName -CertPublicBytes $certInfo.RawPublic
        $state.App = $appResult.App
        $state.Sp  = $appResult.ServicePrincipal

        $exoInfo = Step-AssignExchangePermission -AppObject $state.App -SpObject $state.Sp
        $state.ExoInfo = $exoInfo

        Step-AssignExchangeAdminRole -SpObject $state.Sp

        Step-SaveConfig -AppObject $state.App -SpObject $state.Sp -CertInfo $certInfo `
                        -TenantId $tenantId -Organization $organization -ConfigPath $ConfigPath

        Step-WaitAndTestExo -AppObject $state.App -CertInfo $certInfo -Organization $organization

        Step-AddGraphPermissions -AppObject $state.App -SpObject $state.Sp -ExoInfo $exoInfo

        Step-AssignDirectoryRoles -SpObject $state.Sp

        Show-WizardStep -Number 28 -Title 'Reconectar Graph con cert para validar token aplicación'
        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            Connect-MgGraph -ClientId $state.App.AppId `
                            -CertificateThumbprint $certInfo.Thumbprint `
                            -TenantId $tenantId -NoWelcome -ErrorAction Stop | Out-Null
            Write-Log "Connect-MgGraph (cert) OK." 'OK'
        }
        catch {
            Write-Log "Connect-MgGraph (cert) falló: $($_.Exception.Message)" 'WARN'
            Write-Log 'Esto puede deberse a propagación, esperando 30s y reintentando...' 'WARN'
            Start-Sleep -Seconds 30
            try {
                Connect-MgGraph -ClientId $state.App.AppId `
                                -CertificateThumbprint $certInfo.Thumbprint `
                                -TenantId $tenantId -NoWelcome -ErrorAction Stop | Out-Null
                Write-Log "Reintentado OK." 'OK'
            }
            catch {
                Write-Log "Reintento falló también. La configuración está guardada; usa el menú principal cuando se haya propagado." 'ERROR'
            }
        }

        Step-FinalTest

        Show-CertSummary -CertInfo $certInfo -AppObject $state.App -ConfigPath $ConfigPath -TenantId $tenantId -Organization $organization
    }
    catch {
        Write-Host ''
        Write-Log ("Asistente abortado: {0}" -f $_.Exception.Message) 'ERROR'
        Write-Host ''
        Write-Host '  Recursos parciales que pueden haber quedado en Entra ID:' -ForegroundColor Yellow
        if ($state.App) { Write-Host ("    App Registration: {0} (AppId={1})" -f $state.App.DisplayName, $state.App.AppId) -ForegroundColor Yellow }
        if ($state.Sp)  { Write-Host ("    Service Principal: {0}" -f $state.Sp.Id) -ForegroundColor Yellow }
        if ($state.Cert) { Write-Host ("    Certificado en CurrentUser\My: {0}" -f $state.Cert.Thumbprint) -ForegroundColor Yellow }
        Write-Host '  Si quieres reintentar, elimina manualmente esos recursos antes.' -ForegroundColor Yellow
    }
}

function Show-CertSummary {
    param($CertInfo, $AppObject, [string]$ConfigPath, [string]$TenantId, [string]$Organization)

    Write-Host ''
    Show-Header -Title 'CERTIFICADO LISTO' -Subtitle 'Resumen final del asistente'

    Write-Host '  ╔════════════════════════════════════════════════════════════╗' -ForegroundColor Green
    Write-Host '  ║              CERT + APP REG CREADOS CON ÉXITO              ║' -ForegroundColor Green
    Write-Host '  ╚════════════════════════════════════════════════════════════╝' -ForegroundColor Green
    Write-Host ''
    Write-Host ("  Certificado") -ForegroundColor Cyan
    Write-Host ("    Thumbprint   : {0}" -f $CertInfo.Thumbprint) -ForegroundColor White
    Write-Host ("    Subject      : {0}" -f $CertInfo.Subject)
    Write-Host ("    Expira       : {0}" -f $CertInfo.NotAfter)
    Write-Host ("    .cer público : {0}" -f $CertInfo.CerPath)
    Write-Host ''
    Write-Host ("  App Registration") -ForegroundColor Cyan
    Write-Host ("    DisplayName  : {0}" -f $AppObject.DisplayName) -ForegroundColor White
    Write-Host ("    AppId        : {0}" -f $AppObject.AppId)
    Write-Host ("    ObjectId     : {0}" -f $AppObject.Id)
    Write-Host ''
    Write-Host ("  Tenant         : {0}" -f $TenantId) -ForegroundColor Cyan
    Write-Host ("  Organization   : {0}" -f $Organization) -ForegroundColor Cyan
    Write-Host ("  Config JSON    : {0}" -f $ConfigPath) -ForegroundColor Cyan
    Write-Host ''
    Write-Host '  Para validar manualmente:' -ForegroundColor Yellow
    Write-Host ("    Connect-ExchangeOnline -AppId '{0}' -CertificateThumbprint '{1}' -Organization '{2}'" -f $AppObject.AppId, $CertInfo.Thumbprint, $Organization) -ForegroundColor DarkGray
    Write-Host ("    Connect-MgGraph -ClientId '{0}' -CertificateThumbprint '{1}' -TenantId '{2}'" -f $AppObject.AppId, $CertInfo.Thumbprint, $TenantId) -ForegroundColor DarkGray
    Write-Host ''
    Write-Host '  Ya puedes usar GREX365 con método de conexión = CERT' -ForegroundColor Green
    Write-Host ''
}
