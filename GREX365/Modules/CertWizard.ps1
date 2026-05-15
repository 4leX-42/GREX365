# CertWizard module
# Automated 29-step App Registration + certificate provisioning for Exchange Online + Microsoft Graph.
# Logic preserved verbatim from legacy Cert-Wizard.ps1; UI refreshed to Microsoft Admin Center style.

$global:GREX365_CW_ExoAppId   = '00000002-0000-0ff1-ce00-000000000000'
$global:GREX365_CW_GraphAppId = '00000003-0000-0000-c000-000000000000'

$global:GREX365_CW_RoleTemplate_ExchangeAdmin = '29232cdf-9323-42fd-ade2-1d097af3e4de'
$global:GREX365_CW_RoleTemplate_UserAdmin     = 'fe930be7-5e62-47db-91af-98c3a49a38b1'
$global:GREX365_CW_RoleTemplate_GroupsAdmin   = 'fdd7a751-b60b-444a-984c-02652fe8fa1c'

$global:GREX365_CW_GraphAppRoles = @(
    'User.ReadWrite.All'
    'Group.ReadWrite.All'
    'GroupMember.ReadWrite.All'
    'Directory.ReadWrite.All'
    'Organization.Read.All'
    'RoleManagement.Read.Directory'
    'UserAuthenticationMethod.ReadWrite.All'
    'Policy.Read.All'
)

function Show-WizardStep {
    param(
        [int]$Number,
        [string]$Title,
        [string]$Description = ''
    )
    Write-Host ''
    Write-Indent
    Write-Host ('Paso {0,2}/29  ' -f $Number) -NoNewline -ForegroundColor DarkGray
    Write-Host $Title -ForegroundColor White
    if ($Description) {
        Write-Indent -Level 2
        Write-Host $Description -ForegroundColor DarkGray
    }
}

function Wait-WithProgress {
    param(
        [int]$Seconds = 60,
        [string]$Activity = 'Propagación de permisos'
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
        $value = Read-Input -Prompt $Prompt -Default $Default
        if (-not [string]::IsNullOrWhiteSpace($value)) { return $value.Trim() }
        Show-WarningBlock -Title 'Valor requerido'
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
            Write-Indent -Level 2
            Write-Host 'GUID único del tenant Entra ID. Entra Portal → Microsoft Entra ID → Overview.' -ForegroundColor DarkGray
        }
        'AppName' {
            Write-Indent -Level 2
            Write-Host 'Nombre visible de la App Registration que se creará.' -ForegroundColor DarkGray
        }
        'Years' {
            Write-Indent -Level 2
            Write-Host 'Validez del certificado X.509 self-signed. Rango 1–5 años.' -ForegroundColor DarkGray
        }
        'Organization' {
            Write-Indent -Level 2
            Write-Host 'Dominio principal del tenant (<tenant>.onmicrosoft.com o dominio verificado).' -ForegroundColor DarkGray
        }
    }
}

function Step-EnsureModules {
    Show-WizardStep -Number 1 -Title 'Verificando módulos PowerShell'
    $required = @(
        'ExchangeOnlineManagement'
        'Microsoft.Graph.Authentication'
        'Microsoft.Graph.Applications'
        'Microsoft.Graph.Identity.DirectoryManagement'
        'Microsoft.Graph.Identity.SignIns'
    )
    foreach ($m in $required) { Ensure-ToolkitModule -ModuleName $m }
    Write-Log 'Módulos OK.' -Level OK -Source 'CertWizard'
}

function Step-CreateCertificate {
    param(
        [Parameter(Mandatory = $true)][string]$AppName,
        [Parameter(Mandatory = $true)][int]$ValidityYears,
        [Parameter(Mandatory = $true)][string]$ExportFolder
    )

    Show-WizardStep -Number 2 -Title 'Generando clave RSA 2048 en memoria'
    $rsa = [System.Security.Cryptography.RSA]::Create(2048)
    Write-Log 'Clave RSA 2048 creada.' -Level OK -Source 'CertWizard'

    Show-WizardStep -Number 3 -Title 'Construyendo CertificateRequest self-signed SHA256'
    $subject = "CN=$AppName"
    $req = New-Object System.Security.Cryptography.X509Certificates.CertificateRequest(
        $subject, $rsa,
        [System.Security.Cryptography.HashAlgorithmName]::SHA256,
        [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
    )

    $req.CertificateExtensions.Add((New-Object System.Security.Cryptography.X509Certificates.X509BasicConstraintsExtension($false, $false, 0, $false)))
    $req.CertificateExtensions.Add((New-Object System.Security.Cryptography.X509Certificates.X509KeyUsageExtension(
        [System.Security.Cryptography.X509Certificates.X509KeyUsageFlags]::DigitalSignature, $false)))
    Write-Log 'Request preparado.' -Level OK -Source 'CertWizard'

    Show-WizardStep -Number 4 -Title ('Firmando self-signed (' + $ValidityYears + ' años)')
    $notBefore = (Get-Date).ToUniversalTime().AddMinutes(-5)
    $notAfter  = (Get-Date).ToUniversalTime().AddYears($ValidityYears)
    $tempCert = $req.CreateSelfSigned([System.DateTimeOffset]$notBefore, [System.DateTimeOffset]$notAfter)
    Write-Log ('Cert en memoria. Thumbprint=' + $tempCert.Thumbprint) -Level OK -Source 'CertWizard'

    Show-WizardStep -Number 5 -Title 'Importando a CurrentUser\My'
    $pfxBytes = $tempCert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx, '')
    $flags = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::PersistKeySet -bor `
             [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::UserKeySet
    $persistedCert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($pfxBytes, '', $flags)

    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store('My', 'CurrentUser')
    $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    $store.Add($persistedCert)
    $store.Close()
    Write-Log ('Cert persistido. Thumbprint=' + $persistedCert.Thumbprint) -Level OK -Source 'CertWizard'

    Show-WizardStep -Number 6 -Title 'Exportando parte pública (.cer)'
    if (-not (Test-Path -LiteralPath $ExportFolder)) {
        New-Item -ItemType Directory -Path $ExportFolder -Force | Out-Null
    }
    $cerPath = Join-Path $ExportFolder ("{0}.cer" -f $AppName)
    [System.IO.File]::WriteAllBytes($cerPath, $persistedCert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))
    Write-Log ('Cert público exportado: ' + $cerPath) -Level OK -Source 'CertWizard'

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

function Step-RestrictPrivateKeyAcl {
    param([Parameter(Mandatory = $true)][string]$Thumbprint)

    Show-WizardStep -Number 7 -Title 'Restringiendo ACL de la clave privada'

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
            Write-Log 'No se localizó contenedor de clave privada. ACL hardening omitido.' -Level WARN -Source 'CertWizard'
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
            Write-Log 'No se localizó archivo físico de la clave. ACL hardening omitido.' -Level WARN -Source 'CertWizard'
            return
        }

        $acl = Get-Acl -LiteralPath $keyFile
        $acl.SetAccessRuleProtection($true, $false)
        $rules = @($acl.Access)
        foreach ($r in $rules) { [void]$acl.RemoveAccessRule($r) }

        $userSid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
        [void]$acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($userSid, 'FullControl', 'Allow')))
        $systemSid = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-18')
        [void]$acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($systemSid, 'FullControl', 'Allow')))

        Set-Acl -LiteralPath $keyFile -AclObject $acl
        Write-Log ('ACL aplicada en: ' + $keyFile) -Level OK -Source 'CertWizard'
    } catch {
        Write-Log ('ACL hardening omitido: ' + $_.Exception.Message) -Level WARN -Source 'CertWizard'
    }
}

function Step-ConnectGraphAdmin {
    param([Parameter(Mandatory = $true)][string]$TenantId)

    Show-WizardStep -Number 8 -Title 'Login Microsoft Graph (Global Admin del tenant)'
    Show-WarningBlock -Title 'Login interactivo' -Detail 'Se abrirá el navegador. Inicia sesión con Global Admin del tenant.'

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
    if (-not $ctx -or -not $ctx.Account) { throw 'Login Graph fallido.' }
    Write-Log ('Login Graph OK. Cuenta=' + $ctx.Account + ' | Tenant=' + $ctx.TenantId) -Level OK -Source 'CertWizard'
    return $ctx
}

function Step-CreateAppRegistration {
    param(
        [Parameter(Mandatory = $true)][string]$DisplayName,
        [Parameter(Mandatory = $true)][byte[]]$CertPublicBytes
    )

    Show-WizardStep -Number 9 -Title ('Creando App Registration: ' + $DisplayName)
    $app = New-MgApplication -DisplayName $DisplayName -SignInAudience 'AzureADMyOrg' -ErrorAction Stop
    Write-Log ('App creada. AppId=' + $app.AppId + ' ObjectId=' + $app.Id) -Level OK -Source 'CertWizard'

    Show-WizardStep -Number 10 -Title 'Adjuntando certificado público a la App'
    $keyCred = @{
        Type = 'AsymmetricX509Cert'
        Usage = 'Verify'
        Key = $CertPublicBytes
        DisplayName = "$DisplayName-cert"
    }
    Update-MgApplication -ApplicationId $app.Id -KeyCredentials @($keyCred) -ErrorAction Stop
    Write-Log 'Certificado asociado a la App.' -Level OK -Source 'CertWizard'

    Show-WizardStep -Number 11 -Title 'Creando Service Principal'
    $sp = New-MgServicePrincipal -AppId $app.AppId -ErrorAction Stop
    Write-Log ('SP creado. ObjectId=' + $sp.Id) -Level OK -Source 'CertWizard'

    return [PSCustomObject]@{ App = $app; ServicePrincipal = $sp }
}

function Get-ExoManageAsAppRoleId {
    $exoSp = Get-MgServicePrincipal -Filter "appId eq '$($global:GREX365_CW_ExoAppId)'" -ErrorAction Stop
    if (-not $exoSp) { throw 'No se encontró el SP de Office 365 Exchange Online en el tenant.' }
    $role = $exoSp.AppRoles | Where-Object { $_.Value -eq 'Exchange.ManageAsApp' } | Select-Object -First 1
    if (-not $role) { throw "No se encontró el AppRole 'Exchange.ManageAsApp'." }
    return [PSCustomObject]@{ Sp = $exoSp; RoleId = $role.Id }
}

function Step-AssignExchangePermission {
    param(
        [Parameter(Mandatory = $true)]$AppObject,
        [Parameter(Mandatory = $true)]$SpObject
    )

    Show-WizardStep -Number 12 -Title 'Declarando RequiredResourceAccess Exchange.ManageAsApp'
    $exoInfo = Get-ExoManageAsAppRoleId
    $rra = @{
        ResourceAppId = $global:GREX365_CW_ExoAppId
        ResourceAccess = @(@{ Id = $exoInfo.RoleId; Type = 'Role' })
    }
    Update-MgApplication -ApplicationId $AppObject.Id -RequiredResourceAccess @($rra) -ErrorAction Stop
    Write-Log 'Permiso Exchange.ManageAsApp declarado.' -Level OK -Source 'CertWizard'

    Show-WizardStep -Number 13 -Title 'Admin consent (AppRoleAssignment)'
    $body = @{ PrincipalId = $SpObject.Id; ResourceId = $exoInfo.Sp.Id; AppRoleId = $exoInfo.RoleId }
    try {
        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $SpObject.Id -BodyParameter $body -ErrorAction Stop | Out-Null
        Write-Log 'Consent Exchange.ManageAsApp aplicado.' -Level OK -Source 'CertWizard'
    } catch {
        if ($_.Exception.Message -match 'Permission_Conflict|already exists') {
            Write-Log 'Consent Exchange ya estaba aplicado.' -Level WARN -Source 'CertWizard'
        } else { throw }
    }
    return $exoInfo
}

function Get-OrActivate-DirectoryRole {
    param([Parameter(Mandatory = $true)][string]$RoleTemplateId)
    $role = Get-MgDirectoryRole -Filter "roleTemplateId eq '$RoleTemplateId'" -ErrorAction SilentlyContinue
    if ($role) { return $role }
    Write-Log ("Activando rol templateId=$RoleTemplateId") -Source 'CertWizard'
    return New-MgDirectoryRole -RoleTemplateId $RoleTemplateId -ErrorAction Stop
}

function Add-RoleMember {
    param(
        [Parameter(Mandatory = $true)][string]$RoleId,
        [Parameter(Mandatory = $true)][string]$PrincipalObjectId
    )
    $body = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$PrincipalObjectId" }
    try {
        New-MgDirectoryRoleMemberByRef -DirectoryRoleId $RoleId -BodyParameter $body -ErrorAction Stop
    } catch {
        if ($_.Exception.Message -match 'already exist|already a member|conflicting object') {
            Write-Log 'SP ya era miembro de este rol.' -Level WARN -Source 'CertWizard'
        } else { throw }
    }
}

function Step-AssignExchangeAdminRole {
    param([Parameter(Mandatory = $true)]$SpObject)

    Show-WizardStep -Number 14 -Title 'Activando rol Exchange Administrator'
    $role = Get-OrActivate-DirectoryRole -RoleTemplateId $global:GREX365_CW_RoleTemplate_ExchangeAdmin
    Write-Log ('Rol Exchange Admin OK. Id=' + $role.Id) -Level OK -Source 'CertWizard'

    Show-WizardStep -Number 15 -Title 'Asignando SP al rol Exchange Administrator'
    Add-RoleMember -RoleId $role.Id -PrincipalObjectId $SpObject.Id
    Write-Log 'SP asignado a Exchange Administrator.' -Level OK -Source 'CertWizard'
}

function Step-SaveConfig {
    param(
        [Parameter(Mandatory = $true)]$AppObject,
        [Parameter(Mandatory = $true)]$SpObject,
        [Parameter(Mandatory = $true)]$CertInfo,
        [Parameter(Mandatory = $true)][string]$TenantId,
        [Parameter(Mandatory = $true)][string]$Organization,
        [Parameter(Mandatory = $true)][string]$ConfigPath
    )

    Show-WizardStep -Number 16 -Title 'Persistiendo parámetros (no secretos) en JSON'
    $payload = [PSCustomObject]@{
        TenantId       = $TenantId
        Organization   = $Organization
        AppId          = $AppObject.AppId
        AppObjectId    = $AppObject.Id
        SpObjectId     = $SpObject.Id
        CertThumbprint = $CertInfo.Thumbprint
        CertSubject    = $CertInfo.Subject
        CertNotAfter   = $CertInfo.NotAfter.ToString('o')
        CerPath        = $CertInfo.CerPath
        CreatedAt      = (Get-Date).ToString('o')
    }

    $folder = Split-Path -Parent $ConfigPath
    if (-not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }

    ($payload | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $ConfigPath -Encoding UTF8
    Write-Log ('Configuración guardada: ' + $ConfigPath) -Level OK -Source 'CertWizard'
}

function Step-WaitAndTestExo {
    param(
        [Parameter(Mandatory = $true)]$AppObject,
        [Parameter(Mandatory = $true)]$CertInfo,
        [Parameter(Mandatory = $true)][string]$Organization
    )

    Show-WizardStep -Number 17 -Title 'Esperando propagación (60s)'
    Wait-WithProgress -Seconds 60

    Show-WizardStep -Number 18 -Title 'Conectando Exchange Online (test cert)'
    Connect-ExchangeOnline `
        -AppId                 $AppObject.AppId `
        -CertificateThumbprint $CertInfo.Thumbprint `
        -Organization          $Organization `
        -ShowBanner:$false `
        -ErrorAction Stop
    Write-Log 'Connect-ExchangeOnline OK (cert).' -Level OK -Source 'CertWizard'

    Show-WizardStep -Number 19 -Title 'Validando operación real (Get-OrganizationConfig)'
    $orgConf = Get-OrganizationConfig -ErrorAction Stop
    Write-Log ('Tenant DisplayName: ' + $orgConf.DisplayName) -Level OK -Source 'CertWizard'
}

function Step-AddGraphPermissions {
    param(
        [Parameter(Mandatory = $true)]$AppObject,
        [Parameter(Mandatory = $true)]$SpObject,
        [Parameter(Mandatory = $true)][PSCustomObject]$ExoInfo
    )

    Show-WizardStep -Number 22 -Title 'Reusando contexto Graph'
    $ctx = Get-MgContext
    if (-not $ctx -or -not $ctx.Account) {
        Write-Log 'Contexto Graph perdido, reconectando...' -Level WARN -Source 'CertWizard'
        Connect-MgGraph -TenantId $ctx.TenantId -Scopes @(
            'Application.ReadWrite.All','AppRoleAssignment.ReadWrite.All',
            'RoleManagement.ReadWrite.Directory','Directory.ReadWrite.All'
        ) -ContextScope Process -NoWelcome -ErrorAction Stop
    }

    Show-WizardStep -Number 23 -Title 'Resolviendo IDs de los 8 AppRoles de Microsoft Graph'
    $graphSp = Get-MgServicePrincipal -Filter "appId eq '$($global:GREX365_CW_GraphAppId)'" -ErrorAction Stop
    if (-not $graphSp) { throw 'No se encontró el SP de Microsoft Graph.' }

    $resolved = New-Object System.Collections.Generic.List[object]
    foreach ($value in $global:GREX365_CW_GraphAppRoles) {
        $r = $graphSp.AppRoles | Where-Object { $_.Value -eq $value -and $_.AllowedMemberTypes -contains 'Application' } | Select-Object -First 1
        if (-not $r) { Write-Log ('AppRole no encontrado: ' + $value) -Level ERROR -Source 'CertWizard'; continue }
        $resolved.Add([PSCustomObject]@{ Value = $value; Id = $r.Id })
        Write-Log ($value + ' → ' + $r.Id) -Level OK -Source 'CertWizard'
    }
    if ($resolved.Count -eq 0) { throw 'No se resolvió ningún AppRole de Graph.' }

    Show-WizardStep -Number 24 -Title 'Actualizando manifest (Exchange + Graph)'
    $rraExchange = @{ ResourceAppId = $global:GREX365_CW_ExoAppId;   ResourceAccess = @(@{ Id = $ExoInfo.RoleId; Type = 'Role' }) }
    $rraGraph    = @{ ResourceAppId = $global:GREX365_CW_GraphAppId; ResourceAccess = @($resolved | ForEach-Object { @{ Id = $_.Id; Type = 'Role' } }) }
    Update-MgApplication -ApplicationId $AppObject.Id -RequiredResourceAccess @($rraExchange, $rraGraph) -ErrorAction Stop
    Write-Log 'Manifest actualizado.' -Level OK -Source 'CertWizard'

    Show-WizardStep -Number 25 -Title 'Admin consent para cada AppRole Graph'
    foreach ($r in $resolved) {
        $body = @{ PrincipalId = $SpObject.Id; ResourceId = $graphSp.Id; AppRoleId = $r.Id }
        try {
            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $SpObject.Id -BodyParameter $body -ErrorAction Stop | Out-Null
            Write-Log ('Consent OK: ' + $r.Value) -Level OK -Source 'CertWizard'
        } catch {
            if ($_.Exception.Message -match 'Permission_Conflict|already exists') {
                Write-Log ('Consent ya aplicado: ' + $r.Value) -Level WARN -Source 'CertWizard'
            } else {
                Write-Log ('Consent fallido ' + $r.Value + ': ' + $_.Exception.Message) -Level ERROR -Source 'CertWizard'
            }
        }
    }
}

function Step-AssignDirectoryRoles {
    param([Parameter(Mandatory = $true)]$SpObject)

    Show-WizardStep -Number 26 -Title 'Asignando rol User Administrator'
    $userAdmin = Get-OrActivate-DirectoryRole -RoleTemplateId $global:GREX365_CW_RoleTemplate_UserAdmin
    Add-RoleMember -RoleId $userAdmin.Id -PrincipalObjectId $SpObject.Id
    Write-Log 'User Administrator asignado.' -Level OK -Source 'CertWizard'

    Show-WizardStep -Number 27 -Title 'Asignando rol Groups Administrator'
    $groupsAdmin = Get-OrActivate-DirectoryRole -RoleTemplateId $global:GREX365_CW_RoleTemplate_GroupsAdmin
    Add-RoleMember -RoleId $groupsAdmin.Id -PrincipalObjectId $SpObject.Id
    Write-Log 'Groups Administrator asignado.' -Level OK -Source 'CertWizard'
}

function Step-FinalTest {
    Show-WizardStep -Number 29 -Title 'Test dual: Graph + Exchange'
    try {
        $u = Get-MgUser -Top 3 -ErrorAction Stop
        Write-Log ('Get-MgUser OK (' + @($u).Count + ' usuarios).') -Level OK -Source 'CertWizard'
    } catch { Write-Log ('Get-MgUser falló: ' + $_.Exception.Message) -Level WARN -Source 'CertWizard' }
    try {
        $g = Get-MgGroup -Top 3 -ErrorAction Stop
        Write-Log ('Get-MgGroup OK (' + @($g).Count + ' grupos).') -Level OK -Source 'CertWizard'
    } catch { Write-Log ('Get-MgGroup falló: ' + $_.Exception.Message) -Level WARN -Source 'CertWizard' }
    try {
        $dl = Get-DistributionGroup -ResultSize 3 -ErrorAction Stop
        Write-Log ('Get-DistributionGroup OK (' + @($dl).Count + ').') -Level OK -Source 'CertWizard'
    } catch { Write-Log ('Get-DistributionGroup falló: ' + $_.Exception.Message) -Level WARN -Source 'CertWizard' }
}

function Show-CertSummary {
    param($CertInfo, $AppObject, [string]$ConfigPath, [string]$TenantId, [string]$Organization)

    Show-Header -Title 'GREX365' -Subtitle 'Certificado provisionado'

    Show-Section -Title 'Certificado'
    Write-KeyValue -Key 'Thumbprint'    -Value $CertInfo.Thumbprint
    Write-KeyValue -Key 'Subject'       -Value $CertInfo.Subject
    Write-KeyValue -Key 'Expira'        -Value ($CertInfo.NotAfter.ToString('yyyy-MM-dd'))
    Write-KeyValue -Key 'Cert público'  -Value $CertInfo.CerPath

    Show-Section -Title 'App Registration'
    Write-KeyValue -Key 'DisplayName' -Value $AppObject.DisplayName
    Write-KeyValue -Key 'AppId'       -Value $AppObject.AppId
    Write-KeyValue -Key 'ObjectId'    -Value $AppObject.Id

    Show-Section -Title 'Tenant'
    Write-KeyValue -Key 'TenantId'     -Value $TenantId
    Write-KeyValue -Key 'Organization' -Value $Organization
    Write-KeyValue -Key 'Config JSON'  -Value $ConfigPath

    Show-Section -Title 'Validación manual (opcional)'
    Write-Indent
    Write-Host ("Connect-ExchangeOnline -AppId '{0}' -CertificateThumbprint '{1}' -Organization '{2}'" -f $AppObject.AppId, $CertInfo.Thumbprint, $Organization) -ForegroundColor DarkGray
    Write-Indent
    Write-Host ("Connect-MgGraph -ClientId '{0}' -CertificateThumbprint '{1}' -TenantId '{2}'" -f $AppObject.AppId, $CertInfo.Thumbprint, $TenantId) -ForegroundColor DarkGray
    Write-Host ''
}

function Test-GuidString {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    $g = [guid]::Empty
    return [guid]::TryParse($Value.Trim(), [ref]$g)
}

function Start-CertificateWizard {
    param(
        [Parameter(Mandatory = $true)][string]$CsvStepsPath,
        [Parameter(Mandatory = $true)][string]$ConfigPath
    )

    Show-Header -Title 'GREX365' -Subtitle 'Asistente de certificado (29 pasos)'

    Write-Indent
    Write-Host 'Provisión automática de certificado + App Registration con permisos completos' -ForegroundColor White
    Write-Indent
    Write-Host 'para Exchange Online y Microsoft Graph (conexión desatendida).' -ForegroundColor White
    Write-Host ''
    Show-WarningBlock -Title 'Requisitos' -Detail "Cuenta Global Admin del tenant + acceso a internet.`nLa clave privada se queda en CurrentUser\My (no sale del equipo)."

    $proceed = Read-Input -Prompt '¿Continuar? (S/N)'
    if ($proceed -notmatch '^[Ss]') {
        Write-Log 'Asistente cancelado.' -Level WARN -Source 'CertWizard'
        return
    }

    Write-Host ''
    Show-Section -Title 'Datos requeridos'

    Write-Indent
    Write-Host '1) TenantId' -ForegroundColor Cyan
    Show-CertWizardFieldHint -Field 'TenantId'
    $tenantId = $null
    while (-not (Test-GuidString -Value $tenantId)) {
        $tenantId = Read-NonEmptyInput -Prompt 'TenantId (GUID)'
        if (-not (Test-GuidString -Value $tenantId)) {
            Show-WarningBlock -Title 'Formato inválido' -Detail 'Esperaba un GUID.'
        }
    }

    Write-Host ''
    Write-Indent
    Write-Host '2) Nombre de la App Registration' -ForegroundColor Cyan
    Show-CertWizardFieldHint -Field 'AppName'
    $appName = Read-NonEmptyInput -Prompt 'Nombre de la App' -Default 'GREX365-Unattended'
    $appName = $appName -replace '[^a-zA-Z0-9 _\-]', ''
    if ([string]::IsNullOrWhiteSpace($appName)) { $appName = 'GREX365-Unattended' }

    Write-Host ''
    Write-Indent
    Write-Host '3) Validez del certificado' -ForegroundColor Cyan
    Show-CertWizardFieldHint -Field 'Years'
    $yearsRaw = Read-NonEmptyInput -Prompt 'Años (1-5)' -Default '2'
    $years = 0
    if (-not [int]::TryParse($yearsRaw, [ref]$years) -or $years -lt 1 -or $years -gt 5) {
        $years = 2
        Show-WarningBlock -Title 'Valor fuera de rango' -Detail 'Usando 2 años.'
    }

    Write-Host ''
    Write-Indent
    Write-Host '4) Dominio de la organización' -ForegroundColor Cyan
    Show-CertWizardFieldHint -Field 'Organization'
    $organization = Read-NonEmptyInput -Prompt 'Dominio (ej: contoso.onmicrosoft.com)'

    $exportFolder = Split-Path -Parent $ConfigPath

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
        Step-SaveConfig -AppObject $state.App -SpObject $state.Sp -CertInfo $certInfo -TenantId $tenantId -Organization $organization -ConfigPath $ConfigPath
        Step-WaitAndTestExo -AppObject $state.App -CertInfo $certInfo -Organization $organization
        Step-AddGraphPermissions -AppObject $state.App -SpObject $state.Sp -ExoInfo $exoInfo
        Step-AssignDirectoryRoles -SpObject $state.Sp

        Show-WizardStep -Number 28 -Title 'Reconectando Graph con cert (validación token aplicación)'
        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            Connect-MgGraph -ClientId $state.App.AppId -CertificateThumbprint $certInfo.Thumbprint -TenantId $tenantId -NoWelcome -ErrorAction Stop | Out-Null
            Write-Log 'Connect-MgGraph (cert) OK.' -Level OK -Source 'CertWizard'
        } catch {
            Write-Log ('Connect-MgGraph (cert) falló: ' + $_.Exception.Message) -Level WARN -Source 'CertWizard'
            Write-Log 'Reintento en 30s...' -Level WARN -Source 'CertWizard'
            Start-Sleep -Seconds 30
            try {
                Connect-MgGraph -ClientId $state.App.AppId -CertificateThumbprint $certInfo.Thumbprint -TenantId $tenantId -NoWelcome -ErrorAction Stop | Out-Null
                Write-Log 'Reintento OK.' -Level OK -Source 'CertWizard'
            } catch {
                Write-Log 'Reintento falló también. La configuración está guardada.' -Level ERROR -Source 'CertWizard'
            }
        }

        Step-FinalTest
        Show-CertSummary -CertInfo $certInfo -AppObject $state.App -ConfigPath $ConfigPath -TenantId $tenantId -Organization $organization
    } catch {
        Show-ErrorBlock -Title 'Asistente abortado' -Detail $_.Exception.Message
        Write-Host ''
        Write-Indent
        Write-Host 'Recursos parciales que pueden haber quedado en Entra:' -ForegroundColor Yellow
        if ($state.App)  { Write-Indent -Level 2; Write-Host ('App Registration: ' + $state.App.DisplayName + ' (AppId=' + $state.App.AppId + ')') -ForegroundColor Yellow }
        if ($state.Sp)   { Write-Indent -Level 2; Write-Host ('Service Principal: ' + $state.Sp.Id) -ForegroundColor Yellow }
        if ($state.Cert) { Write-Indent -Level 2; Write-Host ('Certificado en CurrentUser\My: ' + $state.Cert.Thumbprint) -ForegroundColor Yellow }
        Write-Host ''
    }
}
