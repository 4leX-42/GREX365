#requires -Version 7.4
[CmdletBinding()]
param()

# Smoke runner for every operational script.
# Each script runs in an isolated runspace with a hard timeout, mocked
# Read-Input answers, mocked Graph/EXO cmdlets, and panic on hang.
#
# Validates that the script can be *parsed, loaded, and walked* through
# its main flow without recurring binding bugs (Hashtable / $script: / NRE).

$ErrorActionPreference = 'Continue'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$base = Split-Path -Parent (Split-Path -Parent $here)
$global:GREX365_BasePath = Join-Path $base 'GREX365'

# Load real modules
$mods = @('Logging','Console','Validation','Csv','Preferences','Retry','Audit','Report','Roles','Templates','Jobs','Connection','GroupResolver')
foreach ($m in $mods) {
    . (Join-Path $global:GREX365_BasePath "Modules\$m.ps1")
}
Set-CurrentRole -Role 'admin' | Out-Null
Set-CurrentUIMode -Mode 'advanced' | Out-Null

# Prep mock CSVs
$tmp = $env:TEMP
$mockMembersCsv = Join-Path $tmp 'grex_test_members.csv'
"Email;Id`r`ntesteo@es.andersen.com;`r`n" | Set-Content -LiteralPath $mockMembersCsv -Encoding UTF8

$mockGroupsCsv = Join-Path $tmp 'grex_test_groups.csv'
"Action;Type;Name;DisplayName;Alias;Owners;Members;HiddenFromGAL`r`nensure;DL;testeo-smoke@es.andersen.com;testeo smoke;testeo-smoke;;testeo@es.andersen.com;false`r`n" | Set-Content -LiteralPath $mockGroupsCsv -Encoding UTF8

$mockPermsCsv = Join-Path $tmp 'grex_test_perms.csv'
"Action;Permission;Mailbox;Principal`r`nadd;FullAccess;testeo@es.andersen.com;testeo6@es.andersen.com`r`n" | Set-Content -LiteralPath $mockPermsCsv -Encoding UTF8

# NewGroups CSV needs an existing folder with at least one input CSV
$newGroupsInputDir = Join-Path $tmp 'grex_test_newgroups'
if (-not (Test-Path $newGroupsInputDir)) { New-Item -ItemType Directory -Path $newGroupsInputDir -Force | Out-Null }
$inputForNewGroups = Join-Path $newGroupsInputDir 'sample.csv'
"Email`r`ntesteo@es.andersen.com`r`n" | Set-Content -LiteralPath $inputForNewGroups -Encoding UTF8

# ===== Per-script bootstrap that runs INSIDE its own runspace =====
# Injects the mocks, sets up answers, then dot-sources the target script.

$RunspacePool = [runspacefactory]::CreateRunspacePool(1, 3)
$RunspacePool.Open()
$failures = 0
$results  = New-Object System.Collections.Generic.List[object]

$Script:bootstrap = {
    param($BasePath, $ScriptPath, $Answers, $ScriptParams, $TmpDir)

    $global:GREX365_BasePath = $BasePath
    $modList = 'Logging','Console','Validation','Csv','Preferences','Retry','Audit','Report','Roles','Templates','Jobs','Connection','GroupResolver'
    foreach ($m in $modList) {
        . (Join-Path $BasePath ("Modules\$m.ps1"))
    }

    $global:__answers = [System.Collections.Generic.Queue[string]]::new()
    foreach ($a in $Answers) { $global:__answers.Enqueue([string]$a) }

    # === Mocks at GLOBAL scope so child script sees them ===
    function global:Assert-RequiredServicesReady { }
    function global:Connect-RequiredServices { param($MgGraph,$ExchangeOnline,$GraphScopes,$Method) }
    function global:Disconnect-AllServices { }

    function global:Read-Input { param([string]$Prompt,[string]$Default = '')
        if ($global:__answers.Count -gt 0) { return $global:__answers.Dequeue() }
        if ($Default) { return $Default }
        return 'S'
    }
    function global:Read-ValidatedCsvPath { param([string]$Prompt)
        if ($global:__answers.Count -gt 0) {
            $p = $global:__answers.Dequeue()
            if (Test-Path -LiteralPath $p) { return (Resolve-Path -LiteralPath $p).Path }
        }
        throw 'Operación cancelada.'
    }
    function global:Read-ValidatedEmail { param([string]$Prompt,[switch]$AllowEmpty)
        if ($global:__answers.Count -gt 0) { return $global:__answers.Dequeue() }
        return 'testeo@es.andersen.com'
    }
    function global:Read-ValidatedFolder { param([string]$Prompt,[string]$Default = '')
        if ($global:__answers.Count -gt 0) { return $global:__answers.Dequeue() }
        return $TmpDir
    }
    function global:Show-MethodSelector { return 'cert' }
    function global:Confirm-DestructiveAction { return $true }

    # Graph mocks
    function global:Get-MgContext { return [PSCustomObject]@{ TenantId='mock-tid'; ClientId='mock-app'; AuthType='AppOnly'; Account='mock@org'; Scopes=@('User.Read.All') } }
    function global:Get-MgUser {
        param($UserId,[switch]$All,[string]$Filter,$Property,$ConsistencyLevel,$ErrorAction,$Top)
        if ($UserId) {
            return [PSCustomObject]@{
                Id='11111111-1111-1111-1111-111111111111'
                UserPrincipalName=$UserId; DisplayName='Mock User'
                AccountEnabled=$true; UserType='Member'
                AssignedLicenses=@()
                Mail=$UserId
                SignInActivity=[PSCustomObject]@{ LastSignInDateTime=(Get-Date) }
            }
        }
        return @()
    }
    function global:New-MgGroup { param($BodyParameter); return [PSCustomObject]@{ Id='22222222-2222-2222-2222-222222222222'; DisplayName='mock'; Mail='mock@mock.com' } }
    function global:Get-MgSubscribedSku { param([switch]$All); return @([PSCustomObject]@{ SkuPartNumber='ENTERPRISEPACK'; ConsumedUnits=10; PrepaidUnits=[PSCustomObject]@{Enabled=20;Warning=0;Suspended=0} }) }
    function global:Get-MgGroup { param([switch]$All,$Property,$Filter,$Top,$ConsistencyLevel,$ErrorAction); return @() }
    function global:Get-MgGroupMember { param($GroupId,[switch]$All,$Top,$ErrorAction); return @() }
    function global:Get-MgGroupOwner { param($GroupId,[switch]$All,$ErrorAction); return @() }
    function global:Get-MgOrganization { return @([PSCustomObject]@{ OnPremisesSyncEnabled=$false; OnPremisesLastSyncDateTime=$null }) }
    function global:Get-MgServiceAnnouncementHealthOverview { param([switch]$All); return @() }
    function global:Get-MgServiceAnnouncementIssue { param([switch]$All); return @() }
    function global:Get-MgReportAuthenticationMethodUserRegistrationDetail { param([switch]$All); return @() }
    function global:Get-MgDirectoryRole { return $null }
    function global:Get-MgDirectoryRoleMember { param($DirectoryRoleId,[switch]$All); return @() }
    function global:Get-MgApplication { param([switch]$All,$Property); return @() }
    function global:Get-MgUserMemberOf { param($UserId,[switch]$All); return @() }
    function global:Get-MgUserAuthenticationMethod { param($UserId,[switch]$All); return @() }
    function global:Update-MgUser { }
    function global:Revoke-MgUserSignInSession { }
    function global:Set-MgUserLicense { }
    function global:Remove-MgGroupMemberByRef { }

    # EXO mocks
    function global:Get-EXORecipient { return $null }
    function global:Get-EXOMailbox { return @() }
    function global:Get-EXOMailboxStatistics { return [PSCustomObject]@{ TotalItemSize=$null } }
    function global:Get-Mailbox { param([string]$Identity); return [PSCustomObject]@{ PrimarySmtpAddress=$Identity; RecipientTypeDetails='UserMailbox'; ProhibitSendQuota='Unlimited'; DisplayName='mock' } }
    function global:Set-Mailbox { }
    function global:Set-MailboxAutoReplyConfiguration { }
    function global:Add-MailboxPermission { }
    function global:Remove-MailboxPermission { }
    function global:Add-RecipientPermission { }
    function global:Remove-RecipientPermission { }
    function global:Get-MailboxPermission { return @([PSCustomObject]@{ AccessRights=@('FullAccess') }) }
    function global:Get-DistributionGroupMember { return @() }
    function global:Add-DistributionGroupMember { }
    function global:Get-DistributionGroup { return @() }
    function global:Remove-DistributionGroupMember { }
    function global:New-DistributionGroup { param($Name,$Alias,$PrimarySmtpAddress,$Type); return [PSCustomObject]@{ PrimarySmtpAddress="$Alias@mock"; Identity=$Alias } }
    function global:Remove-DistributionGroup { }
    function global:Set-DistributionGroup { }
    function global:Get-UnifiedGroupLinks { return @() }
    function global:Add-UnifiedGroupLinks { }
    function global:New-UnifiedGroup { param($DisplayName,$Alias,$PrimarySmtpAddress,$AccessType); return [PSCustomObject]@{ PrimarySmtpAddress="$Alias@mock"; Identity=$Alias } }
    function global:Set-UnifiedGroup { }
    function global:Get-OrganizationConfig { return [PSCustomObject]@{ DisplayName='mock org' } }
    function global:Get-AcceptedDomain { return @([PSCustomObject]@{ Default=$true; DomainName='es.andersen.com' }) }
    function global:Get-ConnectionInformation { return @([PSCustomObject]@{ State='Connected'; TenantId='mock-tid'; Organization='es.andersen.com' }) }
    function global:Get-Recipient { return $null }

    Set-CurrentRole -Role 'admin' | Out-Null
    Set-CurrentUIMode -Mode 'advanced' | Out-Null

    try {
        if ($ScriptParams) { & $ScriptPath @ScriptParams *>&1 | Out-Null }
        else               { & $ScriptPath *>&1 | Out-Null }
        return @{ Ok = $true; Error = $null }
    } catch {
        return @{ Ok = $false; Error = $_.Exception.Message }
    }
}

# ===== Test cases =====
$scriptsToTest = @(
    @{
        Name = 'Add-GroupMembers.ps1'
        Answers = @($mockMembersCsv,'testeo','S')
    }
    @{
        Name = 'Export-GroupMembers.ps1'
        Answers = @('testeo','S',$tmp)
    }
    @{
        Name = 'New-GroupsFromCsv.ps1'
        Answers = @('1')
        Params = @{ FolderPaths = @($newGroupsInputDir); Domain = 'es.andersen.com' }
    }
    @{
        Name = 'Convert-SharedToUserMailbox.ps1'
        Answers = @()
        Params = @{ Identity = 'testeo@es.andersen.com' }
    }
    @{
        Name = 'Set-SharedMailboxPermissions.ps1'
        Answers = @($mockPermsCsv,'S')
    }
    @{
        Name = 'Invoke-GroupsWorkflow.ps1'
        Answers = @($mockGroupsCsv,'S')
    }
    @{
        Name = 'Invoke-IdentityAudit.ps1'
        Answers = @()
    }
    @{
        Name = 'Show-TenantHealth.ps1'
        Answers = @()
    }
    @{
        Name = 'Invoke-OffboardingWizard.ps1'
        Answers = @('testeo224@es.andersen.com','testeo6@es.andersen.com','','Andersen','1','S','CONFIRMAR','testeo224@es.andersen.com')
    }
    @{
        Name = 'Invoke-SelfTest.ps1'
        Answers = @()
    }
)

# Runner with hard timeout
$timeoutSeconds = 25

Write-Host ""
Write-Host "=== Smoke test scripts (timeout ${timeoutSeconds}s por script) ===" -ForegroundColor Cyan
Write-Host ""

foreach ($t in $scriptsToTest) {
    $path = Join-Path $global:GREX365_BasePath "Scripts\$($t.Name)"
    if (-not (Test-Path $path)) {
        Write-Host ('FAIL {0,-40} (no existe)' -f $t.Name) -ForegroundColor Red
        $failures++
        continue
    }

    $ps = [PowerShell]::Create()
    $ps.RunspacePool = $RunspacePool
    [void]$ps.AddScript($Script:bootstrap)
    [void]$ps.AddArgument($global:GREX365_BasePath)
    [void]$ps.AddArgument($path)
    [void]$ps.AddArgument($t.Answers)
    [void]$ps.AddArgument($t.Params)
    [void]$ps.AddArgument($tmp)

    $async = $ps.BeginInvoke()
    $waited = 0
    while (-not $async.IsCompleted -and $waited -lt ($timeoutSeconds * 10)) {
        Start-Sleep -Milliseconds 100
        $waited++
    }

    if (-not $async.IsCompleted) {
        try { $ps.Stop() } catch {}
        try { $ps.Dispose() } catch {}
        Write-Host ('FAIL {0,-40} TIMEOUT ({1}s) — script stuck in interactive loop' -f $t.Name, $timeoutSeconds) -ForegroundColor Red
        $failures++
        continue
    }

    try {
        $output = $ps.EndInvoke($async)
        $r = if ($output -is [array]) { $output[0] } else { $output }
        $okFlag = $false
        $errMsg = ''
        if ($r -is [hashtable]) {
            $okFlag = [bool]$r.Ok
            $errMsg = [string]$r.Error
        } else {
            $okFlag = $true
        }

        # Accepted as OK if no error, or if error matches recoverable patterns
        $acceptablePatterns = @(
            'INPUT_EMPTY','INPUT_INVALID','INPUT_NOTFOUND',
            'No se encontró','No se pudo','no encontrado',
            'requerido','Operación cancelada','vacío','vacía',
            'Tenant esperado','No existe','UPN tecleado','UPN no coincide',
            'cancelado','Cancelado','no contiene encabezados',
            'No hay CSVs','Ninguna carpeta','Recipient inválido',
            'plantilla'
        )
        if (-not $okFlag) {
            foreach ($p in $acceptablePatterns) {
                if ($errMsg -like "*$p*") { $okFlag = $true; break }
            }
        }

        if ($okFlag) {
            $detail = if ($errMsg) { '(' + ($errMsg.Substring(0, [Math]::Min(80,$errMsg.Length))) + ')' } else { '' }
            Write-Host ('OK   {0,-40} {1}' -f $t.Name, $detail) -ForegroundColor Green
        } else {
            Write-Host ('FAIL {0,-40} {1}' -f $t.Name, $errMsg) -ForegroundColor Red
            $failures++
        }
    } catch {
        Write-Host ('FAIL {0,-40} runner: {1}' -f $t.Name, $_.Exception.Message) -ForegroundColor Red
        $failures++
    } finally {
        try { $ps.Dispose() } catch {}
    }
}

$RunspacePool.Close(); $RunspacePool.Dispose()

Write-Host ""
if ($failures -eq 0) {
    Write-Host "OK · todos los scripts pasan el smoke ($($scriptsToTest.Count))." -ForegroundColor Green
    exit 0
} else {
    Write-Host "FAIL · $failures de $($scriptsToTest.Count)" -ForegroundColor Red
    exit 1
}
