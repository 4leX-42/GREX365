# Audit module
# Append-only structured JSON-lines audit trail for compliance / forensics.
# File: GREX365/logs/audit/YYYY-MM-DD.jsonl
# Each line is a single JSON object. Compatible with Splunk / ELK / Sentinel ingestion.
#
# Lifecycle:
#   Start-AuditOperation -Operation 'Add-GroupMembers' -Metadata @{ group = '...' }
#     ... operation runs, optionally emits Write-AuditEvent ...
#   Stop-AuditOperation -Result OK -Properties @{ added = 5; failed = 0 }

if (-not (Get-Variable -Name GREX365_AuditContext -Scope Global -ErrorAction SilentlyContinue)) {
    $global:GREX365_AuditContext = $null
}
if (-not (Get-Variable -Name GREX365_AuditDisabled -Scope Global -ErrorAction SilentlyContinue)) {
    $global:GREX365_AuditDisabled = $false
}

function Get-AuditFolder {
    if (-not $global:GREX365_BasePath) { throw "BasePath no inicializado. Audit requiere Main.ps1." }
    $folder = Join-Path $global:GREX365_BasePath 'logs\audit'
    if (-not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
    return $folder
}

function Get-AuditFile {
    $folder = Get-AuditFolder
    return (Join-Path $folder ((Get-Date -Format 'yyyy-MM-dd') + '.jsonl'))
}

function Get-AuditTenantSnapshot {
    $tenant = $null; $domain = $null; $account = $null
    try {
        if ($global:GREX365_SessionStateCache) {
            $s = $global:GREX365_SessionStateCache
            $tenant  = $s.TenantId
            $domain  = $s.TenantDomain
            $account = $s.Account
        }
    } catch {}
    return [PSCustomObject]@{
        TenantId     = [string]$tenant
        TenantDomain = [string]$domain
        Account      = [string]$account
    }
}

function Start-AuditOperation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Operation,
        [hashtable]$Metadata
    )

    $tenantSnap = Get-AuditTenantSnapshot
    $correlationId = ([Guid]::NewGuid().ToString('N')).Substring(0, 12)

    $global:GREX365_AuditContext = [PSCustomObject]@{
        CorrelationId = $correlationId
        Operation     = $Operation
        Started       = Get-Date
        TenantId      = $tenantSnap.TenantId
        TenantDomain  = $tenantSnap.TenantDomain
        Account       = $tenantSnap.Account
        User          = [string]$env:USERNAME
        Host          = [string]$env:COMPUTERNAME
        Metadata      = if ($Metadata) { $Metadata } else { @{} }
    }

    Write-AuditEvent -EventType 'OperationStart' -Properties @{
        metadata = $Metadata
    }
    return $global:GREX365_AuditContext
}

function Stop-AuditOperation {
    [CmdletBinding()]
    param(
        [ValidateSet('OK','ERROR','PARTIAL','CANCELLED')]
        [string]$Result = 'OK',
        [hashtable]$Properties
    )

    if (-not $global:GREX365_AuditContext) { return }
    $ctx = $global:GREX365_AuditContext

    $duration = ((Get-Date) - $ctx.Started).TotalMilliseconds
    $props = @{
        result      = $Result
        duration_ms = [int]$duration
    }
    if ($Properties) {
        foreach ($k in $Properties.Keys) { $props[$k] = $Properties[$k] }
    }
    Write-AuditEvent -EventType 'OperationEnd' -Properties $props
    $global:GREX365_AuditContext = $null
}

function Write-AuditEvent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$EventType,
        [hashtable]$Properties
    )

    if ($global:GREX365_AuditDisabled) { return }

    $entry = [ordered]@{
        timestamp  = (Get-Date).ToString('o')
        eventType  = $EventType
        toolkit    = 'GREX365'
        psVersion  = $PSVersionTable.PSVersion.ToString()
        user       = [string]$env:USERNAME
        host       = [string]$env:COMPUTERNAME
    }

    if ($global:GREX365_AuditContext) {
        $entry['correlationId'] = $global:GREX365_AuditContext.CorrelationId
        $entry['operation']     = $global:GREX365_AuditContext.Operation
        $entry['tenantId']      = $global:GREX365_AuditContext.TenantId
        $entry['tenantDomain']  = $global:GREX365_AuditContext.TenantDomain
        $entry['account']       = $global:GREX365_AuditContext.Account
    } else {
        $snap = Get-AuditTenantSnapshot
        $entry['correlationId'] = '-'
        $entry['tenantId']      = $snap.TenantId
        $entry['tenantDomain']  = $snap.TenantDomain
        $entry['account']       = $snap.Account
    }

    if ($Properties) {
        foreach ($k in $Properties.Keys) { $entry[$k] = $Properties[$k] }
    }

    $json = $entry | ConvertTo-Json -Compress -Depth 8

    $file = Get-AuditFile
    $attempts = 0
    while ($attempts -lt 5) {
        try {
            Add-Content -LiteralPath $file -Value $json -Encoding UTF8 -ErrorAction Stop
            return
        } catch {
            $attempts++
            Start-Sleep -Milliseconds (50 + (Get-Random -Maximum 80))
        }
    }
}

function Get-AuditContext {
    return $global:GREX365_AuditContext
}

function Get-AuditCorrelationId {
    if ($global:GREX365_AuditContext) { return $global:GREX365_AuditContext.CorrelationId }
    return $null
}
