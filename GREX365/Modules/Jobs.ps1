# Jobs module
# Persistent background job queue using Start-ThreadJob.
# State file: logs/jobs/queue.json — survives sessions, idempotent registration.
# Note: Start-ThreadJob lives in same process; jobs lost on full pwsh exit.
# This module tracks descriptors so progress can be queried/resumed.

function Get-JobsFolder {
    if (-not $global:GREX365_BasePath) { throw 'BasePath no inicializado.' }
    $folder = Join-Path $global:GREX365_BasePath 'logs\jobs'
    if (-not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
    return $folder
}

function Get-JobsQueueFile {
    return (Join-Path (Get-JobsFolder) 'queue.json')
}

function Read-JobsQueue {
    $file = Get-JobsQueueFile
    if (-not (Test-Path -LiteralPath $file)) { return @() }
    try {
        $raw = Get-Content -LiteralPath $file -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return @() }
        return @($raw | ConvertFrom-Json)
    } catch { return @() }
}

function Save-JobsQueue {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Queue)
    $file = Get-JobsQueueFile
    $json = if ($Queue.Count -eq 0) { '[]' } else { ($Queue | ConvertTo-Json -Depth 6) }
    Set-Content -LiteralPath $file -Value $json -Encoding UTF8
}

function Upsert-JobDescriptor {
    param(
        [Parameter(Mandatory = $true)][PSCustomObject]$Job
    )
    $queue = @(Read-JobsQueue)
    $idx = -1
    for ($i = 0; $i -lt $queue.Count; $i++) {
        if ($queue[$i].Id -eq $Job.Id) { $idx = $i; break }
    }
    if ($idx -ge 0) { $queue[$idx] = $Job } else { $queue += $Job }
    Save-JobsQueue -Queue $queue
}

function Start-ToolkitJob {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][scriptblock]$ScriptBlock,
        [object[]]$ArgumentList,
        [hashtable]$Metadata
    )

    if (-not (Get-Command Start-ThreadJob -ErrorAction SilentlyContinue)) {
        throw 'Start-ThreadJob no disponible. Instala el módulo ThreadJob.'
    }

    $id = [Guid]::NewGuid().ToString('N').Substring(0, 12)
    $tj = Start-ThreadJob -Name ("GREX365-$Name-$id") -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList

    $desc = [PSCustomObject]@{
        Id          = $id
        Name        = $Name
        ThreadJobId = [int]$tj.Id
        State       = [string]$tj.State
        Started     = (Get-Date).ToString('o')
        Finished    = $null
        Metadata    = if ($Metadata) { $Metadata } else { @{} }
    }
    Upsert-JobDescriptor -Job $desc

    Write-Log ("Job '$Name' arrancado (id=$id thread=$($tj.Id))") -Level OK -Source 'Jobs'
    return $desc
}

function Sync-JobsState {
    $queue = @(Read-JobsQueue)
    $updated = $false
    for ($i = 0; $i -lt $queue.Count; $i++) {
        $d = $queue[$i]
        if ($d.State -in @('Completed','Failed','Stopped')) { continue }
        $tj = Get-Job -Id $d.ThreadJobId -ErrorAction SilentlyContinue
        if (-not $tj) { continue }
        if ($tj.State -ne $d.State) {
            $d.State = [string]$tj.State
            if ($tj.State -in @('Completed','Failed','Stopped')) {
                $d.Finished = (Get-Date).ToString('o')
            }
            $updated = $true
        }
    }
    if ($updated) { Save-JobsQueue -Queue $queue }
    return $queue
}

function Get-ToolkitJobs {
    return (Sync-JobsState)
}

function Get-ToolkitJobOutput {
    param([Parameter(Mandatory = $true)][string]$Id)
    $queue = Read-JobsQueue
    $d = $queue | Where-Object { $_.Id -eq $Id } | Select-Object -First 1
    if (-not $d) { throw "Job no encontrado: $Id" }
    $tj = Get-Job -Id $d.ThreadJobId -ErrorAction SilentlyContinue
    if (-not $tj) { throw "ThreadJob no localizable (id=$($d.ThreadJobId)). Posiblemente sesión reiniciada." }
    return (Receive-Job -Id $tj.Id -Keep)
}

function Stop-ToolkitJob {
    param([Parameter(Mandatory = $true)][string]$Id)
    $queue = @(Read-JobsQueue)
    $d = $queue | Where-Object { $_.Id -eq $Id } | Select-Object -First 1
    if (-not $d) { throw "Job no encontrado: $Id" }
    Stop-Job -Id $d.ThreadJobId -ErrorAction SilentlyContinue
    Sync-JobsState | Out-Null
    Write-Log ("Job '$($d.Name)' detenido (id=$Id)") -Level WARN -Source 'Jobs'
}

function Remove-FinishedJobs {
    $queue = @(Sync-JobsState)
    $keep  = $queue | Where-Object { $_.State -notin @('Completed','Failed','Stopped') }
    Save-JobsQueue -Queue @($keep)
    foreach ($d in $queue) {
        if ($d.State -in @('Completed','Failed','Stopped')) {
            try { Remove-Job -Id $d.ThreadJobId -Force -ErrorAction SilentlyContinue } catch {}
        }
    }
}
