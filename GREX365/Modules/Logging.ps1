# Logging module
# Single source of truth for diagnostic output across the toolkit.
# Console output + optional per-session file logging in GREX365/logs/.

$global:GREX365_LogSession = $null

function Get-LogFolder {
    if (-not $global:GREX365_BasePath) { throw "BasePath no inicializado. Logging requiere Main.ps1." }
    $folder = Join-Path $global:GREX365_BasePath 'logs'
    if (-not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
    return $folder
}

function Start-LogSession {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [switch]$AlwaysPersist
    )

    $safeName = ($Name -replace '[^a-zA-Z0-9_-]', '_')
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $folder = Get-LogFolder
    $file = Join-Path $folder ("{0}_{1}.log" -f $safeName, $stamp)

    $global:GREX365_LogSession = [PSCustomObject]@{
        Name           = $Name
        File           = $file
        Buffer         = (New-Object System.Collections.Generic.List[string])
        Started        = Get-Date
        AlwaysPersist  = [bool]$AlwaysPersist
        HasErrors      = $false
        HasSuccess     = $false
    }
    return $global:GREX365_LogSession
}

function Save-LogSession {
    param([switch]$Force)

    if (-not $global:GREX365_LogSession) { return }
    if ($global:GREX365_LogSession.Buffer.Count -eq 0) { return }

    $shouldPersist = $Force.IsPresent -or $global:GREX365_LogSession.AlwaysPersist -or $global:GREX365_LogSession.HasSuccess -or $global:GREX365_LogSession.HasErrors
    if (-not $shouldPersist) { return }

    $folder = Split-Path -Path $global:GREX365_LogSession.File -Parent
    if (-not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
    $global:GREX365_LogSession.Buffer | Set-Content -LiteralPath $global:GREX365_LogSession.File -Encoding UTF8
    return $global:GREX365_LogSession.File
}

function Stop-LogSession {
    param([switch]$Persist)

    if (-not $global:GREX365_LogSession) { return $null }
    $file = $null
    if ($Persist.IsPresent) {
        $file = Save-LogSession -Force
    }
    $global:GREX365_LogSession = $null
    return $file
}

function Get-LogSessionFile {
    if (-not $global:GREX365_LogSession) { return $null }
    return $global:GREX365_LogSession.File
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR','DEBUG')]
        [string]$Level = 'INFO',
        [string]$Source = ''
    )

    $color = switch ($Level) {
        'OK'    { 'Green' }
        'WARN'  { 'Yellow' }
        'ERROR' { 'Red' }
        'DEBUG' { 'DarkGray' }
        default { 'Gray' }
    }

    $time = Get-Date -Format 'HH:mm:ss'
    $tag = switch ($Level) {
        'OK'    { ' OK  ' }
        'WARN'  { 'WARN ' }
        'ERROR' { 'FAIL ' }
        'DEBUG' { 'DBG  ' }
        default { 'INFO ' }
    }

    $prefix = "  $time  $tag "
    if ($Source) { $prefix = "$prefix [$Source] " }

    Write-Host $prefix -NoNewline -ForegroundColor DarkGray
    Write-Host $Message -ForegroundColor $color

    if ($global:GREX365_LogSession) {
        $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $line = "[$stamp] [$Level]"
        if ($Source) { $line += " [$Source]" }
        $line += " $Message"
        [void]$global:GREX365_LogSession.Buffer.Add($line)

        if ($Level -eq 'ERROR') { $global:GREX365_LogSession.HasErrors = $true }
        if ($Level -eq 'OK')    { $global:GREX365_LogSession.HasSuccess = $true }
    }
}
