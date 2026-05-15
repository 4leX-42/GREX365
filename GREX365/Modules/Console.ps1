# Console module
# Microsoft Admin Center–style minimalist console primitives.
# Replaces Ui.ps1 cyberpunk decorations with structured enterprise output.

$global:GREX365_ConsolePadding = 2
$global:GREX365_ConsoleRuleWidth = 60

function Get-ConsoleWidth {
    try { return [Console]::WindowWidth } catch { return 100 }
}

function Get-RuleWidth {
    $w = Get-ConsoleWidth
    if ($w -lt 40) { return ($w - 4) }
    if ($w -gt 90) { return 80 }
    return ($w - 6)
}

function Write-Indent {
    param([int]$Level = 1)
    Write-Host (' ' * ($global:GREX365_ConsolePadding * $Level)) -NoNewline
}

function Write-Rule {
    param(
        [int]$Width = 0,
        [string]$Color = 'DarkGray'
    )
    if ($Width -le 0) { $Width = Get-RuleWidth }
    Write-Indent
    Write-Host ('─' * $Width) -ForegroundColor $Color
}

function Write-KeyValue {
    param(
        [Parameter(Mandatory = $true)][string]$Key,
        [string]$Value = '',
        [int]$KeyWidth = 18,
        [string]$KeyColor = 'DarkGray',
        [string]$ValueColor = 'Gray',
        [int]$IndentLevel = 1
    )

    $keyText = $Key.PadRight($KeyWidth)
    Write-Indent -Level $IndentLevel
    Write-Host $keyText -NoNewline -ForegroundColor $KeyColor
    Write-Host ': ' -NoNewline -ForegroundColor $KeyColor
    Write-Host $Value -ForegroundColor $ValueColor
}

function Write-Status {
    param(
        [Parameter(Mandatory = $true)][string]$Label,
        [Parameter(Mandatory = $true)][string]$State,
        [int]$LabelWidth = 12,
        [int]$IndentLevel = 1
    )

    $color = switch ($State.ToUpperInvariant()) {
        'CONNECTED'    { 'Green' }
        'ONLINE'       { 'Green' }
        'OK'           { 'Green' }
        'DISCONNECTED' { 'DarkGray' }
        'OFFLINE'      { 'DarkGray' }
        'NOT SET'      { 'DarkGray' }
        'WARN'         { 'Yellow' }
        'FAIL'         { 'Red' }
        default        { 'Gray' }
    }

    Write-Indent -Level $IndentLevel
    Write-Host ($Label.PadRight($LabelWidth)) -NoNewline -ForegroundColor DarkGray
    Write-Host ': ' -NoNewline -ForegroundColor DarkGray
    Write-Host $State -ForegroundColor $color
}

function Show-Section {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [string]$Color = 'White',
        [switch]$NoTopSpace
    )

    if (-not $NoTopSpace) { Write-Host '' }
    Write-Indent
    Write-Host $Title.ToUpperInvariant() -ForegroundColor $Color
    Write-Rule
}

function Show-Header {
    param(
        [string]$Title = 'GREX365',
        [string]$Subtitle = '',
        [hashtable[]]$StatusLines,
        [string]$ActiveMethod = $null,
        [switch]$NoClear
    )

    if (-not $NoClear) { try { Clear-Host } catch {} }
    Write-Host ''

    Write-Indent
    Write-Host $Title -NoNewline -ForegroundColor White
    if ($Subtitle) {
        Write-Host ('   ' + $Subtitle) -ForegroundColor DarkGray
    } else {
        Write-Host ''
    }
    Write-Rule

    $state = Get-SessionState

    $authLabel = if ($ActiveMethod) {
        Format-MethodLabel -Method $ActiveMethod
    } else { 'Not set' }

    $tenantValue = if ($state.TenantId) { $state.TenantDomain } else { '—' }
    $graphValue  = if ($state.GraphConnected) { 'Connected' } else { 'Disconnected' }
    $exoValue    = if ($state.ExoConnected)   { 'Connected' } else { 'Disconnected' }

    Write-Status -Label 'Tenant'   -State $tenantValue -LabelWidth 10
    Write-Status -Label 'Auth'     -State $authLabel   -LabelWidth 10
    Write-Status -Label 'Graph'    -State $graphValue  -LabelWidth 10
    Write-Status -Label 'Exchange' -State $exoValue    -LabelWidth 10

    if ($StatusLines) {
        foreach ($line in $StatusLines) {
            Write-Status -Label ([string]$line.Label) -State ([string]$line.Value) -LabelWidth 10
        }
    }

    Write-Rule
    Write-Host ''
}

function Show-ErrorBlock {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [string]$Detail = '',
        [string]$Hint = ''
    )

    Write-Host ''
    Write-Indent
    Write-Host '[ERROR] ' -NoNewline -ForegroundColor Red
    Write-Host $Title -ForegroundColor Red

    if ($Detail) {
        foreach ($line in ($Detail -split "`r?`n")) {
            Write-Indent -Level 2
            Write-Host $line -ForegroundColor Yellow
        }
    }

    if ($Hint) {
        Write-Indent -Level 2
        Write-Host ('Hint: ' + $Hint) -ForegroundColor DarkCyan
    }
    Write-Host ''
}

function Show-WarningBlock {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [string]$Detail = ''
    )

    Write-Host ''
    Write-Indent
    Write-Host '[WARN]  ' -NoNewline -ForegroundColor Yellow
    Write-Host $Title -ForegroundColor Yellow

    if ($Detail) {
        foreach ($line in ($Detail -split "`r?`n")) {
            Write-Indent -Level 2
            Write-Host $line -ForegroundColor Gray
        }
    }
    Write-Host ''
}

function Read-Input {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [string]$Default = ''
    )

    Write-Indent
    Write-Host $Prompt -NoNewline -ForegroundColor White
    if ($Default) {
        Write-Host (' [' + $Default + ']') -NoNewline -ForegroundColor DarkGray
    }
    Write-Host ' ' -NoNewline
    Write-Host '> ' -NoNewline -ForegroundColor DarkCyan
    $val = Read-Host
    if ([string]::IsNullOrWhiteSpace($val)) { return $Default }
    return $val
}

function Wait-ForKey {
    param(
        [string]$Message = 'Pulsa ENTER o ESC para continuar'
    )

    Write-Host ''
    Write-Indent
    Write-Host $Message -ForegroundColor DarkGray

    do {
        $key = [System.Console]::ReadKey($true)
        switch ($key.Key) {
            'Enter'  { return }
            'Escape' { return }
        }
    } while ($true)
}

# Session state helpers (used by header) — pure read, no side effects.
# Cached to avoid hammering Get-MgContext / Get-ConnectionInformation / Get-AcceptedDomain
# on every menu repaint. TTL is short and the cache is invalidated by Connect/Disconnect.
#
# NOTE: must be $global: (not $script:) because Get-SessionState is invoked from child
# scripts launched via `& $path` (Set-StrictMode Latest). With $script:, the variable
# scope rebinds to the child script where it was never initialized → StrictMode throws.

if (-not (Get-Variable -Name GREX365_SessionStateCache -Scope Global -ErrorAction SilentlyContinue)) {
    $global:GREX365_SessionStateCache = $null
}
if (-not (Get-Variable -Name GREX365_SessionStateCacheTime -Scope Global -ErrorAction SilentlyContinue)) {
    $global:GREX365_SessionStateCacheTime = [datetime]::MinValue
}
if (-not (Get-Variable -Name GREX365_SessionStateCacheTTL -Scope Global -ErrorAction SilentlyContinue)) {
    $global:GREX365_SessionStateCacheTTL = [timespan]::FromSeconds(3)
}

function Reset-SessionStateCache {
    $global:GREX365_SessionStateCache = $null
    $global:GREX365_SessionStateCacheTime = [datetime]::MinValue
}

function Get-SessionState {
    param([switch]$Force)

    if (-not $Force -and $global:GREX365_SessionStateCache -and
        ((Get-Date) - $global:GREX365_SessionStateCacheTime) -lt $global:GREX365_SessionStateCacheTTL) {
        return $global:GREX365_SessionStateCache
    }

    $tenant = $null; $tenantDomain = $null; $account = $null
    $graphConnected = $false; $exoConnected = $false; $exoOrg = $null

    try {
        if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
            $ctx = Get-MgContext -ErrorAction SilentlyContinue
            if ($ctx) {
                if ($ctx.TenantId) { $tenant = [string]$ctx.TenantId }
                $isAppOnly = ([string]$ctx.AuthType -match 'AppOnly')
                if ($isAppOnly -and $ctx.ClientId -and $ctx.TenantId) {
                    $graphConnected = $true
                    $account = "App-only ($($ctx.ClientId))"
                }
                elseif ($ctx.Account) {
                    $graphConnected = $true
                    $account = [string]$ctx.Account
                }
            }
        }
    } catch {}

    try {
        if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
            $info = Get-ConnectionInformation -ErrorAction SilentlyContinue
            if ($info | Where-Object { $_.State -eq 'Connected' }) {
                $exoConnected = $true
                $first = ($info | Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1)
                if ($first -and $first.Organization) { $tenantDomain = [string]$first.Organization }
            }
        }
    } catch {}

    if ($exoConnected -and -not $tenantDomain) {
        try {
            if (Get-Command Get-AcceptedDomain -ErrorAction SilentlyContinue) {
                $dom = Get-AcceptedDomain -ErrorAction SilentlyContinue | Where-Object { $_.Default } | Select-Object -First 1
                if ($dom) { $tenantDomain = [string]$dom.DomainName }
            }
        } catch {}
    }

    if (-not $tenantDomain -and $tenant) { $tenantDomain = $tenant }

    $result = [PSCustomObject]@{
        TenantId       = $tenant
        TenantDomain   = $tenantDomain
        Account        = $account
        ExoConnected   = $exoConnected
        GraphConnected = $graphConnected
        ExoOrgName     = $exoOrg
    }

    $global:GREX365_SessionStateCache = $result
    $global:GREX365_SessionStateCacheTime = Get-Date
    return $result
}

function Format-MethodLabel {
    param([string]$Method)
    switch ($Method) {
        'cert'        { return 'Certificate' }
        'traditional' { return 'Device code' }
        default       { return 'Not set' }
    }
}

# Status panel — verbose system overview (used by preferences screen).

function Show-StatusPanel {
    $conn = Get-ToolkitConnectionState

    Show-Section -Title 'Conexión activa'
    Write-KeyValue -Key 'Tenant ID'       -Value ($conn.TenantId        ? $conn.TenantId        : '—')
    Write-KeyValue -Key 'Cuenta'          -Value ($conn.Account          ? $conn.Account          : '—')
    Write-KeyValue -Key 'Exchange Org'    -Value ($conn.ExchangeOrgName  ? $conn.ExchangeOrgName  : '—')
    Write-KeyValue -Key 'Dominio default' -Value ($conn.DefaultDomain    ? $conn.DefaultDomain    : '—')
    Write-KeyValue -Key 'Graph'           -Value ($conn.GraphConnected   ? 'Connected' : 'Disconnected')
    Write-KeyValue -Key 'Exchange'        -Value ($conn.ExoConnected     ? 'Connected' : 'Disconnected')

    Show-Section -Title 'Módulos requeridos'
    foreach ($mod in (Get-ToolkitModuleStatus)) {
        $state = $mod.Installed ? ('OK · v' + $mod.Version) : 'No instalado'
        Write-KeyValue -Key $mod.Name -Value $state -KeyWidth 42
    }

    Show-Section -Title 'Archivos de configuración'
    foreach ($file in (Get-ToolkitConfigFiles)) {
        $state = $file.Exists ? ('Presente · ' + $file.Path) : 'Ausente'
        Write-KeyValue -Key $file.Name -Value $state -KeyWidth 26
    }
    Write-Host ''
}

# CSV format hint — discreet, no fanfare.

function Show-CsvFormatHint {
    param(
        [ValidateSet('EmailId','EmailGroupName')]
        [string]$Schema
    )

    $hint = 'Guía de formato CSV → docs/CSV-Schemas.html'
    Write-Indent -Level 2
    Write-Host $hint -ForegroundColor DarkCyan
}

# --- ADMIN CENTER DEEP LINKS ---
# Generate clickable hyperlinks (OSC 8) to the relevant Microsoft admin portal
# for objects created/touched by the toolkit.

function Get-AdminCenterUrl {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Microsoft365Group','UnifiedGroup','DistributionList','DistributionGroup','MailEnabledSecurityGroup','SecurityGroup','UserMailbox','SharedMailbox','Mailbox','User')]
        [string]$Type,

        [Parameter(Mandatory)]
        [string]$Id
    )

    # EAC deep-link compound: ruta principal + sub-ruta GroupDetails que abre el flyout
    # del grupo en la pestaña Members (/2). Confirmado por probar URLs reales.
    # Para SecurityGroup puro (no mail-enabled) EAC no aplica; cae en M365 Admin Center.
    switch ($Type) {
        { $_ -in 'Microsoft365Group','UnifiedGroup' } {
            return "https://admin.exchange.microsoft.com/#/groups/microsoft365/$Id/general/:/GroupDetails/$Id/2"
        }
        { $_ -in 'DistributionList','DistributionGroup' } {
            return "https://admin.exchange.microsoft.com/#/groups/distributionlist/$Id/general/:/GroupDetails/$Id/2"
        }
        'MailEnabledSecurityGroup' {
            return "https://admin.exchange.microsoft.com/#/groups/mailenabledsecurity/$Id/general/:/GroupDetails/$Id/2"
        }
        'SecurityGroup' {
            return "https://admin.microsoft.com/AdminPortal/Home#/groups/:/GroupDetails/$Id/Members"
        }
        { $_ -in 'UserMailbox','SharedMailbox','Mailbox' } {
            return "https://admin.microsoft.com/AdminPortal/Home#/users/UserDetails/$Id/Mail"
        }
        'User' {
            return "https://admin.microsoft.com/AdminPortal/Home#/users/UserDetails/$Id"
        }
    }
}

function Format-Hyperlink {
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][string]$Text
    )

    # PS 7.2+ exposes $PSStyle.FormatHyperlink for OSC 8 sequences.
    if ($PSStyle -and ($PSStyle.PSObject.Methods.Name -contains 'FormatHyperlink')) {
        return $PSStyle.FormatHyperlink($Text, $Url)
    }
    $esc = [char]27
    return "$esc]8;;$Url$esc\$Text$esc]8;;$esc\"
}

function Show-AdminLink {
    param(
        [Parameter(Mandatory)][string]$Type,
        [Parameter(Mandatory)][string]$Id,
        [string]$Label,
        [string]$DisplayName = ''
    )

    if (-not $Id) { return }

    if (-not $Label) {
        $Label = switch -Wildcard ($Type) {
            'Microsoft365*'        { 'Abrir grupo M365 en Admin Center' }
            'UnifiedGroup'         { 'Abrir grupo M365 en Admin Center' }
            'Distribution*'        { 'Abrir DL en Admin Center' }
            'MailEnabled*'         { 'Abrir grupo mail-enabled en Admin Center' }
            'SecurityGroup'        { 'Abrir grupo de seguridad en Admin Center' }
            { $_ -like '*Mailbox' } { 'Abrir buzón en Admin Center' }
            'User'                 { 'Abrir usuario en Admin Center' }
            default                { 'Abrir en Admin Center' }
        }
    }

    $url = Get-AdminCenterUrl -Type $Type -Id $Id
    if (-not $url) { return }

    $linkText = '→ ' + $Label
    if ($DisplayName) { $linkText += " ({0})" -f $DisplayName }

    $hyperlink = Format-Hyperlink -Url $url -Text $linkText
    Write-Indent
    Write-Host $hyperlink -ForegroundColor Cyan
}
