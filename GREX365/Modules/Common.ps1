# --- LOGGING ---

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet('INFO','OK','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    switch ($Level) {
        'INFO'  { $color = 'Gray' }
        'OK'    { $color = 'Green' }
        'WARN'  { $color = 'Yellow' }
        'ERROR' { $color = 'Red' }
        default { $color = 'White' }
    }

    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

function Ensure-Folder {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Pause-Toolkit {
    Write-Host ""
    Read-Host "Pulsa ENTER para volver al menú principal" | Out-Null
}

function Wait-ForMenuReturn {
    param(
        [string]$Message = "Pulsa ENTER o ESC para volver al menú principal"
    )

    Write-Host ""
    Write-Host $Message -ForegroundColor DarkGray

    do {
        $key = [System.Console]::ReadKey($true)
        switch ($key.Key) {
            'Enter'  { return }
            'Escape' { return }
        }
    } while ($true)
}

# --- HINT FORMATO CSV ---
# Pintar bajo el prompt de ruta CSV. Discreto.

function Show-CsvFormatHint {
    param(
        [Parameter(Mandatory = $false)]
        [ValidateSet('EmailId','EmailGroupName')]
        [string]$Schema
    )

    $hint = '▌ Guía completa · Instrucciones_CSV.html (raíz del script)'
    try {
        $bw  = [Console]::WindowWidth
        $pad = $bw - $hint.Length - 2
        if ($pad -lt 0) { $pad = 0 }
        Write-Host ((' ' * $pad) + $hint) -ForegroundColor DarkCyan
    }
    catch {
        Write-Host ('  ' + $hint) -ForegroundColor DarkCyan
    }
}

# --- ERROR PANEL CYBERPUNK ---
# Bloque ASCII discreto, sin timestamps. Para fallos de UX (input vacío,
# validación, cancelaciones). Acepta título + razón (multilínea OK).

function Show-ErrorPanel {
    param(
        [string]$Title  = 'ERROR',
        [Parameter(Mandatory = $true)][string]$Reason,
        [string]$Hint   = ''
    )

    $w = 70
    $top = '╔' + ('═' * ($w - 2)) + '╗'
    $sep = '╟' + ('─' * ($w - 2)) + '╢'
    $bot = '╚' + ('═' * ($w - 2)) + '╝'

    $title = ('▓▒░  ' + $Title.ToUpperInvariant() + '  ░▒▓')
    $titlePad = $w - 4 - $title.Length
    if ($titlePad -lt 0) { $titlePad = 0 }

    Write-Host ''
    Write-Host ('  ' + $top) -ForegroundColor Red
    Write-Host '  ║ ' -ForegroundColor Red -NoNewline
    Write-Host $title -ForegroundColor Magenta -NoNewline
    Write-Host ((' ' * $titlePad) + ' ') -NoNewline
    Write-Host '║' -ForegroundColor Red
    Write-Host ('  ' + $sep) -ForegroundColor DarkRed

    foreach ($line in ($Reason -split "`r?`n")) {
        $body = '  > ' + $line
        if ($body.Length -gt ($w - 4)) { $body = $body.Substring(0, $w - 4) }
        $padR = $w - 4 - $body.Length
        Write-Host '  ║ ' -ForegroundColor Red -NoNewline
        Write-Host $body -ForegroundColor Yellow -NoNewline
        Write-Host ((' ' * $padR) + ' ') -NoNewline
        Write-Host '║' -ForegroundColor Red
    }

    if (-not [string]::IsNullOrWhiteSpace($Hint)) {
        Write-Host ('  ' + $sep) -ForegroundColor DarkRed
        $hintBody = '  ↳ ' + $Hint
        if ($hintBody.Length -gt ($w - 4)) { $hintBody = $hintBody.Substring(0, $w - 4) }
        $padH = $w - 4 - $hintBody.Length
        Write-Host '  ║ ' -ForegroundColor Red -NoNewline
        Write-Host $hintBody -ForegroundColor DarkCyan -NoNewline
        Write-Host ((' ' * $padH) + ' ') -NoNewline
        Write-Host '║' -ForegroundColor Red
    }

    Write-Host ('  ' + $bot) -ForegroundColor Red
    Write-Host ''
}

# --- ESTADO DEL SISTEMA (panel de preferencias) ---

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

function Get-ToolkitModuleStatus {
    $list = Get-RequiredToolkitModules
    $result = New-Object System.Collections.Generic.List[object]

    foreach ($name in $list) {
        $available = Get-Module -ListAvailable -Name $name -ErrorAction SilentlyContinue |
                     Sort-Object Version -Descending |
                     Select-Object -First 1

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

    $prefsPath = $null
    $certPath  = $null
    try { $prefsPath = Get-PreferencesPath } catch {}
    try { $certPath  = Get-CertParamsPath  } catch {}

    foreach ($entry in @(
        @{ Name = 'user_preferences.json'; Path = $prefsPath; Description = 'Método de conexión + UPN admin tradicional' }
        @{ Name = 'exo-app-params.json';   Path = $certPath;  Description = 'Parámetros App Registration + thumbprint cert' }
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
    $graphConnected = $false

    try {
        if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
            $ctx = Get-MgContext -ErrorAction SilentlyContinue
            if ($ctx) {
                if ($ctx.TenantId) { $tenant = [string]$ctx.TenantId }
                $isAppOnly = ([string]$ctx.AuthType -match 'AppOnly')
                if ($isAppOnly) {
                    if ($ctx.ClientId -and $ctx.TenantId) {
                        $account = "AppOnly: $($ctx.ClientId)"
                        $graphConnected = $true
                    }
                }
                elseif ($ctx.Account) {
                    $account = [string]$ctx.Account
                    $graphConnected = $true
                }
            }
        }
    } catch {}

    $exoConnected = $false
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
                $dom = Get-AcceptedDomain -ErrorAction SilentlyContinue |
                       Where-Object { $_.Default } | Select-Object -First 1
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

# --- BÚSQUEDA INTELIGENTE DE GRUPOS ---
# Funciones extraídas para uso compartido por scripts 1 y 2.

function Normalize-SearchText {
    param([Parameter(Mandatory = $true)][string]$Value)

    $trimmed = $Value.Trim().ToLowerInvariant()
    $formD = $trimmed.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder

    foreach ($ch in $formD.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    return $sb.ToString().Normalize([Text.NormalizationForm]::FormC)
}

function Get-SearchPrefixVariants {
    param([Parameter(Mandatory = $true)][string]$Value)

    $raw = $Value.Trim().ToLowerInvariant()
    $normalized = Normalize-SearchText -Value $Value
    $variants = New-Object System.Collections.Generic.List[string]

    foreach ($candidate in @($raw, $normalized)) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        if ($candidate -notin $variants) { $variants.Add($candidate) }

        if ($candidate.Length -ge 5) {
            $p5 = $candidate.Substring(0, 5)
            if ($p5 -notin $variants) { $variants.Add($p5) }
        }
        if ($candidate.Length -ge 6) {
            $p6 = $candidate.Substring(0, 6)
            if ($p6 -notin $variants) { $variants.Add($p6) }
        }
    }

    return $variants
}

function Get-SearchScore {
    param(
        [Parameter(Mandatory = $true)][string]$Needle,
        [AllowNull()][string]$Name,
        [AllowNull()][string]$Mail,
        [switch]$SearchWasExactMail
    )

    $needleN = Normalize-SearchText -Value $Needle
    $nameN = if ($Name) { Normalize-SearchText -Value $Name } else { '' }
    $mailN = if ($Mail) { Normalize-SearchText -Value $Mail } else { '' }

    if ($SearchWasExactMail) {
        if ($mailN -eq $needleN) { return 100 }
        if ($mailN -like "$needleN*") { return 90 }
        if ($nameN -like "$needleN*") { return 80 }
        if ($mailN -like "*$needleN*") { return 70 }
        if ($nameN -like "*$needleN*") { return 60 }
        return 0
    }

    if ($nameN -eq $needleN) { return 95 }
    if ($mailN -eq $needleN) { return 94 }
    if ($mailN -like "$needleN@*") { return 93 }
    if ($nameN -like "$needleN*") { return 90 }
    if ($mailN -like "$needleN*") { return 85 }
    if ($nameN -like "*$needleN*") { return 70 }
    if ($mailN -like "*$needleN*") { return 65 }
    return 0
}

function Get-GraphGroupCandidates {
    param(
        [Parameter(Mandatory = $true)][string]$SearchText,
        [switch]$SearchWasExactMail,
        [int]$MaxResults = 20
    )

    $candidates = New-Object System.Collections.Generic.List[object]
    $normalized = Normalize-SearchText -Value $SearchText
    $seen = @{}

    $addCandidate = {
        param($g)
        if (-not $g) { return }
        if (-not $g.Mail) { return }

        $score = Get-SearchScore -Needle $normalized -Name $g.DisplayName -Mail $g.Mail -SearchWasExactMail:$SearchWasExactMail
        if ($score -le 0) { return }

        $type = if ($g.GroupTypes -contains 'Unified') {
            'Microsoft365Group'
        }
        elseif ($g.MailEnabled -and $g.SecurityEnabled) {
            'MailEnabledSecurityGroup'
        }
        elseif ($g.MailEnabled) {
            'DistributionList'
        }
        else {
            'SecurityGroup'
        }

        $idKey = if ($g.Id) { ([string]$g.Id).Trim().ToLowerInvariant() } else { '' }
        if ([string]::IsNullOrWhiteSpace($idKey)) {
            $idKey = ('mail|' + ([string]$g.Mail).Trim().ToLowerInvariant())
        }
        if ($seen.ContainsKey($idKey)) { return }
        $seen[$idKey] = $true

        $candidates.Add([PSCustomObject]@{
            DisplayName        = [string]$g.DisplayName
            Alias              = $null
            PrimarySmtpAddress = [string]$g.Mail
            Identity           = [string]$g.Id
            GroupId            = [string]$g.Id
            GroupType          = $type
            Source             = 'Graph'
            Score              = $score
            RawObject          = $g
        })
    }

    if ($SearchWasExactMail) {
        try {
            $safe = $SearchText.Trim().ToLowerInvariant().Replace("'", "''")
            $groups = Get-MgGroup -Top $MaxResults -Filter "mail eq '$safe'" -ConsistencyLevel eventual -Property Id,DisplayName,Mail,GroupTypes,MailEnabled,SecurityEnabled -ErrorAction Stop
            foreach ($g in @($groups)) { & $addCandidate $g }
        } catch {}
        return $candidates
    }

    foreach ($variant in @(Get-SearchPrefixVariants -Value $SearchText)) {
        $safeV = $variant.Replace("'", "''")
        try {
            $byMail = Get-MgGroup -Top $MaxResults -Filter "startswith(mail,'$safeV')" -ConsistencyLevel eventual -Property Id,DisplayName,Mail,GroupTypes,MailEnabled,SecurityEnabled -ErrorAction Stop
            foreach ($g in @($byMail)) { & $addCandidate $g }
        } catch {}
        try {
            $byName = Get-MgGroup -Top $MaxResults -Filter "startswith(displayName,'$safeV')" -ConsistencyLevel eventual -Property Id,DisplayName,Mail,GroupTypes,MailEnabled,SecurityEnabled -ErrorAction Stop
            foreach ($g in @($byName)) { & $addCandidate $g }
        } catch {}
        try {
            $expr = '"displayName:{0}"' -f $safeV
            $bySearch = Get-MgGroup -Top $MaxResults -Search $expr -ConsistencyLevel eventual -Property Id,DisplayName,Mail,GroupTypes,MailEnabled,SecurityEnabled -ErrorAction Stop
            foreach ($g in @($bySearch)) { & $addCandidate $g }
        } catch {}

        if ($candidates.Count -ge $MaxResults) { break }
    }

    return $candidates
}

function Get-ExchangeGroupCandidates {
    param(
        [Parameter(Mandatory = $true)][string]$SearchText,
        [switch]$SearchWasExactMail,
        [int]$MaxResults = 20
    )

    $candidates = New-Object System.Collections.Generic.List[object]
    $normalized = Normalize-SearchText -Value $SearchText
    $seen = @{}

    $addCandidate = {
        param($g, [string]$Type)
        if (-not $g) { return }

        $mail = if ($g.PrimarySmtpAddress) { [string]$g.PrimarySmtpAddress } else { $null }
        $score = Get-SearchScore -Needle $normalized -Name $g.DisplayName -Mail $mail -SearchWasExactMail:$SearchWasExactMail
        if ($score -le 0) { return }

        $idValue = if ($g.ExternalDirectoryObjectId) { [string]$g.ExternalDirectoryObjectId } elseif ($g.Guid) { [string]$g.Guid } elseif ($g.Identity) { [string]$g.Identity } else { '' }
        $idKey = $idValue.Trim().ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($idKey)) {
            $mailKey = if ($mail) { $mail.ToLowerInvariant() } else { '' }
            $idKey = ('type|' + $Type.ToLowerInvariant() + '|mail|' + $mailKey + '|name|' + ([string]$g.DisplayName).Trim().ToLowerInvariant())
        }
        if ($seen.ContainsKey($idKey)) { return }
        $seen[$idKey] = $true

        $candidates.Add([PSCustomObject]@{
            DisplayName        = [string]$g.DisplayName
            Alias              = [string]$g.Alias
            PrimarySmtpAddress = $mail
            Identity           = [string]$g.Identity
            GroupId            = $idValue
            GroupType          = $Type
            Source             = 'Exchange'
            Score              = $score
            RawObject          = $g
        })
    }

    if ($SearchWasExactMail) {
        try {
            $rcpt = Get-EXORecipient -PrimarySmtpAddress $SearchText -ResultSize 1 -Properties DisplayName,Alias,PrimarySmtpAddress,ExternalDirectoryObjectId,RecipientTypeDetails,Identity -ErrorAction Stop
            foreach ($g in @($rcpt)) {
                $type = switch -Wildcard ($g.RecipientTypeDetails.ToString()) {
                    'GroupMailbox'                { 'Microsoft365Group' }
                    'MailUniversalSecurityGroup'  { 'MailEnabledSecurityGroup' }
                    default                        { 'DistributionList' }
                }
                & $addCandidate $g $type
            }
        } catch {}
        if ($candidates.Count -gt 0) { return $candidates }
    }

    try {
        $rcpt = Get-EXORecipient -Anr $SearchText -ResultSize $MaxResults -Properties DisplayName,Alias,PrimarySmtpAddress,ExternalDirectoryObjectId,RecipientTypeDetails,Identity -ErrorAction Stop |
            Where-Object { $_.RecipientTypeDetails -in @('MailUniversalDistributionGroup','MailUniversalSecurityGroup','GroupMailbox') }
        foreach ($g in @($rcpt)) {
            $type = switch ($g.RecipientTypeDetails.ToString()) {
                'GroupMailbox'                { 'Microsoft365Group' }
                'MailUniversalSecurityGroup'  { 'MailEnabledSecurityGroup' }
                default                        { 'DistributionList' }
            }
            & $addCandidate $g $type
        }
    }
    catch {
        try {
            $rcpt = Get-Recipient -RecipientTypeDetails MailUniversalDistributionGroup,MailUniversalSecurityGroup,GroupMailbox -Anr $SearchText -ResultSize $MaxResults -ErrorAction Stop
            foreach ($g in @($rcpt)) {
                $type = switch ($g.RecipientTypeDetails.ToString()) {
                    'GroupMailbox'                { 'Microsoft365Group' }
                    'MailUniversalSecurityGroup'  { 'MailEnabledSecurityGroup' }
                    default                        { 'DistributionList' }
                }
                & $addCandidate $g $type
            }
        } catch {}
    }

    if (-not $SearchWasExactMail) {
        $safeC = $SearchText.Replace("'", "''")
        try {
            $more = Get-Recipient -RecipientTypeDetails MailUniversalDistributionGroup,MailUniversalSecurityGroup,GroupMailbox `
                -ResultSize $MaxResults `
                -Filter "Alias -like '*$safeC*' -or Name -like '*$safeC*' -or DisplayName -like '*$safeC*' -or PrimarySmtpAddress -like '*$safeC*'" `
                -ErrorAction Stop
            foreach ($g in @($more)) {
                $type = switch ($g.RecipientTypeDetails.ToString()) {
                    'GroupMailbox'                { 'Microsoft365Group' }
                    'MailUniversalSecurityGroup'  { 'MailEnabledSecurityGroup' }
                    default                        { 'DistributionList' }
                }
                & $addCandidate $g $type
            }
        } catch {}
    }

    return $candidates
}

function Merge-GroupCandidates {
    param([Parameter(Mandatory = $true)][System.Collections.IEnumerable]$Candidates)

    $merged = New-Object System.Collections.Generic.List[object]
    $seen = @{}

    foreach ($item in @($Candidates | Sort-Object -Property @{Expression='Score';Descending=$true}, 'DisplayName', 'PrimarySmtpAddress')) {
        if (-not $item) { continue }

        $mailKey = if ($item.PrimarySmtpAddress) { Normalize-SearchText -Value ([string]$item.PrimarySmtpAddress) } else { '' }
        $idKey   = if ($item.GroupId) { ([string]$item.GroupId).Trim().ToLowerInvariant() } else { '' }
        $nameKey = if ($item.DisplayName) { Normalize-SearchText -Value ([string]$item.DisplayName) } else { '' }
        $typeKey = if ($item.GroupType) { ([string]$item.GroupType).Trim().ToLowerInvariant() } else { '' }

        if ($idKey) {
            $key = "id|$idKey"
        }
        elseif ($typeKey -eq 'microsoft365group' -and $mailKey) {
            $key = "m365mail|$mailKey"
        }
        else {
            $key = "type|$typeKey|mail|$mailKey|name|$nameKey"
        }

        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $item
            continue
        }

        $existing = $seen[$key]
        if (($item.Score -as [int]) -gt ($existing.Score -as [int])) {
            $seen[$key] = $item
            continue
        }
        if (($item.Score -as [int]) -eq ($existing.Score -as [int]) -and $existing.Source -eq 'Graph' -and $item.Source -eq 'Exchange') {
            $seen[$key] = $item
        }
    }

    foreach ($v in ($seen.Values | Sort-Object -Property @{Expression='Score';Descending=$true}, 'DisplayName', 'PrimarySmtpAddress', 'GroupType')) {
        $merged.Add($v)
    }
    return $merged
}

function Get-GroupTypeBadge {
    param([string]$GroupType)
    switch ($GroupType) {
        'Microsoft365Group'         { return @{ Label='[M365]'; Color='Cyan' } }
        'DistributionList'          { return @{ Label=' [DL] '; Color='Yellow' } }
        'MailEnabledSecurityGroup'  { return @{ Label='[MSEC]'; Color='Magenta' } }
        'SecurityGroup'             { return @{ Label='[SEC] '; Color='DarkMagenta' } }
        default                     { return @{ Label='[ ?  ]'; Color='Gray' } }
    }
}

function Show-GroupSelectionMenu {
    param(
        [Parameter(Mandatory = $true)][System.Collections.IList]$Options,
        [Parameter(Mandatory = $true)][string]$SearchText
    )

    if (-not $Options -or $Options.Count -eq 0) { return $null }

    $selected = 0
    $numberBuffer = ''

    while ($true) {
        Clear-Host
        if (Get-Command Show-Header -ErrorAction SilentlyContinue) {
            try { Show-Header } catch {}
        }

        Write-Centered 'COINCIDENCIAS ENCONTRADAS' Yellow
        Write-Host ''
        Write-Centered ("Búsqueda: {0}" -f $SearchText) Gray
        Write-Centered 'Usa flechas ↑ ↓, número y ENTER, o ESC = cancelar' DarkGray
        if ($numberBuffer) {
            Write-Centered ("Selección numérica: {0}" -f $numberBuffer) Cyan
        }
        Write-Host ''

        for ($i = 0; $i -lt $Options.Count; $i++) {
            $item = $Options[$i]
            $prefix = if ($i -eq $selected) { '➜' } else { ' ' }
            $badge = Get-GroupTypeBadge -GroupType $item.GroupType
            $line = "{0} [{1,2}] {2} {3} | {4}" -f $prefix, ($i + 1), $badge.Label, $item.DisplayName, $item.PrimarySmtpAddress
            $color = if ($i -eq $selected) { 'Cyan' } else { 'Gray' }
            Write-Centered $line $color
        }

        $key = [System.Console]::ReadKey($true)

        switch ($key.Key) {
            'UpArrow'   { if ($selected -gt 0) { $selected-- }; $numberBuffer = '' }
            'DownArrow' { if ($selected -lt ($Options.Count - 1)) { $selected++ }; $numberBuffer = '' }
            'Enter' {
                if ($numberBuffer -match '^\d+$') {
                    $n = [int]$numberBuffer
                    if ($n -ge 1 -and $n -le $Options.Count) { return $Options[$n - 1] }
                    $numberBuffer = ''
                } else { return $Options[$selected] }
            }
            'Escape' { return $null }
            'Backspace' {
                if ($numberBuffer.Length -gt 0) {
                    $numberBuffer = $numberBuffer.Substring(0, $numberBuffer.Length - 1)
                }
            }
            default {
                if ($key.KeyChar -match '\d') {
                    $numberBuffer += [string]$key.KeyChar
                    if ($numberBuffer -match '^\d+$') {
                        $preview = [int]$numberBuffer
                        if ($preview -ge 1 -and $preview -le $Options.Count) {
                            $selected = $preview - 1
                        }
                    }
                }
            }
        }
    }
}

function Find-Group {
    param(
        [Parameter(Mandatory = $true)][string]$SearchText,
        [int]$MaxResults = 20
    )

    $needle = Normalize-SearchText -Value $SearchText
    $exactMail = $needle -like '*@*'

    if (-not $exactMail -and $needle.Length -lt 3) {
        Write-Log "Si no usas correo completo, introduce al menos 3 caracteres." 'ERROR'
        return [PSCustomObject]@{
            Found              = $false
            Cancelled          = $false
            GroupType          = $null
            DisplayName        = $null
            Alias              = $null
            PrimarySmtpAddress = $needle
            Identity           = $null
            GroupId            = $null
            RawObject          = $null
            MatchCount         = 0
            Source             = $null
        }
    }

    Write-Log "Buscando: $needle" 'INFO'
    $all = New-Object System.Collections.Generic.List[object]

    if (Get-Command Get-EXORecipient -ErrorAction SilentlyContinue) {
        Write-Log "Buscando coincidencias en Exchange..." 'INFO'
        foreach ($it in @(Get-ExchangeGroupCandidates -SearchText $needle -SearchWasExactMail:$exactMail -MaxResults $MaxResults)) {
            $all.Add($it)
        }
    }

    if (Get-Command Get-MgGroup -ErrorAction SilentlyContinue) {
        Write-Log "Buscando coincidencias en Graph..." 'INFO'
        foreach ($it in @(Get-GraphGroupCandidates -SearchText $needle -SearchWasExactMail:$exactMail -MaxResults $MaxResults)) {
            $all.Add($it)
        }
    }

    $matches = Merge-GroupCandidates -Candidates $all
    Write-Log ("{0} coincidencia(s)." -f @($matches).Count) 'INFO'

    if (-not $matches -or $matches.Count -eq 0) {
        return [PSCustomObject]@{
            Found              = $false
            Cancelled          = $false
            GroupType          = $null
            DisplayName        = $null
            Alias              = $null
            PrimarySmtpAddress = $needle
            Identity           = $null
            GroupId            = $null
            RawObject          = $null
            MatchCount         = 0
            Source             = $null
        }
    }

    $selected = if ($matches.Count -eq 1) {
        $matches[0]
    } else {
        Show-GroupSelectionMenu -Options $matches -SearchText $needle
    }

    if (-not $selected) {
        return [PSCustomObject]@{
            Found              = $false
            Cancelled          = $true
            GroupType          = $null
            DisplayName        = $null
            Alias              = $null
            PrimarySmtpAddress = $needle
            Identity           = $null
            GroupId            = $null
            RawObject          = $null
            MatchCount         = $matches.Count
            Source             = $null
        }
    }

    return [PSCustomObject]@{
        Found              = $true
        Cancelled          = $false
        GroupType          = $selected.GroupType
        DisplayName        = $selected.DisplayName
        Alias              = $selected.Alias
        PrimarySmtpAddress = $selected.PrimarySmtpAddress
        Identity           = $selected.Identity
        GroupId            = $selected.GroupId
        RawObject          = $selected.RawObject
        MatchCount         = $matches.Count
        Source             = $selected.Source
    }
}

# Mantenido para compatibilidad con extraccion_user.ps1 (alias).
function Resolve-GroupByMail {
    param([Parameter(Mandatory = $true)][string]$GroupMail)
    return Find-Group -SearchText $GroupMail
}

# Wrapper para script 1: mapea tipos a DistributionGroup/UnifiedGroup
# que es lo que esperan Get-ExistingMemberMap y Add-MemberToTargetGroup.
function Resolve-GroupForMembership {
    param([Parameter(Mandatory = $true)][string]$GroupMail)

    $info = Find-Group -SearchText $GroupMail
    if ($info.Cancelled) { return $info }
    if (-not $info.Found) { return $info }

    $opType = switch ($info.GroupType) {
        'Microsoft365Group'         { 'UnifiedGroup' }
        'DistributionList'          { 'DistributionGroup' }
        'MailEnabledSecurityGroup'  { 'DistributionGroup' }
        default                     { $info.GroupType }
    }

    return [PSCustomObject]@{
        Found              = $true
        Cancelled          = $false
        OperationGroupType = $opType
        DisplayType        = $info.GroupType
        DisplayName        = $info.DisplayName
        Identity           = $info.Identity
        Id                 = $info.GroupId
        PrimarySmtpAddress = $info.PrimarySmtpAddress
        Source             = $info.Source
    }
}
