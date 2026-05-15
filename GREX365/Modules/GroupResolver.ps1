# GroupResolver module
# Intelligent group lookup across Graph and Exchange Online.
# Ported intact from Common.ps1 (preserves all scoring and dedup logic).

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

        $type = if ($g.GroupTypes -contains 'Unified') { 'Microsoft365Group' }
                elseif ($g.MailEnabled -and $g.SecurityEnabled) { 'MailEnabledSecurityGroup' }
                elseif ($g.MailEnabled) { 'DistributionList' }
                else { 'SecurityGroup' }

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

        $idValue = if ($g.ExternalDirectoryObjectId) { [string]$g.ExternalDirectoryObjectId }
                   elseif ($g.Guid) { [string]$g.Guid }
                   elseif ($g.Identity) { [string]$g.Identity }
                   else { '' }
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
                    default                       { 'DistributionList' }
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
                default                       { 'DistributionList' }
            }
            & $addCandidate $g $type
        }
    } catch {
        try {
            $rcpt = Get-Recipient -RecipientTypeDetails MailUniversalDistributionGroup,MailUniversalSecurityGroup,GroupMailbox -Anr $SearchText -ResultSize $MaxResults -ErrorAction Stop
            foreach ($g in @($rcpt)) {
                $type = switch ($g.RecipientTypeDetails.ToString()) {
                    'GroupMailbox'                { 'Microsoft365Group' }
                    'MailUniversalSecurityGroup'  { 'MailEnabledSecurityGroup' }
                    default                       { 'DistributionList' }
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
                    default                       { 'DistributionList' }
                }
                & $addCandidate $g $type
            }
        } catch {}
    }

    return $candidates
}

function Merge-GroupCandidates {
    param([Parameter(Mandatory = $true)][System.Collections.IEnumerable]$Candidates)

    $seen = @{}
    foreach ($item in @($Candidates | Sort-Object -Property @{Expression='Score';Descending=$true}, 'DisplayName', 'PrimarySmtpAddress')) {
        if (-not $item) { continue }

        $mailKey = if ($item.PrimarySmtpAddress) { Normalize-SearchText -Value ([string]$item.PrimarySmtpAddress) } else { '' }
        $idKey   = if ($item.GroupId) { ([string]$item.GroupId).Trim().ToLowerInvariant() } else { '' }
        $nameKey = if ($item.DisplayName) { Normalize-SearchText -Value ([string]$item.DisplayName) } else { '' }
        $typeKey = if ($item.GroupType) { ([string]$item.GroupType).Trim().ToLowerInvariant() } else { '' }

        if ($idKey) { $key = "id|$idKey" }
        elseif ($typeKey -eq 'microsoft365group' -and $mailKey) { $key = "m365mail|$mailKey" }
        else { $key = "type|$typeKey|mail|$mailKey|name|$nameKey" }

        if (-not $seen.ContainsKey($key)) { $seen[$key] = $item; continue }

        $existing = $seen[$key]
        if (($item.Score -as [int]) -gt ($existing.Score -as [int])) { $seen[$key] = $item; continue }
        if (($item.Score -as [int]) -eq ($existing.Score -as [int]) -and $existing.Source -eq 'Graph' -and $item.Source -eq 'Exchange') {
            $seen[$key] = $item
        }
    }

    $merged = New-Object System.Collections.Generic.List[object]
    foreach ($v in ($seen.Values | Sort-Object -Property @{Expression='Score';Descending=$true}, 'DisplayName', 'PrimarySmtpAddress', 'GroupType')) {
        $merged.Add($v)
    }
    return $merged
}

function Get-GroupTypeBadge {
    param([string]$GroupType)
    switch ($GroupType) {
        'Microsoft365Group'         { return @{ Label = 'M365'; Color = 'Cyan' } }
        'DistributionList'          { return @{ Label = ' DL '; Color = 'Yellow' } }
        'MailEnabledSecurityGroup'  { return @{ Label = 'MSEC'; Color = 'DarkCyan' } }
        'SecurityGroup'             { return @{ Label = ' SEC'; Color = 'DarkGray' } }
        default                     { return @{ Label = '  ? '; Color = 'Gray' } }
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
        try { Clear-Host } catch {}
        Write-Host ''
        Write-Indent
        Write-Host 'Coincidencias' -ForegroundColor White
        Write-Rule

        Write-Indent
        Write-Host ('Búsqueda : ' + $SearchText) -ForegroundColor DarkGray
        Write-Indent
        Write-Host '↑↓ navegar    Enter seleccionar    Esc cancelar' -ForegroundColor DarkGray
        if ($numberBuffer) {
            Write-Indent
            Write-Host ('Número: ' + $numberBuffer) -ForegroundColor Cyan
        }
        Write-Host ''

        for ($i = 0; $i -lt $Options.Count; $i++) {
            $item = $Options[$i]
            $badge = Get-GroupTypeBadge -GroupType $item.GroupType
            $prefix = if ($i -eq $selected) { '> ' } else { '  ' }
            $color  = if ($i -eq $selected) { 'Cyan' } else { 'Gray' }

            Write-Indent
            Write-Host $prefix -NoNewline -ForegroundColor Cyan
            Write-Host ('{0,2}  ' -f ($i + 1)) -NoNewline -ForegroundColor DarkGray
            Write-Host ('[' + $badge.Label + ']  ') -NoNewline -ForegroundColor $badge.Color
            Write-Host ($item.DisplayName + '  ') -NoNewline -ForegroundColor $color
            Write-Host $item.PrimarySmtpAddress -ForegroundColor DarkGray
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
                if ($numberBuffer.Length -gt 0) { $numberBuffer = $numberBuffer.Substring(0, $numberBuffer.Length - 1) }
            }
            default {
                if ($key.KeyChar -match '\d') {
                    $numberBuffer += [string]$key.KeyChar
                    if ($numberBuffer -match '^\d+$') {
                        $preview = [int]$numberBuffer
                        if ($preview -ge 1 -and $preview -le $Options.Count) { $selected = $preview - 1 }
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
        Write-Log 'Si no usas correo completo, introduce al menos 3 caracteres.' -Level ERROR -Source 'GroupResolver'
        return [PSCustomObject]@{
            Found = $false; Cancelled = $false; GroupType = $null
            DisplayName = $null; Alias = $null; PrimarySmtpAddress = $needle
            Identity = $null; GroupId = $null; RawObject = $null
            MatchCount = 0; Source = $null
        }
    }

    Write-Log ('Buscando: ' + $needle) -Source 'GroupResolver'
    $all = New-Object System.Collections.Generic.List[object]

    if (Get-Command Get-EXORecipient -ErrorAction SilentlyContinue) {
        foreach ($it in @(Get-ExchangeGroupCandidates -SearchText $needle -SearchWasExactMail:$exactMail -MaxResults $MaxResults)) {
            $all.Add($it)
        }
    }
    if (Get-Command Get-MgGroup -ErrorAction SilentlyContinue) {
        foreach ($it in @(Get-GraphGroupCandidates -SearchText $needle -SearchWasExactMail:$exactMail -MaxResults $MaxResults)) {
            $all.Add($it)
        }
    }

    $matches = Merge-GroupCandidates -Candidates $all
    Write-Log ('Coincidencias: ' + @($matches).Count) -Source 'GroupResolver'

    if (-not $matches -or $matches.Count -eq 0) {
        return [PSCustomObject]@{
            Found = $false; Cancelled = $false; GroupType = $null
            DisplayName = $null; Alias = $null; PrimarySmtpAddress = $needle
            Identity = $null; GroupId = $null; RawObject = $null
            MatchCount = 0; Source = $null
        }
    }

    $selected = if ($matches.Count -eq 1) { $matches[0] } else { Show-GroupSelectionMenu -Options $matches -SearchText $needle }

    if (-not $selected) {
        return [PSCustomObject]@{
            Found = $false; Cancelled = $true; GroupType = $null
            DisplayName = $null; Alias = $null; PrimarySmtpAddress = $needle
            Identity = $null; GroupId = $null; RawObject = $null
            MatchCount = $matches.Count; Source = $null
        }
    }

    return [PSCustomObject]@{
        Found = $true; Cancelled = $false
        GroupType = $selected.GroupType
        DisplayName = $selected.DisplayName
        Alias = $selected.Alias
        PrimarySmtpAddress = $selected.PrimarySmtpAddress
        Identity = $selected.Identity
        GroupId = $selected.GroupId
        RawObject = $selected.RawObject
        MatchCount = $matches.Count
        Source = $selected.Source
    }
}

function Resolve-GroupByMail {
    param([Parameter(Mandatory = $true)][string]$GroupMail)
    return Find-Group -SearchText $GroupMail
}

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

function Resolve-GroupBySearch {
    param([Parameter(Mandatory = $true)][string]$Prompt)

    Write-Indent
    Write-Host $Prompt -ForegroundColor White
    Write-Indent -Level 2
    Write-Host 'Acepta correo, nombre o alias. Mostrará coincidencias DL / M365 / Security Group.' -ForegroundColor DarkGray

    $raw = Read-Input -Prompt 'Búsqueda'
    $searchText = Normalize-Input -Value $raw
    if ([string]::IsNullOrWhiteSpace($searchText)) {
        throw 'Búsqueda de grupo vacía. Operación cancelada.'
    }

    $resolved = Resolve-GroupForMembership -GroupMail $searchText
    if ($resolved.Cancelled) { throw 'Selección de grupo cancelada por el usuario.' }
    if (-not $resolved.Found) { throw ('No se encontró ningún grupo coincidente con: ' + $searchText) }

    return [PSCustomObject]@{
        GroupType          = [string]$resolved.OperationGroupType
        DisplayType        = [string]$resolved.DisplayType
        Identity           = [string]$resolved.Identity
        Id                 = [string]$resolved.Id
        DisplayName        = [string]$resolved.DisplayName
        PrimarySmtpAddress = [string]$resolved.PrimarySmtpAddress
    }
}
