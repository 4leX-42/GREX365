#requires -Version 5.1
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:LogFile = $null
$script:LogBuffer = New-Object System.Collections.Generic.List[string]
$script:EnableFileLogging = $false
$script:ResultCsv = $null

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet('INFO','WARN','ERROR','OK')]
        [string]$Level = 'INFO'
    )

    $time = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$time] [$Level] $Message"

    switch ($Level) {
        'INFO'  { Write-Host $line -ForegroundColor Cyan }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'OK'    { Write-Host $line -ForegroundColor Green }
    }

    [void]$script:LogBuffer.Add($line)
}

function Save-SuccessLog {
    if (-not $script:EnableFileLogging) { return }
    if ([string]::IsNullOrWhiteSpace($script:LogFile)) { return }
    if ($script:LogBuffer.Count -eq 0) { return }

    $folder = Split-Path -Path $script:LogFile -Parent
    if (-not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }

    $script:LogBuffer | Set-Content -Path $script:LogFile -Encoding UTF8
}

$Host.UI.RawUI.WindowTitle = 'Inyección de usuarios | G365-DL'

function Get-CenterPadding {
    param([Parameter(Mandatory = $true)][string]$Text)
    $width = [Console]::WindowWidth
    $pad = [Math]::Floor(($width - $Text.Length) / 2)
    if ($pad -lt 0) { $pad = 0 }
    return (' ' * $pad)
}

function Write-Centered {
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [ConsoleColor]$Color = [ConsoleColor]::Gray
    )

    $pad = Get-CenterPadding -Text $Text
    Write-Host ($pad + $Text) -ForegroundColor $Color
}

function Show-Header {
    Clear-Host
    Write-Host ''
    Write-Centered '◀◀◀ ════════════  Grex365  ════════════ ▶▶▶' -Color DarkCyan
    Write-Host ''
    Write-Centered 'INYECCION DE USUARIOS | G365-DL' -Color Green
    Write-Centered '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━' -Color Cyan
    Write-Centered 'IMPORTACION MASIVA | EMAIL + ID' -Color DarkCyan
    Write-Host ''
}

function Normalize-QuotedInput {
    param([AllowNull()][string]$Value)

    if ($null -eq $Value) { return '' }

    $value = $Value.Trim()
    if ($value.Length -ge 2) {
        $first = $value.Substring(0, 1)
        $last  = $value.Substring($value.Length - 1, 1)
        if ((($first -eq '"') -and ($last -eq '"')) -or (($first -eq "'") -and ($last -eq "'"))) {
            $value = $value.Substring(1, $value.Length - 2).Trim()
        }
    }

    return $value
}

function Test-LooksLikeCsvPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    if ($Path.Length -lt 7) { return $false }
    if ($Path -notmatch '(?i)\.csv$') { return $false }
    if (-not [System.IO.Path]::IsPathRooted($Path)) { return $false }

    try {
        $invalidChars = [System.IO.Path]::GetInvalidPathChars()
        foreach ($char in $invalidChars) {
            if ($Path.Contains([string]$char)) { return $false }
        }
    }
    catch {
        return $false
    }

    return $true
}

function Read-ValidatedCsvPath {
    param([Parameter(Mandatory = $true)][string]$Prompt)

    Write-Host ''
    Write-Host "  $Prompt" -ForegroundColor White
    if (Get-Command Show-CsvFormatHint -ErrorAction SilentlyContinue) {
        Show-CsvFormatHint -Schema 'EmailId'
    }
    $value = Normalize-QuotedInput -Value (Read-Host '  >')

    if ([string]::IsNullOrWhiteSpace($value)) {
        throw 'La ruta del CSV está vacía. Operación cancelada.'
    }
    if (-not (Test-LooksLikeCsvPath -Path $value)) {
        throw 'La ruta indicada no tiene formato válido de CSV. Ejemplo esperado: C:\Temp\fichero.csv'
    }
    if (-not (Test-Path -LiteralPath $value -PathType Leaf)) {
        throw "No existe el CSV de entrada: $value"
    }

    return (Resolve-Path -LiteralPath $value).Path
}

function Resolve-GroupBySearch {
    param([Parameter(Mandatory = $true)][string]$Prompt)

    Write-Host ''
    Write-Host "  $Prompt" -ForegroundColor White
    Write-Host '  Acepta correo, nombre o alias. Mostrará coincidencias DL / M365 / SecGroup para elegir con flechas o número.' -ForegroundColor DarkGray

    $searchText = Normalize-QuotedInput -Value (Read-Host '  >')
    if ([string]::IsNullOrWhiteSpace($searchText)) {
        throw 'Búsqueda de grupo vacía. Operación cancelada.'
    }

    if (-not (Get-Command Resolve-GroupForMembership -ErrorAction SilentlyContinue)) {
        throw "Función 'Resolve-GroupForMembership' no disponible. Ejecuta el script desde Main.ps1."
    }

    $resolved = Resolve-GroupForMembership -GroupMail $searchText
    if ($resolved.Cancelled) {
        throw 'Selección de grupo cancelada por el usuario.'
    }
    if (-not $resolved.Found) {
        throw "No se encontró ningún grupo coincidente con: $searchText"
    }

    return [PSCustomObject]@{
        GroupType          = [string]$resolved.OperationGroupType
        Identity           = [string]$resolved.Identity
        Id                 = [string]$resolved.Id
        DisplayName        = [string]$resolved.DisplayName
        PrimarySmtpAddress = [string]$resolved.PrimarySmtpAddress
    }
}

function Test-GuidValue {
    param([Parameter(Mandatory = $true)][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    $parsedGuid = [guid]::Empty
    return [guid]::TryParse($Value.Trim(), [ref]$parsedGuid)
}

function Test-EmailValue {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    return ($Value.Trim() -match '^[^@\s]+@[^@\s]+\.[^@\s]+$')
}

function New-ResultObject {
    param(
        [string]$InputEmail,
        [string]$InputId,
        [string]$Email,
        [string]$Id,
        [string]$Estado,
        [string]$Detalle = '',
        [string]$RecipientType = '',
        [string]$PrimarySmtpAddress = '',
        [string]$DisplayName = '',
        [string]$ResolutionSource = ''
    )

    [PSCustomObject]@{
        InputEmail         = $InputEmail
        InputId            = $InputId
        ResolvedEmail      = $Email
        ResolvedId         = $Id
        Estado             = $Estado
        Detalle            = $Detalle
        DisplayName        = $DisplayName
        PrimarySmtpAddress = $PrimarySmtpAddress
        RecipientType      = $RecipientType
        ResolutionSource   = $ResolutionSource
    }
}

function Get-ColumnIndex {
    param(
        [Parameter(Mandatory = $true)][string[]]$Header,
        [Parameter(Mandatory = $true)][string[]]$CandidateNames
    )

    for ($i = 0; $i -lt $Header.Length; $i++) {
        $name = [string]$Header[$i]
        foreach ($candidate in $CandidateNames) {
            if ($name.Trim().ToLowerInvariant() -eq $candidate.Trim().ToLowerInvariant()) {
                return $i
            }
        }
    }

    return -1
}

function Open-SharedFileStream {
    param([Parameter(Mandatory = $true)][string]$Path)

    return [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
}

function Get-FileEncodingObject {
    param([Parameter(Mandatory = $true)][string]$Path)

    $fs = $null
    try {
        $fs = Open-SharedFileStream -Path $Path
        $buffer = New-Object byte[] 4
        $read = $fs.Read($buffer, 0, 4)

        if ($read -ge 2 -and $buffer[0] -eq 0xFF -and $buffer[1] -eq 0xFE) { return [System.Text.Encoding]::Unicode }
        if ($read -ge 2 -and $buffer[0] -eq 0xFE -and $buffer[1] -eq 0xFF) { return [System.Text.Encoding]::BigEndianUnicode }
        if ($read -ge 3 -and $buffer[0] -eq 0xEF -and $buffer[1] -eq 0xBB -and $buffer[2] -eq 0xBF) { return [System.Text.Encoding]::UTF8 }
        return [System.Text.Encoding]::UTF8
    }
    finally {
        if ($null -ne $fs) { $fs.Dispose() }
    }
}

function Get-CsvDelimiter {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][System.Text.Encoding]$Encoding
    )

    $fs = $null
    $reader = $null
    try {
        $fs = Open-SharedFileStream -Path $Path
        $reader = New-Object System.IO.StreamReader($fs, $Encoding, $true)
        $firstLine = $reader.ReadLine()
    }
    catch {
        return ';'
    }
    finally {
        if ($null -ne $reader) { $reader.Dispose() }
        elseif ($null -ne $fs) { $fs.Dispose() }
    }

    if ([string]::IsNullOrWhiteSpace($firstLine)) { return ';' }

    $counts = @{ ';' = 0; ',' = 0; "`t" = 0 }
    $inQuote = $false

    foreach ($ch in $firstLine.ToCharArray()) {
        if ($ch -eq '"') {
            $inQuote = -not $inQuote
            continue
        }
        if ($inQuote) { continue }
        $key = [string]$ch
        if ($counts.ContainsKey($key)) { $counts[$key]++ }
    }

    $best = $counts.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1
    if ($null -ne $best -and $best.Value -gt 0) { return $best.Key }

    return ';'
}

function Split-CsvLine {
    param(
        [Parameter(Mandatory = $true)][string]$Line,
        [Parameter(Mandatory = $true)][string]$Delimiter
    )

    $fields = New-Object System.Collections.Generic.List[string]
    $sb = New-Object System.Text.StringBuilder
    $inQuotes = $false
    $i = 0

    while ($i -lt $Line.Length) {
        $ch = $Line[$i]

        if ($ch -eq '"') {
            if ($inQuotes -and ($i + 1) -lt $Line.Length -and $Line[$i + 1] -eq '"') {
                [void]$sb.Append('"')
                $i += 2
                continue
            }

            $inQuotes = -not $inQuotes
            $i++
            continue
        }

        if (([string]$ch) -eq $Delimiter -and -not $inQuotes) {
            [void]$fields.Add($sb.ToString())
            [void]$sb.Clear()
            $i++
            continue
        }

        [void]$sb.Append($ch)
        $i++
    }

    [void]$fields.Add($sb.ToString())
    return $fields.ToArray()
}

function Get-SharedTextLines {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][System.Text.Encoding]$Encoding
    )

    $lines = New-Object System.Collections.Generic.List[string]
    $fs = $null
    $reader = $null
    try {
        $fs = Open-SharedFileStream -Path $Path
        $reader = New-Object System.IO.StreamReader($fs, $Encoding, $true)
        while (-not $reader.EndOfStream) {
            [void]$lines.Add($reader.ReadLine())
        }
    }
    finally {
        if ($null -ne $reader) { $reader.Dispose() }
        elseif ($null -ne $fs) { $fs.Dispose() }
    }

    return $lines.ToArray()
}

function Import-EmailIdCsvRobust {
    param([Parameter(Mandatory = $true)][string]$Path)

    $encoding = Get-FileEncodingObject -Path $Path
    Write-Log ("Encoding detectado: {0}" -f $encoding.EncodingName)

    $delimiter = Get-CsvDelimiter -Path $Path -Encoding $encoding
    Write-Log ("Delimitador detectado: '{0}'" -f $(if ($delimiter -eq "`t") { 'TAB' } else { $delimiter }))

    $lines = @(Get-SharedTextLines -Path $Path -Encoding $encoding)
    if ($lines.Count -eq 0) { throw 'El CSV está vacío.' }

    $nonEmpty = @($lines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($nonEmpty.Count -eq 0) { throw 'El CSV está vacío o solo contiene líneas en blanco.' }

    $headerFields = @(Split-CsvLine -Line $nonEmpty[0] -Delimiter $delimiter)
    if ($headerFields.Count -eq 0) { throw 'No se pudo leer la cabecera del CSV.' }

    if ($null -ne $headerFields[0]) {
        $headerFields[0] = ([string]$headerFields[0]).TrimStart([char]0xFEFF).Trim()
    }

    $headerFields = @($headerFields | ForEach-Object { if ($null -eq $_) { '' } else { ([string]$_).Trim() } })
    Write-Log ("Cabecera leída: {0}" -f ($headerFields -join ' | '))

    $emailCandidates = @('Email','Mail','UserPrincipalName','UPN','Correo','PrimarySmtpAddress','WindowsLiveID','Login','EmailAddress')
    $idCandidates    = @('Id','ID','ObjectId','ExternalDirectoryObjectId','Guid','GUID','EntraId','AzureObjectId')

    $emailIndex = Get-ColumnIndex -Header $headerFields -CandidateNames $emailCandidates
    $idIndex    = Get-ColumnIndex -Header $headerFields -CandidateNames $idCandidates

    if ($emailIndex -lt 0 -and $idIndex -lt 0) {
        if ($headerFields.Count -ge 2) {
            $emailIndex = 0
            $idIndex = 1
            Write-Log 'No se detectó cabecera estándar. Se usarán columna 1 como Email y columna 2 como Id.' 'WARN'
        }
        elseif ($headerFields.Count -eq 1) {
            $emailIndex = 0
            $idIndex = -1
            Write-Log 'Solo se detectó una columna. Se intentará inferir si contiene Email o Id.' 'WARN'
        }
        else {
            throw 'No se pudo mapear la cabecera del CSV.'
        }
    }

    $emailHeaderName = if ($emailIndex -ge 0 -and $emailIndex -lt $headerFields.Count) { $headerFields[$emailIndex] } else { '' }
    $idHeaderName    = if ($idIndex -ge 0 -and $idIndex -lt $headerFields.Count) { $headerFields[$idIndex] } else { '' }

    Write-Log ("Cabecera detectada correctamente -> Email='{0}' | Id='{1}'" -f $emailHeaderName, $idHeaderName) 'OK'

    $rows = New-Object System.Collections.Generic.List[object]

    for ($lineIndex = 1; $lineIndex -lt $nonEmpty.Count; $lineIndex++) {
        $line = [string]$nonEmpty[$lineIndex]
        $fields = @(Split-CsvLine -Line $line -Delimiter $delimiter)
        if ($fields.Count -eq 0) { continue }

        $email = ''
        $id = ''

        if ($emailIndex -ge 0 -and $emailIndex -lt $fields.Count) {
            $email = [string]$fields[$emailIndex]
        }
        if ($idIndex -ge 0 -and $idIndex -lt $fields.Count) {
            $id = [string]$fields[$idIndex]
        }

        if ($emailIndex -lt 0 -or $idIndex -lt 0) {
            foreach ($field in $fields) {
                $value = ([string]$field).Trim()
                if ([string]::IsNullOrWhiteSpace($value)) { continue }
                if ([string]::IsNullOrWhiteSpace($email) -and (Test-EmailValue -Value $value)) {
                    $email = $value
                    continue
                }
                if ([string]::IsNullOrWhiteSpace($id) -and (Test-GuidValue -Value $value)) {
                    $id = $value
                    continue
                }
            }
        }

        [void]$rows.Add([PSCustomObject]@{
            Email   = $email.Trim()
            Id      = $id.Trim()
            RawText = $line
        })
    }

    return $rows.ToArray()
}

function Split-MultiEmailValue {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return @() }

    $normalized = $Value -replace '[\r\n]+', ' '
    $matches = [regex]::Matches($normalized, '(?i)\b[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}\b')
    $emails = New-Object System.Collections.Generic.List[string]

    foreach ($m in $matches) {
        $mail = $m.Value.Trim().ToLowerInvariant()
        if (-not [string]::IsNullOrWhiteSpace($mail) -and -not $emails.Contains($mail)) {
            [void]$emails.Add($mail)
        }
    }

    if ($emails.Count -gt 0) { return $emails.ToArray() }

    $fallback = $Value.Trim()
    if (-not [string]::IsNullOrWhiteSpace($fallback)) { return @($fallback) }

    return @()
}

function Resolve-GraphIdentity {
    param(
        [string]$Email,
        [string]$Id,
        [hashtable]$Cache
    )

    $resolvedEmail = if ($null -ne $Email) { $Email.Trim() } else { '' }
    $resolvedId = if ($null -ne $Id) { $Id.Trim() } else { '' }

    if (-not [string]::IsNullOrWhiteSpace($resolvedEmail)) {
        $emailKey = 'EMAIL::' + $resolvedEmail.ToLowerInvariant()
        if ($Cache.ContainsKey($emailKey)) { return $Cache[$emailKey] }
    }
    if (-not [string]::IsNullOrWhiteSpace($resolvedId)) {
        $idKey = 'ID::' + $resolvedId.ToLowerInvariant()
        if ($Cache.ContainsKey($idKey)) { return $Cache[$idKey] }
    }

    try {
        if (-not [string]::IsNullOrWhiteSpace($resolvedId) -and (Test-GuidValue -Value $resolvedId)) {
            $u = Get-MgUser -UserId $resolvedId -Property Id,Mail,UserPrincipalName -ErrorAction Stop
            if ($null -ne $u) {
                if ([string]::IsNullOrWhiteSpace($resolvedEmail)) {
                    $resolvedEmail = @($u.Mail, $u.UserPrincipalName) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
                }

                $result = [PSCustomObject]@{
                    Email  = [string]$resolvedEmail
                    Id     = [string]$u.Id
                    Source = 'GraphById'
                }

                if (-not [string]::IsNullOrWhiteSpace($result.Email)) { $Cache['EMAIL::' + $result.Email.ToLowerInvariant()] = $result }
                if (-not [string]::IsNullOrWhiteSpace($result.Id)) { $Cache['ID::' + $result.Id.ToLowerInvariant()] = $result }
                return $result
            }
        }
    }
    catch {}

    try {
        if (-not [string]::IsNullOrWhiteSpace($resolvedEmail)) {
            $safeEmail = $resolvedEmail.Replace("'", "''")
            $users = @(Get-MgUser -Filter "mail eq '$safeEmail' or userPrincipalName eq '$safeEmail'" -Property Id,Mail,UserPrincipalName -ConsistencyLevel eventual -All -ErrorAction Stop)
            $u = $users | Select-Object -First 1
            if ($null -ne $u) {
                $finalEmail = @($u.Mail, $u.UserPrincipalName, $resolvedEmail) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
                $result = [PSCustomObject]@{
                    Email  = [string]$finalEmail
                    Id     = [string]$u.Id
                    Source = 'GraphByEmail'
                }

                if (-not [string]::IsNullOrWhiteSpace($result.Email)) { $Cache['EMAIL::' + $result.Email.ToLowerInvariant()] = $result }
                if (-not [string]::IsNullOrWhiteSpace($result.Id)) { $Cache['ID::' + $result.Id.ToLowerInvariant()] = $result }
                return $result
            }
        }
    }
    catch {}

    $fallback = [PSCustomObject]@{
        Email  = [string]$resolvedEmail
        Id     = [string]$resolvedId
        Source = 'Unchanged'
    }

    if (-not [string]::IsNullOrWhiteSpace($resolvedEmail)) { $Cache['EMAIL::' + $resolvedEmail.ToLowerInvariant()] = $fallback }
    if (-not [string]::IsNullOrWhiteSpace($resolvedId)) { $Cache['ID::' + $resolvedId.ToLowerInvariant()] = $fallback }

    return $fallback
}


function Get-ExistingMemberMap {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('DistributionGroup','UnifiedGroup')][string]$GroupType,
        [Parameter(Mandatory = $true)][string]$Identity
    )

    $map = @{}

    if ($GroupType -eq 'DistributionGroup') {
        $members = @(Get-DistributionGroupMember -Identity $Identity -ResultSize Unlimited -ErrorAction Stop)
        foreach ($member in $members) {
            if ($null -ne $member.PrimarySmtpAddress) {
                $smtp = $member.PrimarySmtpAddress.ToString().Trim()
                if (-not [string]::IsNullOrWhiteSpace($smtp)) { $map[$smtp.ToLowerInvariant()] = $true }
            }
        }
    }
    else {
        $members = @(Get-UnifiedGroupLinks -Identity $Identity -LinkType Members -ResultSize Unlimited -ErrorAction Stop)
        foreach ($member in $members) {
            if ($null -ne $member.PrimarySmtpAddress) {
                $smtp = $member.PrimarySmtpAddress.ToString().Trim()
                if (-not [string]::IsNullOrWhiteSpace($smtp)) { $map[$smtp.ToLowerInvariant()] = $true }
            }
            elseif ($null -ne $member.WindowsLiveID) {
                $smtp = $member.WindowsLiveID.ToString().Trim()
                if (-not [string]::IsNullOrWhiteSpace($smtp)) { $map[$smtp.ToLowerInvariant()] = $true }
            }
        }
    }

    return $map
}

function Add-MemberToTargetGroup {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('DistributionGroup','UnifiedGroup')][string]$GroupType,
        [Parameter(Mandatory = $true)][string]$Identity,
        [Parameter(Mandatory = $true)][string]$MemberSmtp
    )

    if ($GroupType -eq 'DistributionGroup') {
        Add-DistributionGroupMember -Identity $Identity -Member $MemberSmtp -ErrorAction Stop
        return
    }

    Add-UnifiedGroupLinks -Identity $Identity -LinkType Members -Links $MemberSmtp -ErrorAction Stop
}

function Resolve-RecipientByIdOrEmail {
    param(
        [string]$Id,
        [string]$Email
    )

    $recipient = @()

    if (-not [string]::IsNullOrWhiteSpace($Id)) {
        $safeId = $Id.Replace("'", "''")
        $recipient = @(Get-EXORecipient -ResultSize Unlimited -Filter "ExternalDirectoryObjectId -eq '$safeId'" -ErrorAction SilentlyContinue)
    }

    if ((-not $recipient -or @($recipient).Count -eq 0) -and -not [string]::IsNullOrWhiteSpace($Email)) {
        try {
            $recipient = @(Get-EXORecipient -Identity $Email -ErrorAction Stop)
        }
        catch {}
    }

    if ($recipient -and @($recipient).Count -gt 0) {
        return @($recipient)[0]
    }

    return $null
}

Show-Header

try {
    $InputCsv   = Read-ValidatedCsvPath -Prompt 'Ruta completa del CSV de entrada'
    $GroupEmail = Read-ValidatedEmail -Prompt 'Correo del grupo destino'

    $OutputFolder = Split-Path -Path $InputCsv -Parent
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $script:ResultCsv = Join-Path -Path $OutputFolder -ChildPath "resultado_add_group_$timestamp.csv"
    $script:LogFile   = Join-Path -Path $OutputFolder -ChildPath "resultado_add_group_$timestamp.log"

    Write-Log 'Iniciando proceso.'
    Write-Log "CSV de entrada: $InputCsv"
    Write-Log "Carpeta de salida: $OutputFolder"
    Write-Log "Correo de grupo introducido: $GroupEmail"
    Write-Log "CSV de resultados: $script:ResultCsv"

    Write-Log 'Resolviendo grupo destino por correo...'
    $group = Resolve-TargetGroup -GroupEmail $GroupEmail
    Write-Log ("Grupo validado: {0} | {1} | ID={2} | Tipo={3}" -f $group.DisplayName, $group.PrimarySmtpAddress, $group.Id, $group.GroupType) 'OK'

    Write-Log 'Leyendo CSV con acceso compartido...'
    $rows = @(Import-EmailIdCsvRobust -Path $InputCsv)
    Write-Log ("Registros detectados: {0}" -f $rows.Count)

    Write-Log 'Cargando miembros actuales del grupo...'
    $existingMemberMap = Get-ExistingMemberMap -GroupType $group.GroupType -Identity $group.Identity
    Write-Log ("Miembros actuales cargados: {0}" -f $existingMemberMap.Count)

    $results = New-Object System.Collections.Generic.List[object]
    $noAgregados = New-Object System.Collections.Generic.List[object]
    $seenIds = @{}
    $graphCache = @{}

    $totalProcesados = 0
    $agregados = 0
    $agregadosSinExchange = 0
    $yaExistian = 0
    $noResueltos = 0
    $idNoResueltos = 0
    $guidInvalidos = 0
    $duplicadosCsv = 0
    $recipientSinSmtp = 0
    $registrosVacios = 0
    $errores = 0
    $currentIndex = 0

    foreach ($row in $rows) {
        $baseEmail = if ($null -ne $row.Email) { [string]$row.Email } else { '' }
        $baseId    = if ($null -ne $row.Id) { [string]$row.Id } else { '' }

        $expandedEmails = @(Split-MultiEmailValue -Value $baseEmail)
        if ($expandedEmails.Count -eq 0) {
            $expandedEmails = @('')
        }

        foreach ($emailItem in $expandedEmails) {
            $currentIndex++
            $totalProcesados++

            $inputEmail = [string]$emailItem
            $inputId = [string]$baseId
            $email = $inputEmail.Trim()
            $id = $inputId.Trim()
            $resolutionSource = ''

            $percentComplete = if ($rows.Count -gt 0) { [math]::Round(($currentIndex / $rows.Count) * 100, 2) } else { 0 }
            Write-Progress -Activity 'Añadiendo usuarios al grupo' -Status "Procesando $currentIndex" -PercentComplete $percentComplete

            Write-Log ("Procesando InputEmail='{0}' | InputId='{1}'" -f $email, $id)

            if ([string]::IsNullOrWhiteSpace($email) -and [string]::IsNullOrWhiteSpace($id)) {
                $registrosVacios++
                $resultObj = New-ResultObject -InputEmail $inputEmail -InputId $inputId -Email $email -Id $id -Estado 'REGISTRO_VACIO' -Detalle 'Fila sin datos útiles'
                [void]$results.Add($resultObj)
                [void]$noAgregados.Add($resultObj)
                Write-Log "NO AÑADIDO | Email='$email' | Motivo=REGISTRO_VACIO" 'WARN'
                continue
            }

            if (([string]::IsNullOrWhiteSpace($id) -and -not [string]::IsNullOrWhiteSpace($email)) -or ([string]::IsNullOrWhiteSpace($email) -and -not [string]::IsNullOrWhiteSpace($id))) {
                $resolved = Resolve-GraphIdentity -Email $email -Id $id -Cache $graphCache
                $email = [string]$resolved.Email
                $id = [string]$resolved.Id
                $resolutionSource = [string]$resolved.Source
                Write-Log ("Graph enriquecido -> Email='{0}' | Id='{1}' | Source={2}" -f $email, $id, $resolutionSource)
            }

            if ([string]::IsNullOrWhiteSpace($id)) {
                $idNoResueltos++
                $resultObj = New-ResultObject -InputEmail $inputEmail -InputId $inputId -Email $email -Id $id -Estado 'ID_NO_RESUELTO' -Detalle 'No se pudo resolver el Id a través de Graph' -ResolutionSource $resolutionSource
                [void]$results.Add($resultObj)
                [void]$noAgregados.Add($resultObj)
                Write-Log "NO AÑADIDO | Email='$email' | Motivo=ID_NO_RESUELTO" 'WARN'
                continue
            }

            if (-not (Test-GuidValue -Value $id)) {
                $guidInvalidos++
                $resultObj = New-ResultObject -InputEmail $inputEmail -InputId $inputId -Email $email -Id $id -Estado 'ID_GUID_INVALIDO' -Detalle 'El valor no tiene formato GUID válido' -ResolutionSource $resolutionSource
                [void]$results.Add($resultObj)
                [void]$noAgregados.Add($resultObj)
                Write-Log "NO AÑADIDO | Email='$email' | Id='$id' | Motivo=ID_GUID_INVALIDO" 'WARN'
                continue
            }

            if ($seenIds.ContainsKey($id.ToLowerInvariant())) {
                $duplicadosCsv++
                $resultObj = New-ResultObject -InputEmail $inputEmail -InputId $inputId -Email $email -Id $id -Estado 'DUPLICADO_EN_CSV' -Detalle 'Ese Id ya apareció antes en el CSV' -ResolutionSource $resolutionSource
                [void]$results.Add($resultObj)
                [void]$noAgregados.Add($resultObj)
                Write-Log "NO AÑADIDO | Email='$email' | Id='$id' | Motivo=DUPLICADO_EN_CSV" 'WARN'
                continue
            }
            else {
                $seenIds[$id.ToLowerInvariant()] = $true
            }

            try {
                $recipientObj = Resolve-RecipientByIdOrEmail -Id $id -Email $email

                if ($null -eq $recipientObj) {
                    # El usuario tiene ID válido en Graph pero no resuelve en Exchange (típicamente sin licencia).
                    # Se intenta la adición directa por email de Graph (o Id como último recurso).
                    $fallbackIdentity = if (-not [string]::IsNullOrWhiteSpace($email)) { $email } else { $id }
                    $fallbackKey = $fallbackIdentity.ToLowerInvariant()

                    if ($existingMemberMap.ContainsKey($fallbackKey)) {
                        $yaExistian++
                        $resultObj = New-ResultObject -InputEmail $inputEmail -InputId $inputId -Email $email -Id $id -Estado 'YA_EXISTE_EN_GRUPO' -Detalle 'El usuario ya pertenece al grupo (verificado por email/id, sin recipient en Exchange)' -ResolutionSource $resolutionSource
                        [void]$results.Add($resultObj)
                        [void]$noAgregados.Add($resultObj)
                        Write-Log "NO AÑADIDO | Email='$email' | Id='$id' | Motivo=YA_EXISTE_EN_GRUPO" 'WARN'
                        continue
                    }

                    try {
                        Add-MemberToTargetGroup -GroupType $group.GroupType -Identity $group.Identity -MemberSmtp $fallbackIdentity
                        $existingMemberMap[$fallbackKey] = $true
                        $agregados++
                        $agregadosSinExchange++
                        $resultObj = New-ResultObject -InputEmail $inputEmail -InputId $inputId -Email $email -Id $id -Estado 'AGREGADO_SIN_EXCHANGE' -Detalle 'Añadido directamente por email/id de Graph sin recipient en Exchange (posiblemente sin licencia)' -ResolutionSource $resolutionSource
                        [void]$results.Add($resultObj)
                        Write-Log "AÑADIDO (SIN EXCHANGE) | Email='$email' | Id='$id' | Identity='$fallbackIdentity'" 'OK'
                    }
                    catch {
                        $noResueltos++
                        $errorFallback = $_.Exception.Message
                        $resultObj = New-ResultObject -InputEmail $inputEmail -InputId $inputId -Email $email -Id $id -Estado 'NO_RESUELTO_EN_EXCHANGE' -Detalle "No resuelto en Exchange y el intento directo también falló: $errorFallback" -ResolutionSource $resolutionSource
                        [void]$results.Add($resultObj)
                        [void]$noAgregados.Add($resultObj)
                        Write-Log "NO AÑADIDO | Email='$email' | Id='$id' | Motivo=NO_RESUELTO_EN_EXCHANGE" 'WARN'
                    }
                    continue
                }

                $recipientAddress = ''
                if ($null -ne $recipientObj.PrimarySmtpAddress) {
                    $recipientAddress = $recipientObj.PrimarySmtpAddress.ToString().Trim()
                }

                if ([string]::IsNullOrWhiteSpace($recipientAddress)) {
                    # Recipient existe en Exchange pero sin PrimarySmtpAddress.
                    # Se intenta añadir por email de Graph como fallback.
                    $fallbackIdentity = if (-not [string]::IsNullOrWhiteSpace($email)) { $email } else { $id }
                    $fallbackKey = $fallbackIdentity.ToLowerInvariant()

                    if ($existingMemberMap.ContainsKey($fallbackKey)) {
                        $yaExistian++
                        $resultObj = New-ResultObject -InputEmail $inputEmail -InputId $inputId -Email $email -Id $id -Estado 'YA_EXISTE_EN_GRUPO' -Detalle 'El usuario ya pertenece al grupo (verificado por email/id, sin SMTP en Exchange)' -RecipientType ([string]$recipientObj.RecipientTypeDetails) -DisplayName ([string]$recipientObj.DisplayName) -ResolutionSource $resolutionSource
                        [void]$results.Add($resultObj)
                        [void]$noAgregados.Add($resultObj)
                        Write-Log "NO AÑADIDO | Email='$email' | Id='$id' | Motivo=YA_EXISTE_EN_GRUPO" 'WARN'
                        continue
                    }

                    try {
                        Add-MemberToTargetGroup -GroupType $group.GroupType -Identity $group.Identity -MemberSmtp $fallbackIdentity
                        $existingMemberMap[$fallbackKey] = $true
                        $agregados++
                        $agregadosSinExchange++
                        $resultObj = New-ResultObject -InputEmail $inputEmail -InputId $inputId -Email $email -Id $id -Estado 'AGREGADO_SIN_SMTP' -Detalle 'Añadido por email/id de Graph sin PrimarySmtpAddress en Exchange (posiblemente sin licencia)' -RecipientType ([string]$recipientObj.RecipientTypeDetails) -DisplayName ([string]$recipientObj.DisplayName) -ResolutionSource $resolutionSource
                        [void]$results.Add($resultObj)
                        Write-Log "AÑADIDO (SIN SMTP) | Email='$email' | Id='$id' | Identity='$fallbackIdentity'" 'OK'
                    }
                    catch {
                        $recipientSinSmtp++
                        $errorFallback = $_.Exception.Message
                        $resultObj = New-ResultObject -InputEmail $inputEmail -InputId $inputId -Email $email -Id $id -Estado 'RECIPIENT_SIN_SMTP' -Detalle "El recipient existe pero no tiene SMTP y el intento directo también falló: $errorFallback" -RecipientType ([string]$recipientObj.RecipientTypeDetails) -DisplayName ([string]$recipientObj.DisplayName) -ResolutionSource $resolutionSource
                        [void]$results.Add($resultObj)
                        [void]$noAgregados.Add($resultObj)
                        Write-Log "NO AÑADIDO | Email='$email' | Id='$id' | Motivo=RECIPIENT_SIN_SMTP" 'WARN'
                    }
                    continue
                }

                $recipientKey = $recipientAddress.ToLowerInvariant()
                if ($existingMemberMap.ContainsKey($recipientKey)) {
                    $yaExistian++
                    $resultObj = New-ResultObject -InputEmail $inputEmail -InputId $inputId -Email $email -Id $id -Estado 'YA_EXISTE_EN_GRUPO' -Detalle 'El usuario ya pertenece al grupo' -RecipientType ([string]$recipientObj.RecipientTypeDetails) -PrimarySmtpAddress $recipientAddress -DisplayName ([string]$recipientObj.DisplayName) -ResolutionSource $resolutionSource
                    [void]$results.Add($resultObj)
                    [void]$noAgregados.Add($resultObj)
                    Write-Log "NO AÑADIDO | Email='$email' | SMTP='$recipientAddress' | Motivo=YA_EXISTE_EN_GRUPO" 'WARN'
                    continue
                }

                Add-MemberToTargetGroup -GroupType $group.GroupType -Identity $group.Identity -MemberSmtp $recipientAddress
                $existingMemberMap[$recipientKey] = $true

                $agregados++
                $resultObj = New-ResultObject -InputEmail $inputEmail -InputId $inputId -Email $email -Id $id -Estado 'AGREGADO' -Detalle 'Añadido correctamente al grupo' -RecipientType ([string]$recipientObj.RecipientTypeDetails) -PrimarySmtpAddress $recipientAddress -DisplayName ([string]$recipientObj.DisplayName) -ResolutionSource $resolutionSource
                [void]$results.Add($resultObj)
                Write-Log "AÑADIDO | Email='$email' | Id='$id' | SMTP='$recipientAddress'" 'OK'
            }
            catch {
                $errores++
                $errorMessage = $_.Exception.Message
                $resultObj = New-ResultObject -InputEmail $inputEmail -InputId $inputId -Email $email -Id $id -Estado 'ERROR_AL_AGREGAR' -Detalle $errorMessage -ResolutionSource $resolutionSource
                [void]$results.Add($resultObj)
                [void]$noAgregados.Add($resultObj)
                Write-Log "NO AÑADIDO | Email='$email' | Id='$id' | Motivo=ERROR_AL_AGREGAR | Error=$errorMessage" 'ERROR'
            }
        }
    }

    Write-Progress -Activity 'Añadiendo usuarios al grupo' -Completed

    if ($results.Count -gt 0) {
        $results | Export-Csv -Path $script:ResultCsv -NoTypeInformation -Encoding UTF8
        Write-Log 'CSV de resultados exportado correctamente.' 'OK'
    }

    $successWithoutErrors = ($agregados -gt 0 -and $errores -eq 0)
    $script:EnableFileLogging = $successWithoutErrors

    if ($successWithoutErrors) {
        Write-Log ("Log persistido porque hubo éxito real: {0} usuario(s) añadido(s) ({1} sin Exchange/SMTP) y 0 errores." -f $agregados, $agregadosSinExchange) 'OK'
        Save-SuccessLog
    }

    Write-Host ''
    Write-Host '================ RESUMEN FINAL ================' -ForegroundColor White
    Write-Host ("Grupo:             {0} <{1}>" -f $group.DisplayName, $group.PrimarySmtpAddress) -ForegroundColor Cyan
    Write-Host ("Tipo grupo:        {0}" -f $group.GroupType) -ForegroundColor Cyan
    Write-Host ("ID grupo:          {0}" -f $group.Id) -ForegroundColor Cyan
    Write-Host ("Total procesados:  {0}" -f $totalProcesados) -ForegroundColor White
    Write-Host ("Añadidos:          {0}" -f $agregados) -ForegroundColor Green
    Write-Host ("  Sin Exchange/SMTP: {0}" -f $agregadosSinExchange) -ForegroundColor DarkYellow
    Write-Host ("Ya existían:       {0}" -f $yaExistian) -ForegroundColor Yellow
    Write-Host ("No resueltos EXO:  {0}" -f $noResueltos) -ForegroundColor Yellow
    Write-Host ("ID no resueltos:   {0}" -f $idNoResueltos) -ForegroundColor Yellow
    Write-Host ("GUID inválidos:    {0}" -f $guidInvalidos) -ForegroundColor Yellow
    Write-Host ("Duplicados CSV:    {0}" -f $duplicadosCsv) -ForegroundColor Yellow
    Write-Host ("Sin SMTP:          {0}" -f $recipientSinSmtp) -ForegroundColor Yellow
    Write-Host ("Registros vacíos:  {0}" -f $registrosVacios) -ForegroundColor Yellow
    Write-Host ("Errores:           {0}" -f $errores) -ForegroundColor Red
    Write-Host ("No añadidos total: {0}" -f $noAgregados.Count) -ForegroundColor Yellow
    Write-Host ("CSV resultado:     {0}" -f $script:ResultCsv) -ForegroundColor Cyan

    if ($successWithoutErrors) {
        Write-Host ("Log:               {0}" -f $script:LogFile) -ForegroundColor Cyan
    }
    else {
        Write-Host 'Log:               no generado' -ForegroundColor DarkGray
    }

    Write-Host '===============================================' -ForegroundColor White
    Write-Host ''

    if ($noAgregados.Count -gt 0) {
        Write-Host '============= USUARIOS NO AÑADIDOS =============' -ForegroundColor Yellow
        foreach ($item in $noAgregados) {
            Write-Host ("InputEmail: {0} | InputId: {1} | ResolvedEmail: {2} | ResolvedId: {3} | Estado: {4}" -f $item.InputEmail, $item.InputId, $item.ResolvedEmail, $item.ResolvedId, $item.Estado) -ForegroundColor Yellow
        }
        Write-Host '===============================================' -ForegroundColor Yellow
        Write-Host ''
    }
}
catch {
    $fatalMessage = $_.Exception.Message
    Write-Host ''
    Write-Host "[FATAL] $fatalMessage" -ForegroundColor Red
    Write-Host ''
    throw
}