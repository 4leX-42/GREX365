# CSV module
# Robust CSV reader: shared file access, encoding detection, delimiter detection,
# flexible header mapping. Single source of truth for all CSV parsing in the toolkit.

function Open-SharedFileStream {
    param([Parameter(Mandatory = $true)][string]$Path)
    return [System.IO.File]::Open(
        $Path,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::Read,
        [System.IO.FileShare]::ReadWrite
    )
}

function Get-FileEncoding {
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
    } finally {
        if ($null -ne $fs) { $fs.Dispose() }
    }
}

function Get-CsvDelimiter {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][System.Text.Encoding]$Encoding
    )

    $fs = $null; $reader = $null
    try {
        $fs = Open-SharedFileStream -Path $Path
        $reader = New-Object System.IO.StreamReader($fs, $Encoding, $true)
        $first = $reader.ReadLine()
    } catch {
        return ';'
    } finally {
        if ($null -ne $reader) { $reader.Dispose() }
        elseif ($null -ne $fs) { $fs.Dispose() }
    }
    if ([string]::IsNullOrWhiteSpace($first)) { return ';' }

    $counts = @{ ';' = 0; ',' = 0; "`t" = 0 }
    $inQuote = $false
    foreach ($ch in $first.ToCharArray()) {
        if ($ch -eq '"') { $inQuote = -not $inQuote; continue }
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
                [void]$sb.Append('"'); $i += 2; continue
            }
            $inQuotes = -not $inQuotes; $i++; continue
        }
        if (([string]$ch) -eq $Delimiter -and -not $inQuotes) {
            [void]$fields.Add($sb.ToString()); [void]$sb.Clear(); $i++; continue
        }
        [void]$sb.Append($ch); $i++
    }
    [void]$fields.Add($sb.ToString())
    return $fields.ToArray()
}

function Read-CsvLines {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][System.Text.Encoding]$Encoding
    )

    $lines = New-Object System.Collections.Generic.List[string]
    $fs = $null; $reader = $null
    try {
        $fs = Open-SharedFileStream -Path $Path
        $reader = New-Object System.IO.StreamReader($fs, $Encoding, $true)
        while (-not $reader.EndOfStream) { [void]$lines.Add($reader.ReadLine()) }
    } finally {
        if ($null -ne $reader) { $reader.Dispose() }
        elseif ($null -ne $fs) { $fs.Dispose() }
    }
    return $lines.ToArray()
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

# Schema → candidate header names.
$global:GREX365_CsvSchemas = @{
    'EmailId' = @{
        Email = @('Email','Mail','UserPrincipalName','UPN','Correo','PrimarySmtpAddress','WindowsLiveID','Login','EmailAddress')
        Id    = @('Id','ID','ObjectId','ExternalDirectoryObjectId','Guid','GUID','EntraId','AzureObjectId')
    }
    'EmailGroupName' = @{
        Email     = @('Email','Mail','UserPrincipalName','UPN','Correo','PrimarySmtpAddress','WindowsLiveID','Login','EmailAddress','MemberEmail')
        GroupName = @('GroupName','Group','Grupo','NombreGrupo','ListName','DL','DistributionList','GroupAlias','Nombre')
    }
}

function Import-FlexibleCsv {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][ValidateSet('EmailId','EmailGroupName')][string]$Schema
    )

    $encoding = Get-FileEncoding -Path $Path
    $delimiter = Get-CsvDelimiter -Path $Path -Encoding $encoding
    $delimDisplay = if ($delimiter -eq "`t") { 'TAB' } else { $delimiter }
    Write-Log ("Encoding={0}  Delimitador='{1}'" -f $encoding.EncodingName, $delimDisplay) -Level DEBUG -Source 'CSV'

    $lines = @(Read-CsvLines -Path $Path -Encoding $encoding)
    $nonEmpty = @($lines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($nonEmpty.Count -eq 0) { throw 'El CSV está vacío.' }

    $hFields = @(Split-CsvLine -Line $nonEmpty[0] -Delimiter $delimiter)
    if ($hFields.Count -gt 0 -and $null -ne $hFields[0]) {
        $hFields[0] = ([string]$hFields[0]).TrimStart([char]0xFEFF).Trim()
    }
    $hFields = @($hFields | ForEach-Object { if ($null -eq $_) { '' } else { ([string]$_).Trim() } })

    $schemaDef = $global:GREX365_CsvSchemas[$Schema]
    $columnIndices = @{}
    foreach ($key in $schemaDef.Keys) {
        $columnIndices[$key] = Get-ColumnIndex -Header $hFields -CandidateNames $schemaDef[$key]
    }

    $allMissing = $true
    foreach ($idx in $columnIndices.Values) {
        if ($idx -ge 0) { $allMissing = $false; break }
    }

    if ($allMissing) {
        if ($hFields.Count -ge 2) {
            $keys = @($schemaDef.Keys | Sort-Object)
            $columnIndices[$keys[0]] = 0
            $columnIndices[$keys[1]] = 1
            Write-Log ("Cabecera no estándar — col 1={0}, col 2={1}" -f $keys[0], $keys[1]) -Level WARN -Source 'CSV'
        }
        elseif ($hFields.Count -eq 1) {
            $keys = @($schemaDef.Keys | Sort-Object)
            $columnIndices[$keys[0]] = 0
            Write-Log ("Solo una columna — usada como {0}" -f $keys[0]) -Level WARN -Source 'CSV'
        }
        else { throw 'No se pudo mapear la cabecera del CSV.' }
    }

    Write-Log ('Cabeceras: ' + ($hFields -join ' | ')) -Level OK -Source 'CSV'

    $rows = New-Object System.Collections.Generic.List[object]
    for ($li = 1; $li -lt $nonEmpty.Count; $li++) {
        $line = [string]$nonEmpty[$li]
        $fields = @(Split-CsvLine -Line $line -Delimiter $delimiter)
        if ($fields.Count -eq 0) { continue }

        $row = [ordered]@{}
        foreach ($key in $columnIndices.Keys) {
            $idx = $columnIndices[$key]
            $val = ''
            if ($idx -ge 0 -and $idx -lt $fields.Count) {
                $val = ([string]$fields[$idx]).Trim()
            }
            $row[$key] = $val
        }
        $row['RawText'] = $line
        [void]$rows.Add([PSCustomObject]$row)
    }

    if ($Schema -eq 'EmailId') {
        foreach ($r in $rows) {
            if ([string]::IsNullOrWhiteSpace($r.Email) -and [string]::IsNullOrWhiteSpace($r.Id)) {
                $candidates = @(Split-CsvLine -Line $r.RawText -Delimiter $delimiter)
                foreach ($field in $candidates) {
                    $value = ([string]$field).Trim()
                    if ([string]::IsNullOrWhiteSpace($value)) { continue }
                    if ([string]::IsNullOrWhiteSpace($r.Email) -and (Test-Email -Value $value)) { $r.Email = $value; continue }
                    if ([string]::IsNullOrWhiteSpace($r.Id) -and (Test-Guid -Value $value))   { $r.Id = $value; continue }
                }
            }
        }
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
