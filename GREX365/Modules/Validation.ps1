# Validation module
# Input validation and normalization. Single source of truth.

function Normalize-Input {
    param([AllowNull()][string]$Value)

    if ($null -eq $Value) { return '' }
    $v = $Value.Trim()
    if ($v.Length -ge 2) {
        $first = $v.Substring(0, 1)
        $last  = $v.Substring($v.Length - 1, 1)
        if ((($first -eq '"') -and ($last -eq '"')) -or (($first -eq "'") -and ($last -eq "'"))) {
            $v = $v.Substring(1, $v.Length - 2).Trim()
        }
    }
    return $v
}

function Test-Email {
    param([AllowNull()][string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    return ($Value.Trim() -match '^[^@\s]+@[^@\s]+\.[^@\s]+$')
}

function Test-Guid {
    param([AllowNull()][string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    $g = [guid]::Empty
    return [guid]::TryParse($Value.Trim(), [ref]$g)
}

function Test-Domain {
    param([AllowNull()][string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    return ($Value.Trim() -match '^[a-zA-Z0-9]([a-zA-Z0-9\-]{0,61}[a-zA-Z0-9])?(\.[a-zA-Z0-9]([a-zA-Z0-9\-]{0,61}[a-zA-Z0-9])?)+$')
}

function Test-CsvPathFormat {
    param([Parameter(Mandatory = $true)][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    if ($Path.Length -lt 5) { return $false }
    if ($Path -notmatch '(?i)\.csv$') { return $false }
    if (-not [System.IO.Path]::IsPathRooted($Path)) { return $false }

    try {
        $invalid = [System.IO.Path]::GetInvalidPathChars()
        foreach ($c in $invalid) {
            if ($Path.Contains([string]$c)) { return $false }
        }
    } catch { return $false }
    return $true
}

function Read-ValidatedCsvPath {
    param([Parameter(Mandatory = $true)][string]$Prompt)

    Write-Indent
    Write-Host $Prompt -ForegroundColor White
    if (Get-Command Show-CsvFormatHint -ErrorAction SilentlyContinue) {
        Show-CsvFormatHint
    }
    $raw = Read-Input -Prompt 'Ruta CSV'
    $value = Normalize-Input -Value $raw

    if ([string]::IsNullOrWhiteSpace($value)) {
        throw 'La ruta del CSV está vacía. Operación cancelada.'
    }
    if (-not (Test-CsvPathFormat -Path $value)) {
        throw 'La ruta indicada no tiene formato válido de CSV. Ejemplo: C:\Temp\fichero.csv'
    }
    if (-not (Test-Path -LiteralPath $value -PathType Leaf)) {
        throw ('No existe el CSV de entrada: ' + $value)
    }
    return (Resolve-Path -LiteralPath $value).Path
}

function Read-ValidatedFolder {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [string]$Default = ''
    )

    $raw = Read-Input -Prompt $Prompt -Default $Default
    $value = Normalize-Input -Value $raw
    if ([string]::IsNullOrWhiteSpace($value)) {
        throw 'Ruta de carpeta vacía. Operación cancelada.'
    }
    if (-not (Test-Path -LiteralPath $value)) {
        try { New-Item -ItemType Directory -Path $value -Force | Out-Null }
        catch { throw ('No se pudo crear la carpeta: ' + $value) }
    }
    return (Resolve-Path -LiteralPath $value).Path
}

function Read-ValidatedEmail {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [switch]$AllowEmpty
    )

    while ($true) {
        $raw = Read-Input -Prompt $Prompt
        $value = Normalize-Input -Value $raw

        if ([string]::IsNullOrWhiteSpace($value)) {
            if ($AllowEmpty) { return $null }
            Show-WarningBlock -Title 'Email requerido'
            continue
        }
        if (-not (Test-Email -Value $value)) {
            Show-WarningBlock -Title 'Formato de email inválido' -Detail 'Esperado: usuario@dominio.com'
            continue
        }
        return $value.ToLowerInvariant()
    }
}
