#requires -Version 5.1
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [string[]]$FolderPaths,
    [string]$Domain
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Host.UI.RawUI.WindowTitle = 'Creación de grupos / DL desde CSV | GREX365'

# --- ESTADO ---

$script:LogRows    = New-Object System.Collections.Generic.List[object]
$script:GroupType  = ''
$script:Domain     = ''
$script:WhatIfMode = $WhatIfPreference -ne [System.Management.Automation.ActionPreference]::SilentlyContinue

# --- LOGGING ---

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')][string]$Level = 'INFO'
    )

    $time = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$time][$Level] $Message"
    $color = switch ($Level) { 'OK' { 'Green' } 'WARN' { 'Yellow' } 'ERROR' { 'Red' } default { 'Cyan' } }
    Write-Host $line -ForegroundColor $color
}

function Add-LogRow {
    param(
        [string]$CsvFile = '', [string]$GroupName = '', [string]$GroupEmail = '',
        [string]$Action = '', [string]$UserEmail = '', [string]$Detail = ''
    )
    [void]$script:LogRows.Add([PSCustomObject]@{
        Timestamp  = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        CsvFile    = $CsvFile
        GroupName  = $GroupName
        GroupEmail = $GroupEmail
        Action     = $Action
        UserEmail  = $UserEmail
        Detail     = $Detail
    })
}

# --- VALIDACIÓN DE CONEXIÓN ---

function Assert-RequiredServicesReady {
    $exoOk = $false; $mgOk = $false
    try {
        if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
            $s = Get-ConnectionInformation -ErrorAction SilentlyContinue
            if ($s | Where-Object { $_.State -eq 'Connected' }) { $exoOk = $true }
        }
    } catch {}
    try {
        if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
            $ctx = Get-MgContext -ErrorAction SilentlyContinue
            if ($ctx) {
                $isAppOnly = ([string]$ctx.AuthType -match 'AppOnly')
                if ($isAppOnly -and $ctx.ClientId -and $ctx.TenantId) { $mgOk = $true }
                elseif ($ctx.Account) { $mgOk = $true }
            }
        }
    } catch {}
    if (-not $exoOk -or -not $mgOk) {
        throw "Faltan servicios M365 (EXO=$exoOk, Graph=$mgOk). Ejecuta desde Main.ps1."
    }
}

# --- CABECERA ---

function Get-CenterPadding {
    param([Parameter(Mandatory = $true)][string]$Text)
    $w = [Console]::WindowWidth
    $pad = [Math]::Floor(($w - $Text.Length) / 2)
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
    Write-Centered 'CREACIÓN MASIVA DE GRUPOS / DL DESDE CSV' -Color Green
    Write-Centered '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━' -Color Cyan
    Write-Centered 'M365 Groups (Graph)  //  Distribution Lists (ExO)' -Color DarkCyan
    Write-Host ''
}

# --- ENTRADAS ---

function Normalize-QuotedInput {
    param([AllowNull()][string]$Value)
    if ($null -eq $Value) { return '' }
    $v = $Value.Trim()
    if ($v.Length -ge 2) {
        $f = $v.Substring(0, 1); $l = $v.Substring($v.Length - 1, 1)
        if ((($f -eq '"') -and ($l -eq '"')) -or (($f -eq "'") -and ($l -eq "'"))) {
            $v = $v.Substring(1, $v.Length - 2).Trim()
        }
    }
    return $v
}

function Test-EmailValue {
    param([AllowNull()][string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    return ($Value.Trim() -match '^[^@\s]+@[^@\s]+\.[^@\s]+$')
}

# --- MOTOR DE CSV ---

function Open-SharedFileStream {
    param([Parameter(Mandatory = $true)][string]$Path)
    return [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
}

function Get-FileEncodingObject {
    param([Parameter(Mandatory = $true)][string]$Path)
    $fs = $null
    try {
        $fs = Open-SharedFileStream -Path $Path
        $buf = New-Object byte[] 4
        $read = $fs.Read($buf, 0, 4)
        if ($read -ge 2 -and $buf[0] -eq 0xFF -and $buf[1] -eq 0xFE) { return [System.Text.Encoding]::Unicode }
        if ($read -ge 2 -and $buf[0] -eq 0xFE -and $buf[1] -eq 0xFF) { return [System.Text.Encoding]::BigEndianUnicode }
        if ($read -ge 3 -and $buf[0] -eq 0xEF -and $buf[1] -eq 0xBB -and $buf[2] -eq 0xBF) { return [System.Text.Encoding]::UTF8 }
        return [System.Text.Encoding]::UTF8
    }
    finally { if ($null -ne $fs) { $fs.Dispose() } }
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
    }
    catch { return ';' }
    finally {
        if ($null -ne $reader) { $reader.Dispose() }
        elseif ($null -ne $fs) { $fs.Dispose() }
    }
    if ([string]::IsNullOrWhiteSpace($first)) { return ';' }

    $counts = @{ ';' = 0; ',' = 0; "`t" = 0 }
    $inQ = $false
    foreach ($ch in $first.ToCharArray()) {
        if ($ch -eq '"') { $inQ = -not $inQ; continue }
        if ($inQ) { continue }
        $k = [string]$ch
        if ($counts.ContainsKey($k)) { $counts[$k]++ }
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
    $inQ = $false; $i = 0
    while ($i -lt $Line.Length) {
        $ch = $Line[$i]
        if ($ch -eq '"') {
            if ($inQ -and ($i + 1) -lt $Line.Length -and $Line[$i + 1] -eq '"') {
                [void]$sb.Append('"'); $i += 2; continue
            }
            $inQ = -not $inQ; $i++; continue
        }
        if (([string]$ch) -eq $Delimiter -and -not $inQ) {
            [void]$fields.Add($sb.ToString()); [void]$sb.Clear(); $i++; continue
        }
        [void]$sb.Append($ch); $i++
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
    $fs = $null; $reader = $null
    try {
        $fs = Open-SharedFileStream -Path $Path
        $reader = New-Object System.IO.StreamReader($fs, $Encoding, $true)
        while (-not $reader.EndOfStream) { [void]$lines.Add($reader.ReadLine()) }
    }
    finally {
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
        foreach ($c in $CandidateNames) {
            if ($name.Trim().ToLowerInvariant() -eq $c.Trim().ToLowerInvariant()) { return $i }
        }
    }
    return -1
}

$script:EmailCandidates = @('Email','Mail','UserPrincipalName','UPN','Correo','PrimarySmtpAddress','WindowsLiveID','Login','EmailAddress','MemberEmail')
$script:GroupCandidates = @('GroupName','Group','Grupo','NombreGrupo','ListName','DL','DistributionList','GroupAlias','Nombre')

function Import-GroupCsvRobust {
    param([Parameter(Mandatory = $true)][string]$Path)

    $encoding = Get-FileEncodingObject -Path $Path
    Write-Log "  Encoding: $($encoding.EncodingName)"

    $delimiter = Get-CsvDelimiter -Path $Path -Encoding $encoding
    $delimDisplay = if ($delimiter -eq "`t") { 'TAB' } else { $delimiter }
    Write-Log "  Delimitador: '$delimDisplay'"

    $lines = @(Get-SharedTextLines -Path $Path -Encoding $encoding)
    $nonEmpty = @($lines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($nonEmpty.Count -eq 0) { throw "CSV vacío: $Path" }

    $hFields = @(Split-CsvLine -Line $nonEmpty[0] -Delimiter $delimiter)
    if ($hFields.Count -gt 0 -and $null -ne $hFields[0]) {
        $hFields[0] = ([string]$hFields[0]).TrimStart([char]0xFEFF).Trim()
    }
    $hFields = @($hFields | ForEach-Object { if ($null -eq $_) { '' } else { ([string]$_).Trim() } })
    Write-Log "  Cabeceras: $($hFields -join ' | ')"

    $emailIdx = Get-ColumnIndex -Header $hFields -CandidateNames $script:EmailCandidates
    $groupIdx = Get-ColumnIndex -Header $hFields -CandidateNames $script:GroupCandidates

    if ($emailIdx -lt 0 -and $groupIdx -lt 0) {
        if ($hFields.Count -ge 2) {
            $emailIdx = 0; $groupIdx = 1
            Write-Log '  Cabecera no estándar → col 1=Email, col 2=GroupName' 'WARN'
        }
        elseif ($hFields.Count -eq 1) {
            $emailIdx = 0; $groupIdx = -1
            Write-Log '  Solo una columna → se usará como Email.' 'WARN'
        }
        else { throw "No se pudo mapear la cabecera del CSV: $Path" }
    }

    $emailLabel = if ($emailIdx -ge 0 -and $emailIdx -lt $hFields.Count) { $hFields[$emailIdx] } else { '(ninguna)' }
    $groupLabel = if ($groupIdx -ge 0 -and $groupIdx -lt $hFields.Count) { $hFields[$groupIdx] } else { '(ninguna)' }
    Write-Log "  Mapa columnas → Email='$emailLabel' | GroupName='$groupLabel'" 'OK'

    $rows = New-Object System.Collections.Generic.List[object]
    for ($li = 1; $li -lt $nonEmpty.Count; $li++) {
        $line = [string]$nonEmpty[$li]
        $fields = @(Split-CsvLine -Line $line -Delimiter $delimiter)
        if ($fields.Count -eq 0) { continue }

        $email = ''
        $group = ''
        if ($emailIdx -ge 0 -and $emailIdx -lt $fields.Count) { $email = ([string]$fields[$emailIdx]).Trim() }
        if ($groupIdx -ge 0 -and $groupIdx -lt $fields.Count) { $group = ([string]$fields[$groupIdx]).Trim() }

        if ([string]::IsNullOrWhiteSpace($group) -and $fields.Count -ge 2) {
            $altIdx = if ($emailIdx -eq 0) { 1 } else { 0 }
            $group = ([string]$fields[$altIdx]).Trim()
        }

        [void]$rows.Add([PSCustomObject]@{
            Email     = $email
            GroupName = $group
            RawText   = $line
        })
    }
    return $rows.ToArray()
}

# --- OPERACIONES M365 (Graph) ---

function Get-OrCreate-M365Group {
    param([string]$DisplayName, [string]$GroupEmail, [string]$CsvFile)

    $existing = @(Get-MgGroup -Filter "mail eq '$GroupEmail'" -ErrorAction SilentlyContinue)
    if ($existing.Count -gt 0) {
        Write-Log "  [SKIP] Grupo M365 ya existe: $GroupEmail" 'WARN'
        Add-LogRow -CsvFile $CsvFile -GroupName $DisplayName -GroupEmail $GroupEmail -Action 'Skipped' -Detail 'El grupo ya existía'
        return $existing[0]
    }

    if ($script:WhatIfMode) {
        Write-Log "  [WhatIf] Se crearía grupo M365: $GroupEmail" 'WARN'
        Add-LogRow -CsvFile $CsvFile -GroupName $DisplayName -GroupEmail $GroupEmail -Action 'WhatIf-Created'
        return $null
    }

    if ($PSCmdlet.ShouldProcess($GroupEmail, 'Crear Grupo M365')) {
        $alias = ($GroupEmail -split '@')[0]
        $body = @{
            DisplayName     = $DisplayName
            MailNickname    = $alias
            MailEnabled     = $true
            SecurityEnabled = $false
            GroupTypes      = @('Unified')
            Visibility      = 'Private'
        }
        $mgGroup = New-MgGroup -BodyParameter $body -ErrorAction Stop

        Write-Log "  [CREATED] Grupo M365 creado: $GroupEmail" 'OK'
        Add-LogRow -CsvFile $CsvFile -GroupName $DisplayName -GroupEmail $GroupEmail -Action 'Created'
        return $mgGroup
    }
    return $null
}

function Get-M365GroupMembers {
    param([string]$GroupId)

    $members = @(Get-MgGroupMember -GroupId $GroupId -All -ErrorAction SilentlyContinue)
    $map = @{}
    foreach ($m in $members) {
        try {
            $u = Get-MgUser -UserId $m.Id -Property Mail,UserPrincipalName -ErrorAction SilentlyContinue
            if ($null -ne $u) {
                foreach ($addr in @($u.Mail, $u.UserPrincipalName)) {
                    if (-not [string]::IsNullOrWhiteSpace($addr)) { $map[$addr.ToLowerInvariant()] = $m.Id }
                }
            }
        } catch {}
        $map[$m.Id.ToLowerInvariant()] = $m.Id
    }
    return $map
}

function Add-M365Member {
    param(
        [string]$GroupId, [string]$UserEmail, [string]$GroupEmail,
        [string]$GroupName, [string]$CsvFile, [hashtable]$MemberMap
    )

    $key = $UserEmail.ToLowerInvariant()
    if ($MemberMap.ContainsKey($key)) {
        Write-Log "      [SKIP] Ya es miembro: $UserEmail" 'WARN'
        Add-LogRow -CsvFile $CsvFile -GroupName $GroupName -GroupEmail $GroupEmail -Action 'MemberSkipped' -UserEmail $UserEmail -Detail 'Ya pertenece al grupo'
        return
    }

    $mgUsers = @(Get-MgUser -Filter "mail eq '$UserEmail' or userPrincipalName eq '$UserEmail'" `
                            -Property Id,Mail,UserPrincipalName -ConsistencyLevel eventual -All -ErrorAction SilentlyContinue)
    if ($mgUsers.Count -eq 0) {
        Write-Log "      [ERROR] Usuario no encontrado en Graph: $UserEmail" 'ERROR'
        Add-LogRow -CsvFile $CsvFile -GroupName $GroupName -GroupEmail $GroupEmail -Action 'Error' -UserEmail $UserEmail -Detail 'Usuario no encontrado en Graph'
        return
    }

    $mgUser = $mgUsers[0]
    if ($script:WhatIfMode) {
        Write-Log "      [WhatIf] Se añadiría: $UserEmail" 'WARN'
        Add-LogRow -CsvFile $CsvFile -GroupName $GroupName -GroupEmail $GroupEmail -Action 'WhatIf-MemberAdded' -UserEmail $UserEmail
        return
    }

    if ($PSCmdlet.ShouldProcess($UserEmail, "Añadir al grupo M365 $GroupEmail")) {
        New-MgGroupMember -GroupId $GroupId -DirectoryObjectId $mgUser.Id -ErrorAction Stop
        $MemberMap[$key] = $mgUser.Id
        Write-Log "      [OK] Añadido: $UserEmail" 'OK'
        Add-LogRow -CsvFile $CsvFile -GroupName $GroupName -GroupEmail $GroupEmail -Action 'MemberAdded' -UserEmail $UserEmail
    }
}

# --- OPERACIONES DL (Exchange Online) ---

function Get-OrCreate-DL {
    param([string]$DisplayName, [string]$GroupEmail, [string]$CsvFile)

    $existing = @(Get-DistributionGroup -Identity $GroupEmail -ErrorAction SilentlyContinue)
    if ($existing.Count -gt 0) {
        Write-Log "  [SKIP] DL ya existe: $GroupEmail" 'WARN'
        Add-LogRow -CsvFile $CsvFile -GroupName $DisplayName -GroupEmail $GroupEmail -Action 'Skipped' -Detail 'La DL ya existía'
        return $existing[0]
    }

    if ($script:WhatIfMode) {
        Write-Log "  [WhatIf] Se crearía DL: $GroupEmail" 'WARN'
        Add-LogRow -CsvFile $CsvFile -GroupName $DisplayName -GroupEmail $GroupEmail -Action 'WhatIf-Created'
        return $null
    }

    if ($PSCmdlet.ShouldProcess($GroupEmail, 'Crear Lista de Distribución')) {
        $dl = New-DistributionGroup `
            -Name $DisplayName -PrimarySmtpAddress $GroupEmail `
            -Type Distribution -ErrorAction Stop

        Write-Log "  [CREATED] DL creada: $GroupEmail" 'OK'
        Add-LogRow -CsvFile $CsvFile -GroupName $DisplayName -GroupEmail $GroupEmail -Action 'Created'

        Write-Log '  Esperando propagación en Exchange (3s)...' 'INFO'
        Start-Sleep -Seconds 3
        return $dl
    }
    return $null
}

function Get-DLMemberMap {
    param([string]$Identity)
    $map = @{}
    $members = @(Get-DistributionGroupMember -Identity $Identity -ResultSize Unlimited -ErrorAction SilentlyContinue)
    foreach ($m in $members) {
        if ($null -ne $m.PrimarySmtpAddress) {
            $smtp = $m.PrimarySmtpAddress.ToString().Trim()
            if (-not [string]::IsNullOrWhiteSpace($smtp)) { $map[$smtp.ToLowerInvariant()] = $true }
        }
    }
    return $map
}

function Add-DLMember {
    param(
        [string]$Identity, [string]$UserEmail, [string]$GroupEmail,
        [string]$GroupName, [string]$CsvFile, [hashtable]$MemberMap
    )

    $key = $UserEmail.ToLowerInvariant()
    if ($MemberMap.ContainsKey($key)) {
        Write-Log "      [SKIP] Ya es miembro: $UserEmail" 'WARN'
        Add-LogRow -CsvFile $CsvFile -GroupName $GroupName -GroupEmail $GroupEmail -Action 'MemberSkipped' -UserEmail $UserEmail -Detail 'Ya pertenece a la DL'
        return
    }

    if ($script:WhatIfMode) {
        Write-Log "      [WhatIf] Se añadiría: $UserEmail" 'WARN'
        Add-LogRow -CsvFile $CsvFile -GroupName $GroupName -GroupEmail $GroupEmail -Action 'WhatIf-MemberAdded' -UserEmail $UserEmail
        return
    }

    if ($PSCmdlet.ShouldProcess($UserEmail, "Añadir a la DL $GroupEmail")) {
        Add-DistributionGroupMember -Identity $Identity -Member $UserEmail -ErrorAction Stop
        $MemberMap[$key] = $true
        Write-Log "      [OK] Añadido: $UserEmail" 'OK'
        Add-LogRow -CsvFile $CsvFile -GroupName $GroupName -GroupEmail $GroupEmail -Action 'MemberAdded' -UserEmail $UserEmail
    }
}

# --- LÓGICA PRINCIPAL POR CSV ---

function Invoke-ProcessCsv {
    param([string]$CsvPath)

    $csvName = Split-Path $CsvPath -Leaf
    Write-Host ''
    Write-Host "  ┌─ $csvName" -ForegroundColor DarkGray
    Write-Log "Procesando CSV: $csvName"

    $rows = @()
    try {
        $rows = @(Import-GroupCsvRobust -Path $CsvPath)

        $lastGroupName = ''
        foreach ($row in $rows) {
            if (-not [string]::IsNullOrWhiteSpace($row.GroupName)) {
                $lastGroupName = $row.GroupName
            } elseif (-not [string]::IsNullOrWhiteSpace($lastGroupName)) {
                $row.GroupName = $lastGroupName
            }
        }
    }
    catch {
        Write-Log "  CSV ignorado: $_" 'ERROR'
        Add-LogRow -CsvFile $csvName -Action 'Error' -Detail "Lectura fallida: $_"
        return
    }

    if ($rows.Count -eq 0) {
        Write-Log '  CSV sin filas válidas.' 'WARN'
        return
    }

    $validGroupRows = @($rows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.GroupName) })
    $ignoredCount = $rows.Count - $validGroupRows.Count
    if ($ignoredCount -gt 0) {
        Write-Log "  Se ignoraron $ignoredCount fila(s) sin GroupName." 'WARN'
    }
    if ($validGroupRows.Count -eq 0) {
        Write-Log '  CSV sin filas con GroupName válido.' 'WARN'
        return
    }

    $grouped = $validGroupRows | Group-Object -Property GroupName

    foreach ($grp in $grouped) {
        $groupName  = $grp.Name.Trim()
        $groupEmail = "$groupName@$($script:Domain)"
        Write-Host ''
        Write-Log "  Grupo: $groupEmail  ($($grp.Group.Count) miembro(s) en CSV)"

        $validMembers = @($grp.Group | Where-Object {
            $em = $_.Email.Trim()
            if ([string]::IsNullOrWhiteSpace($em)) {
                Write-Log "    Fila vacía en Email — se omite." 'WARN'
                Add-LogRow -CsvFile $csvName -GroupName $groupName -GroupEmail $groupEmail -Action 'Error' -UserEmail '' -Detail 'Email vacío'
                return $false
            }
            if (-not (Test-EmailValue -Value $em)) {
                Write-Log "    Email inválido: '$em' — se omite." 'WARN'
                Add-LogRow -CsvFile $csvName -GroupName $groupName -GroupEmail $groupEmail -Action 'Error' -UserEmail $em -Detail 'Formato de email inválido'
                return $false
            }
            return $true
        })

        if ($script:GroupType -eq 'M365') {
            $mgGroup = $null
            try { $mgGroup = Get-OrCreate-M365Group -DisplayName $groupName -GroupEmail $groupEmail -CsvFile $csvName }
            catch {
                Write-Log "  Error creando grupo M365 '$groupEmail': $_" 'ERROR'
                Add-LogRow -CsvFile $csvName -GroupName $groupName -GroupEmail $groupEmail -Action 'Error' -Detail "Creación fallida: $_"
                continue
            }
            if ($null -eq $mgGroup) { continue }

            $memberMap = Get-M365GroupMembers -GroupId $mgGroup.Id

            foreach ($row in $validMembers) {
                $userEmail = $row.Email.Trim()
                try {
                    Add-M365Member -GroupId $mgGroup.Id -UserEmail $userEmail `
                                   -GroupEmail $groupEmail -GroupName $groupName `
                                   -CsvFile $csvName -MemberMap $memberMap
                }
                catch {
                    Write-Log "    Error añadiendo '$userEmail': $_" 'ERROR'
                    Add-LogRow -CsvFile $csvName -GroupName $groupName -GroupEmail $groupEmail -Action 'Error' -UserEmail $userEmail -Detail "Adición fallida: $_"
                }
            }
        }
        else {
            $dl = $null
            try { $dl = Get-OrCreate-DL -DisplayName $groupName -GroupEmail $groupEmail -CsvFile $csvName }
            catch {
                Write-Log "  Error creando DL '$groupEmail': $_" 'ERROR'
                Add-LogRow -CsvFile $csvName -GroupName $groupName -GroupEmail $groupEmail -Action 'Error' -Detail "Creación fallida: $_"
                continue
            }
            if ($null -eq $dl) { continue }

            $memberMap = Get-DLMemberMap -Identity $dl.PrimarySmtpAddress.ToString()

            foreach ($row in $validMembers) {
                $userEmail = $row.Email.Trim()
                try {
                    Add-DLMember -Identity $dl.PrimarySmtpAddress.ToString() `
                                 -UserEmail $userEmail -GroupEmail $groupEmail `
                                 -GroupName $groupName -CsvFile $csvName -MemberMap $memberMap
                }
                catch {
                    Write-Log "    Error añadiendo '$userEmail': $_" 'ERROR'
                    Add-LogRow -CsvFile $csvName -GroupName $groupName -GroupEmail $groupEmail -Action 'Error' -UserEmail $userEmail -Detail "Adición fallida: $_"
                }
            }
        }
    }
}

# --- RESUMEN ---

function Show-Summary {
    param([string]$LogCsvPath)

    $created  = @($script:LogRows | Where-Object { $_.Action -eq 'Created' }).Count
    $skipped  = @($script:LogRows | Where-Object { $_.Action -eq 'Skipped' }).Count
    $added    = @($script:LogRows | Where-Object { $_.Action -eq 'MemberAdded' }).Count
    $mSkipped = @($script:LogRows | Where-Object { $_.Action -eq 'MemberSkipped' }).Count
    $errors   = @($script:LogRows | Where-Object { $_.Action -eq 'Error' }).Count

    Write-Host ''
    Write-Host '══════════════════════════════════════════════' -ForegroundColor Magenta
    Write-Host '              RESUMEN FINAL                   ' -ForegroundColor Magenta
    Write-Host '══════════════════════════════════════════════' -ForegroundColor Magenta
    Write-Host ("  Grupos creados      : {0}" -f $created)  -ForegroundColor Green
    Write-Host ("  Grupos omitidos     : {0}" -f $skipped)  -ForegroundColor Yellow
    Write-Host ("  Miembros añadidos   : {0}" -f $added)    -ForegroundColor Green
    Write-Host ("  Miembros existentes : {0}" -f $mSkipped) -ForegroundColor Yellow
    Write-Host ("  Errores             : {0}" -f $errors)   -ForegroundColor Red
    Write-Host ''
    Write-Host ("  Log CSV             : $LogCsvPath") -ForegroundColor Cyan
    Write-Host '══════════════════════════════════════════════' -ForegroundColor Magenta
    Write-Host ''

    if ($script:WhatIfMode) {
        Write-Host '  WhatIf activo — no se realizaron cambios reales.' -ForegroundColor Yellow
        Write-Host ''
    }
}

# --- ENTRADA ---

try {
    Show-Header
    Assert-RequiredServicesReady

    Write-Host '  ¿Qué tipo de grupo quieres crear?' -ForegroundColor Yellow
    Write-Host '    [1]  Grupo de Microsoft 365  (Microsoft Graph)'
    Write-Host '    [2]  Lista de Distribución    (Exchange Online)'
    Write-Host ''
    do {
        $tipoRaw = (Read-Host '  Elige 1 o 2').Trim()
    } while ($tipoRaw -notin '1','2')
    $script:GroupType = if ($tipoRaw -eq '1') { 'M365' } else { 'DL' }
    Write-Log "Tipo seleccionado: $($script:GroupType)"
    Write-Host ''

    if ([string]::IsNullOrWhiteSpace($Domain)) {
        $domInput = (Read-Host '  Dominio para el email del grupo (ej: contoso.com)').Trim()
        if ([string]::IsNullOrWhiteSpace($domInput)) { throw 'Dominio requerido.' }
        $script:Domain = $domInput.TrimStart('@')
    }
    else {
        $script:Domain = $Domain.TrimStart('@')
    }
    Write-Log "Dominio: @$($script:Domain)"
    Write-Host ''

    $allFolders = New-Object System.Collections.Generic.List[string]
    if ($FolderPaths -and $FolderPaths.Count -gt 0) {
        foreach ($fp in $FolderPaths) { [void]$allFolders.Add((Normalize-QuotedInput $fp)) }
    }
    else {
        Write-Host '  Introduce las rutas de carpeta con CSVs (una por línea).'
        Write-Host '  Deja vacío y pulsa Enter cuando termines.' -ForegroundColor DarkGray
        if (Get-Command Show-CsvFormatHint -ErrorAction SilentlyContinue) {
            Show-CsvFormatHint -Schema 'EmailGroupName'
        }
        Write-Host ''
        $idx = 1
        while ($true) {
            $raw = Normalize-QuotedInput (Read-Host "  Carpeta $idx")
            if ([string]::IsNullOrWhiteSpace($raw)) { break }
            if (-not (Test-Path -LiteralPath $raw -PathType Container)) {
                Write-Host "    No existe o no es carpeta: $raw" -ForegroundColor Red
                continue
            }
            [void]$allFolders.Add($raw); $idx++
        }
    }

    if ($allFolders.Count -eq 0) {
        Write-Log 'Ninguna carpeta válida. Operación cancelada.' 'ERROR'
        exit 1
    }

    $allCsvs = New-Object System.Collections.Generic.List[string]
    foreach ($folder in $allFolders) {
        $found = @(Get-ChildItem -LiteralPath $folder -Filter '*.csv' -File | Where-Object { $_.Name -notmatch '^Log_' })
        if ($found.Count -eq 0) {
            Write-Log "Sin CSVs en: $folder" 'WARN'
        }
        else {
            foreach ($f in $found) { [void]$allCsvs.Add($f.FullName) }
            Write-Log "$($found.Count) CSV(s) en: $folder"
        }
    }

    if ($allCsvs.Count -eq 0) {
        Write-Log 'No hay CSVs para procesar.' 'ERROR'
        exit 1
    }

    Write-Log "Total CSVs a procesar: $($allCsvs.Count)"
    Write-Host ''

    foreach ($csvPath in $allCsvs) { Invoke-ProcessCsv -CsvPath $csvPath }

    $logFolder = Split-Path $allCsvs[0] -Parent
    $logCsvPath = Join-Path $logFolder ("Log_NewGroups_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.csv')
    if ($script:LogRows.Count -gt 0) {
        $script:LogRows | Export-Csv -Path $logCsvPath -NoTypeInformation -Encoding UTF8
        Write-Log "Log CSV exportado: $logCsvPath" 'OK'
    }

    Show-Summary -LogCsvPath $logCsvPath
}
catch {
    Write-Host ''
    Write-Host "[FATAL] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
    Write-Host ''
}
