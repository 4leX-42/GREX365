#requires -Version 5.1
[CmdletBinding()]
param()

$Host.UI.RawUI.WindowTitle = "Exportador de miembros | Email + Object ID"

function Write-Log {
    param(
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

function Get-CenterPadding {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    $width = [Console]::WindowWidth
    $pad = [Math]::Floor(($width - $Text.Length) / 2)

    if ($pad -lt 0) { $pad = 0 }
    return (" " * $pad)
}

function Write-Centered {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [ConsoleColor]$Color = [ConsoleColor]::Gray
    )

    $pad = Get-CenterPadding -Text $Text
    Write-Host ($pad + $Text) -ForegroundColor $Color
}

function Show-Header {
    Clear-Host

    $width = (Get-Host).UI.RawUI.WindowSize.Width
    if ($width -lt 96) { $width = 96 }

    $leftAccent  = "◀◀◀"
    $rightAccent = "▶▶▶"
    $rail        = "════════════"
    $sep         = "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

    Write-Host ""
    Write-Centered "$leftAccent $rail  Grex365  $rail $rightAccent" -Color DarkCyan
    Write-Host ""

    Write-Centered "EXPORTAR USUARIOS | EMAIL + OBJECT ID" -Color Green
    Write-Centered $sep -Color Cyan
    Write-Centered "DL & MICROSOFT 365 GROUPS | AUTO-DETECTION" -Color DarkCyan

    Write-Host ""
}

# --- VERIFICACIÓN DE CONEXIÓN ---
# La conexión real se hace en Main.ps1 (Connect-RequiredServices del módulo).
# Aquí solo validamos que estén activas las sesiones que necesitamos.

function Assert-RequiredServicesReady {
    $exoOk = $false
    $mgOk  = $false

    try {
        if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
            $exoSession = Get-ConnectionInformation -ErrorAction SilentlyContinue
            if ($exoSession | Where-Object { $_.State -eq 'Connected' }) { $exoOk = $true }
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
        throw "Faltan servicios M365 conectados (EXO=$exoOk, Graph=$mgOk). Ejecuta este script desde Main.ps1."
    }
}

# Funciones de búsqueda y selección (Normalize-SearchText, Get-Search*, Get-GraphGroupCandidates,
# Get-ExchangeGroupCandidates, Merge-GroupCandidates, Show-GroupSelectionMenu, Resolve-GroupByMail)
# viven en Modules/Common.ps1 — compartidas con Add-MembersToGroup_Fixed.ps1.

function Resolve-UserId {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Email
    )

    $mail = $Email.Trim()

    if ([string]::IsNullOrWhiteSpace($mail)) {
        return $null
    }

    try {
        $safeMail = $mail.Replace("'", "''")
        $user = Get-MgUser -Filter "mail eq '$safeMail'" -ConsistencyLevel eventual -ErrorAction SilentlyContinue
        if ($user) {
            return ([string](($user | Select-Object -First 1).Id)).Trim()
        }
    }
    catch {}

    try {
        $safeMail = $mail.Replace("'", "''")
        $user = Get-MgUser -Filter "userPrincipalName eq '$safeMail'" -ConsistencyLevel eventual -ErrorAction SilentlyContinue
        if ($user) {
            return ([string](($user | Select-Object -First 1).Id)).Trim()
        }
    }
    catch {}

    return $null
}

function Export-CleanCsv {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.List[object]]$Rows,

        [Parameter(Mandatory = $true)]
        [string]$OutputCsv
    )

    $Rows |
        Select-Object @{
            Name = 'Email'
            Expression = { [string]$_.Email }
        }, @{
            Name = 'Id'
            Expression = { [string]$_.Id }
        } |
        Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8 -Delimiter ';'
}

function Export-DistributionGroupMembers {
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupIdentity,

        [Parameter(Mandatory = $true)]
        [string]$OutputCsv
    )

    Write-Log "Obteniendo miembros de la lista de distribución..." "INFO"

    $members = Get-DistributionGroupMember -Identity $GroupIdentity -ResultSize Unlimited -ErrorAction Stop
    $results = New-Object System.Collections.Generic.List[object]

    $total = @($members).Count
    $i = 0

    foreach ($m in $members) {
        $i++
        $mail = $null

        if ($m.PrimarySmtpAddress) {
            $mail = ([string]$m.PrimarySmtpAddress).Trim()
        }

        Write-Log "[$i/$total] Procesando: $mail" "INFO"

        $id = $null
        if ($mail) {
            $id = Resolve-UserId -Email $mail
        }

        $results.Add([PSCustomObject]@{
            Email = $mail
            Id    = $id
        })
    }

    Export-CleanCsv -Rows $results -OutputCsv $OutputCsv
}

function Export-UnifiedGroupMembers {
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupIdentity,

        [Parameter(Mandatory = $true)]
        [string]$OutputCsv
    )

    Write-Log "Obteniendo miembros del grupo de Microsoft 365..." "INFO"

    $members = Get-UnifiedGroupLinks -Identity $GroupIdentity -LinkType Members -ResultSize Unlimited -ErrorAction Stop
    $results = New-Object System.Collections.Generic.List[object]

    $total = @($members).Count
    $i = 0

    foreach ($m in $members) {
        $i++
        $mail = $null
        $id   = $null

        if ($m.PrimarySmtpAddress) {
            $mail = ([string]$m.PrimarySmtpAddress).Trim()
        }

        if ($m.ExternalDirectoryObjectId) {
            $id = ([string]$m.ExternalDirectoryObjectId).Trim()
        }

        if (-not $id -and $mail) {
            $id = Resolve-UserId -Email $mail
        }

        Write-Log "[$i/$total] Procesando: $mail" "INFO"

        $results.Add([PSCustomObject]@{
            Email = $mail
            Id    = $id
        })
    }

    Export-CleanCsv -Rows $results -OutputCsv $OutputCsv
}

try {
    Show-Header
    Assert-RequiredServicesReady

    $groupMail = Read-Host "Introduce el correo del grupo o lista de distribución"
    if ([string]::IsNullOrWhiteSpace($groupMail)) {
        throw "No has introducido ningún correo o texto de búsqueda."
    }

    $groupInfo = Resolve-GroupByMail -GroupMail $groupMail

    if ($groupInfo.Cancelled) {
        Write-Log "Selección cancelada por el usuario." "WARN"
        return
    }

    if (-not $groupInfo.Found) {
        throw "No se encontró ningún grupo ni lista de distribución que coincida con: $groupMail"
    }

    Clear-Host
    Show-Header
    Write-Host ""
    Write-Centered "OBJETO ENCONTRADO" -Color Cyan
    Write-Host ""

    $lines = @(
        ("{0,-10}: {1}" -f "Nombre", $groupInfo.DisplayName)
        ("{0,-10}: {1}" -f "Correo", $groupInfo.PrimarySmtpAddress)
        ("{0,-10}: {1}" -f "Tipo",   $groupInfo.GroupType)
        ("{0,-10}: {1}" -f "ID",     $groupInfo.GroupId)
    )

    $maxLen = ($lines | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum

    foreach ($line in $lines) {
        Write-Centered ($line.PadRight($maxLen)) -Color Gray
    }

    Write-Host ""

    $confirm = Read-Host "¿Quieres exportar los miembros de este grupo? (S/N)"
    if ($confirm -notmatch '^(S|SI|Y|YES)$') {
        Write-Log "Operación cancelada por el usuario." "WARN"
        return
    }

    $outputFolder = Read-Host "Carpeta de salida CSV [ENTER = C:\Temp]"
    if ([string]::IsNullOrWhiteSpace($outputFolder)) {
        $outputFolder = "C:\Temp"
    }

    Ensure-Folder -Path $outputFolder

    $safeName = ($groupInfo.PrimarySmtpAddress -replace '[\\/:*?"<>|]', '_')
    $outputCsv = Join-Path $outputFolder ("{0}_Members_Email_ID.csv" -f $safeName)

    switch ($groupInfo.GroupType) {
        'DistributionList' {
            Export-DistributionGroupMembers -GroupIdentity $groupInfo.Identity -OutputCsv $outputCsv
        }
        'Microsoft365Group' {
            Export-UnifiedGroupMembers -GroupIdentity $groupInfo.Identity -OutputCsv $outputCsv
        }
        default {
            throw "Tipo de grupo no soportado: $($groupInfo.GroupType)"
        }
    }

    Write-Host ""
    Write-Log "CSV generado correctamente: $outputCsv" 'OK'
    Write-Host ""
}
catch {
    Write-Host ""
    Write-Log $_.Exception.Message 'ERROR'
    Write-Host ""
}
