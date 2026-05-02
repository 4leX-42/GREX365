# --- HELPERS DE CONSOLA ---

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
        [string]$Color = 'Gray'
    )

    $pad = Get-CenterPadding -Text $Text
    Write-Host ($pad + $Text) -ForegroundColor $Color
}

function Write-CenteredSegments {
    param(
        [Parameter(Mandatory = $true)][array]$Segments,
        [int]$Offset = 0
    )

    $plain = (($Segments | ForEach-Object { [string]$_.Text }) -join '')
    $basePad = [Math]::Floor(([Console]::WindowWidth - $plain.Length) / 2)
    $padCount = $basePad + $Offset
    if ($padCount -lt 0) { $padCount = 0 }

    Write-Host (' ' * $padCount) -NoNewline
    foreach ($seg in $Segments) {
        Write-Host $seg.Text -NoNewline -ForegroundColor $seg.Color
    }
    Write-Host ''
}

function Write-SideSegments {
    param(
        [Parameter(Mandatory = $true)][array]$LeftSegments,
        [Parameter(Mandatory = $true)][array]$RightSegments,
        [int]$Margin = 2
    )

    $consoleWidth = [Console]::WindowWidth
    $leftText  = (($LeftSegments  | ForEach-Object { [string]$_.Text }) -join '')
    $rightText = (($RightSegments | ForEach-Object { [string]$_.Text }) -join '')

    $usable = $consoleWidth - ($Margin * 2)
    if ($usable -lt 0) { $usable = 0 }
    $gap = $usable - $leftText.Length - $rightText.Length

    if ($gap -lt 1) {
        Write-CenteredSegments -Segments $LeftSegments
        Write-CenteredSegments -Segments $RightSegments
        return
    }

    Write-Host (' ' * $Margin) -NoNewline
    foreach ($seg in $LeftSegments)  { Write-Host $seg.Text -NoNewline -ForegroundColor $seg.Color }
    Write-Host (' ' * $gap) -NoNewline
    foreach ($seg in $RightSegments) { Write-Host $seg.Text -NoNewline -ForegroundColor $seg.Color }
    Write-Host ''
}

function Write-PanelLine {
    param(
        [int]$Width = 76,
        [string]$Color = 'DarkCyan',
        [string]$Left = '╔',
        [string]$Fill = '═',
        [string]$Right = '╗'
    )

    if ($Width -lt 2) { $Width = 2 }
    $line = $Left + ($Fill * ($Width - 2)) + $Right
    Write-Centered -Text $line -Color $Color
}

# --- ESTADO DE SESIÓN (para cabecera) ---

function Get-SessionState {
    $exo = 'OFFLINE'
    $graph = 'OFFLINE'

    try {
        if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
            $info = Get-ConnectionInformation -ErrorAction SilentlyContinue
            if ($info | Where-Object { $_.State -eq 'Connected' }) { $exo = 'ONLINE' }
        }
    } catch {}

    try {
        if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
            $ctx = Get-MgContext -ErrorAction SilentlyContinue
            if ($ctx) {
                $isAppOnly = ([string]$ctx.AuthType -match 'AppOnly')
                if ($isAppOnly -and $ctx.ClientId -and $ctx.TenantId) { $graph = 'ONLINE' }
                elseif ($ctx.Account) { $graph = 'ONLINE' }
            }
        }
    } catch {}

    [PSCustomObject]@{ EXO = $exo; GRAPH = $graph }
}

function Format-MethodLabel {
    param([string]$Method)

    switch ($Method) {
        'cert'        { return 'CERT (App)' }
        'traditional' { return 'DEVICE CODE' }
        default       { return 'NO ELEGIDO' }
    }
}

# --- CABECERA UNIFICADA ---

function Show-Header {
    param(
        [string]$Title = 'GREX365',
        [string]$Subtitle = 'EXO // M365 // GRAPH',
        [string]$ActiveMethod = $null
    )

    $state = Get-SessionState
    $exoColor   = if ($state.EXO   -eq 'ONLINE') { 'Green' } else { 'DarkGray' }
    $graphColor = if ($state.GRAPH -eq 'ONLINE') { 'Green' } else { 'DarkGray' }

    Clear-Host
    Write-Host ''

    Write-SideSegments -Margin 3 -LeftSegments @(
        @{ Text = '['; Color = 'DarkGray' }
        @{ Text = 'EXO'; Color = 'DarkCyan' }
        @{ Text = ':'; Color = 'DarkGray' }
        @{ Text = $state.EXO; Color = $exoColor }
        @{ Text = ']'; Color = 'DarkGray' }
    ) -RightSegments @(
        @{ Text = '['; Color = 'DarkGray' }
        @{ Text = 'GRAPH'; Color = 'DarkCyan' }
        @{ Text = ':'; Color = 'DarkGray' }
        @{ Text = $state.GRAPH; Color = $graphColor }
        @{ Text = ']'; Color = 'DarkGray' }
    )

    if ($ActiveMethod) {
        $methodColor = if ($ActiveMethod -eq 'cert') { 'Cyan' } elseif ($ActiveMethod -eq 'traditional') { 'Yellow' } else { 'DarkGray' }
        Write-CenteredSegments @(
            @{ Text = 'Método activo: '; Color = 'DarkGray' }
            @{ Text = (Format-MethodLabel -Method $ActiveMethod); Color = $methodColor }
        )
    }

    Write-Host ''
    Write-PanelLine -Width 76 -Color 'DarkCyan' -Left '╔' -Fill '═' -Right '╗'

    Write-CenteredSegments @(
        @{ Text = '>> '; Color = 'DarkGray' }
        @{ Text = $Title; Color = 'Cyan' }
        @{ Text = ' <<'; Color = 'DarkGray' }
    )

    Write-Centered -Text $Subtitle -Color 'DarkGray'
    Write-PanelLine -Width 76 -Color 'DarkCyan' -Left '╚' -Fill '═' -Right '╝'
    Write-Host ''
}

function Get-MenuBoxLayout {
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [int]$Width = 52,
        [bool]$Selected = $false,
        [bool]$ComingSoon = $false
    )

    if ($Width -lt 14) { $Width = 14 }

    $marker  = if ($Selected) { '▶  ' } else { '   ' }
    $display = $marker + $Text
    $maxText = $Width - 6
    if ($display.Length -gt $maxText) { $display = $display.Substring(0, $maxText) }

    $padR = $Width - 6 - $display.Length
    if ($padR -lt 0) { $padR = 0 }

    if ($ComingSoon) {
        $bColor = 'DarkGray'; $tColor = 'DarkGray'
        $top   = '╭' + ('─' * ($Width - 2)) + '╮'
        $blank = '│' + (' ' * ($Width - 2)) + '│'
        $bot   = '╰' + ('─' * ($Width - 2)) + '╯'
        $lb = '│  '; $rb = '  │'
    }
    elseif ($Selected) {
        $bColor = 'Cyan'; $tColor = 'White'
        $top   = '╔' + ('═' * ($Width - 2)) + '╗'
        $blank = '║' + (' ' * ($Width - 2)) + '║'
        $bot   = '╚' + ('═' * ($Width - 2)) + '╝'
        $lb = '║  '; $rb = '  ║'
    }
    else {
        $bColor = 'DarkCyan'; $tColor = 'Gray'
        $top   = '╭' + ('─' * ($Width - 2)) + '╮'
        $blank = '│' + (' ' * ($Width - 2)) + '│'
        $bot   = '╰' + ('─' * ($Width - 2)) + '╯'
        $lb = '│  '; $rb = '  │'
    }

    [PSCustomObject]@{
        Top         = $top
        Blank       = $blank
        Bottom      = $bot
        LB          = $lb
        RB          = $rb
        Display     = $display
        RightPad    = $padR
        BorderColor = $bColor
        TextColor   = $tColor
        Width       = $Width
    }
}

function Write-MenuRow {
    param(
        [Parameter(Mandatory = $true)]$Left,
        $Right,
        [int]$Gap = 8
    )

    $totalLen = $Left.Width + $(if ($Right) { $Gap + $Right.Width } else { 0 })
    $consoleW = [Console]::WindowWidth
    $padBase  = [Math]::Floor(($consoleW - $totalLen) / 2)
    if ($padBase -lt 0) { $padBase = 0 }
    $pad    = ' ' * $padBase
    $gapStr = ' ' * $Gap

    # Línea 1: borde superior
    Write-Host $pad -NoNewline
    Write-Host $Left.Top -NoNewline -ForegroundColor $Left.BorderColor
    if ($Right) {
        Write-Host $gapStr -NoNewline
        Write-Host $Right.Top -ForegroundColor $Right.BorderColor
    } else { Write-Host '' }

    # Línea 2: padding superior interno
    Write-Host $pad -NoNewline
    Write-Host $Left.Blank -NoNewline -ForegroundColor $Left.BorderColor
    if ($Right) {
        Write-Host $gapStr -NoNewline
        Write-Host $Right.Blank -ForegroundColor $Right.BorderColor
    } else { Write-Host '' }

    # Línea 3: contenido
    Write-Host $pad -NoNewline
    Write-Host $Left.LB -NoNewline -ForegroundColor $Left.BorderColor
    Write-Host $Left.Display -NoNewline -ForegroundColor $Left.TextColor
    Write-Host (' ' * $Left.RightPad) -NoNewline
    Write-Host $Left.RB -NoNewline -ForegroundColor $Left.BorderColor
    if ($Right) {
        Write-Host $gapStr -NoNewline
        Write-Host $Right.LB -NoNewline -ForegroundColor $Right.BorderColor
        Write-Host $Right.Display -NoNewline -ForegroundColor $Right.TextColor
        Write-Host (' ' * $Right.RightPad) -NoNewline
        Write-Host $Right.RB -ForegroundColor $Right.BorderColor
    } else { Write-Host '' }

    # Línea 4: padding inferior interno
    Write-Host $pad -NoNewline
    Write-Host $Left.Blank -NoNewline -ForegroundColor $Left.BorderColor
    if ($Right) {
        Write-Host $gapStr -NoNewline
        Write-Host $Right.Blank -ForegroundColor $Right.BorderColor
    } else { Write-Host '' }

    # Línea 5: borde inferior
    Write-Host $pad -NoNewline
    Write-Host $Left.Bottom -NoNewline -ForegroundColor $Left.BorderColor
    if ($Right) {
        Write-Host $gapStr -NoNewline
        Write-Host $Right.Bottom -ForegroundColor $Right.BorderColor
    } else { Write-Host '' }
}

function Write-Watermark {
    param(
        [string]$Text = 'GREX365',
        [string]$Color = 'DarkGray'
    )

    try {
        $raw = $Host.UI.RawUI
        $bufWidth = $raw.BufferSize.Width
        $winHeight = $raw.WindowSize.Height
        $winTop = $raw.WindowPosition.Y

        $x = 1
        $y = $winTop + $winHeight - 2
        if ($y -lt 0) { $y = 0 }

        $oldPos = $raw.CursorPosition
        $oldColor = $raw.ForegroundColor
        $raw.CursorPosition = New-Object System.Management.Automation.Host.Coordinates($x, $y)
        $raw.ForegroundColor = $Color
        [Console]::Write($Text)
        $raw.ForegroundColor = $oldColor
        $raw.CursorPosition = $oldPos
    }
    catch {}
}

function Pause-ReturnToMenu {
    Write-Host ''
    Write-Centered -Text 'Pulsa cualquier tecla para volver al menú...' -Color 'DarkGray'
    [void][System.Console]::ReadKey($true)
}

# --- PANEL DE ESTADO ---
# Muestra: conexión activa (tenant/EXO/dominio), módulos requeridos
# (instalados/rutas) y archivos JSON de configuración (rutas/existencia).

function Show-StatusPanel {
    Write-Host ''
    Write-Host '  ┌─ ESTADO DEL SISTEMA ──────────────────────────────────────────────' -ForegroundColor DarkCyan
    Write-Host '  │' -ForegroundColor DarkCyan

    # --- Conexión ---
    $conn = Get-ToolkitConnectionState
    $exoTag   = if ($conn.ExoConnected)   { '[ONLINE] ' } else { '[OFFLINE]' }
    $graphTag = if ($conn.GraphConnected) { '[ONLINE] ' } else { '[OFFLINE]' }
    $exoColor   = if ($conn.ExoConnected)   { 'Green' } else { 'DarkGray' }
    $graphColor = if ($conn.GraphConnected) { 'Green' } else { 'DarkGray' }

    Write-Host '  │ ' -ForegroundColor DarkCyan -NoNewline
    Write-Host 'CONEXIÓN ACTIVA' -ForegroundColor White
    $accountText = if ($conn.Account)         { $conn.Account }         else { '— sin sesión —' }
    $exoOrgText  = if ($conn.ExchangeOrgName) { $conn.ExchangeOrgName } else { '— sin sesión —' }
    $tenantText  = if ($conn.TenantId)        { $conn.TenantId }        else { '—' }
    $domainText  = if ($conn.DefaultDomain)   { $conn.DefaultDomain }   else { '—' }

    Write-Host '  │   Microsoft Graph  : ' -ForegroundColor DarkCyan -NoNewline
    Write-Host $graphTag -ForegroundColor $graphColor -NoNewline
    Write-Host (' ' + $accountText) -ForegroundColor Gray
    Write-Host '  │   Exchange Online  : ' -ForegroundColor DarkCyan -NoNewline
    Write-Host $exoTag -ForegroundColor $exoColor -NoNewline
    Write-Host (' ' + $exoOrgText) -ForegroundColor Gray
    Write-Host '  │   Tenant ID        : ' -ForegroundColor DarkCyan -NoNewline
    Write-Host $tenantText -ForegroundColor Gray
    Write-Host '  │   Dominio default  : ' -ForegroundColor DarkCyan -NoNewline
    Write-Host $domainText -ForegroundColor Gray

    Write-Host '  │' -ForegroundColor DarkCyan

    # --- Módulos ---
    Write-Host '  │ ' -ForegroundColor DarkCyan -NoNewline
    Write-Host 'MÓDULOS REQUERIDOS' -ForegroundColor White
    foreach ($mod in (Get-ToolkitModuleStatus)) {
        $tag = if ($mod.Installed) { '[OK]' } else { '[X] ' }
        $tagColor = if ($mod.Installed) { 'Green' } else { 'Red' }
        $verSuffix = if ($mod.Installed) { " (v$($mod.Version))" } else { ' (no instalado)' }
        Write-Host '  │   ' -ForegroundColor DarkCyan -NoNewline
        Write-Host $tag -ForegroundColor $tagColor -NoNewline
        Write-Host (' ' + $mod.Name + $verSuffix) -ForegroundColor Gray
        if ($mod.Installed -and $mod.Path) {
            Write-Host ('  │       ' + $mod.Path) -ForegroundColor DarkGray
        }
    }

    Write-Host '  │' -ForegroundColor DarkCyan

    # --- Archivos config ---
    Write-Host '  │ ' -ForegroundColor DarkCyan -NoNewline
    Write-Host 'ARCHIVOS DE CONFIGURACIÓN' -ForegroundColor White
    foreach ($file in (Get-ToolkitConfigFiles)) {
        $tag = if ($file.Exists) { '[OK]' } else { '[X] ' }
        $tagColor = if ($file.Exists) { 'Green' } else { 'DarkYellow' }
        Write-Host '  │   ' -ForegroundColor DarkCyan -NoNewline
        Write-Host $tag -ForegroundColor $tagColor -NoNewline
        Write-Host (' ' + $file.Name) -ForegroundColor Gray
        if ($file.Path) {
            Write-Host ('  │       ' + $file.Path) -ForegroundColor DarkGray
        }
        if ($file.Description) {
            Write-Host ('  │       ' + $file.Description) -ForegroundColor DarkGray
        }
    }

    Write-Host '  │' -ForegroundColor DarkCyan
    Write-Host '  └────────────────────────────────────────────────────────────────────' -ForegroundColor DarkCyan
    Write-Host ''
}

# --- MENÚ PRINCIPAL DINÁMICO ---
# $Items: array de hashtables @{ Label; ComingSoon }
# Devuelve PSCustomObject @{ Action; Value }
# Acciones: 'option' (Value=índice 1-based), 'toggle' (Value='cert'|'traditional'),
#           'wizard', 'exit'

function Show-MainMenu {
    param(
        [Parameter(Mandatory = $true)][array]$Items,
        [string]$ActiveMethod
    )

    $selected = 0
    $count = $Items.Count
    $cols = 2
    $rows = [Math]::Ceiling($count / [double]$cols)
    $boxW = 52
    $gap  = 8

    while ($true) {
        Show-Header -Title 'GREX365' -Subtitle 'EXO · M365 · GRAPH' -ActiveMethod $ActiveMethod

        Write-Host ''

        for ($r = 0; $r -lt $rows; $r++) {
            $iL = $r * $cols
            $iR = $iL + 1

            $left = $null; $right = $null
            if ($iL -lt $count) {
                $itemL = $Items[$iL]
                $labelL = if ($itemL -is [hashtable]) { $itemL.Label } else { [string]$itemL }
                $comingL = $false
                if ($itemL -is [hashtable] -and $itemL.ContainsKey('ComingSoon')) { $comingL = [bool]$itemL.ComingSoon }
                $left = Get-MenuBoxLayout -Text ("0{0}  ·  {1}" -f ($iL + 1), $labelL) -Width $boxW -Selected ($iL -eq $selected) -ComingSoon $comingL
            }
            if ($iR -lt $count) {
                $itemR = $Items[$iR]
                $labelR = if ($itemR -is [hashtable]) { $itemR.Label } else { [string]$itemR }
                $comingR = $false
                if ($itemR -is [hashtable] -and $itemR.ContainsKey('ComingSoon')) { $comingR = [bool]$itemR.ComingSoon }
                $right = Get-MenuBoxLayout -Text ("0{0}  ·  {1}" -f ($iR + 1), $labelR) -Width $boxW -Selected ($iR -eq $selected) -ComingSoon $comingR
            }

            Write-MenuRow -Left $left -Right $right -Gap $gap
            Write-Host ''
        }

        # Footer minimalista al final
        Write-Host ''
        $alt = if ($ActiveMethod -eq 'cert') { 'traditional' } else { 'cert' }
        $altKey = if ($alt -eq 'cert') { 'C' } else { 'T' }
        $altLabel = if ($alt -eq 'cert') { 'cert' } else { 'device code' }

        $sepWidth = ($boxW * 2) + $gap
        $sepLine  = '·' * $sepWidth
        Write-Centered -Text $sepLine -Color 'DarkGray'
        Write-Host ''

        Write-CenteredSegments @(
            @{ Text = '↑ ↓ ← →'; Color = 'DarkCyan' }
            @{ Text = '   navegar'; Color = 'DarkGray' }
            @{ Text = '     '; Color = 'DarkGray' }
            @{ Text = $altKey; Color = 'DarkYellow' }
            @{ Text = '   '; Color = 'DarkGray' }
            @{ Text = $altLabel; Color = 'DarkGray' }
            @{ Text = '     '; Color = 'DarkGray' }
            @{ Text = '0'; Color = 'DarkYellow' }
            @{ Text = '   '; Color = 'DarkGray' }
            @{ Text = 'salir'; Color = 'DarkGray' }
        )

        Write-Watermark -Text 'GREX365 // by Andreu' -Color 'DarkGray'

        $key = [System.Console]::ReadKey($true)

        switch ($key.Key) {
            'UpArrow' {
                $newSel = $selected - $cols
                if ($newSel -lt 0) {
                    $col = $selected % $cols
                    $lastRow = $rows - 1
                    $newSel = $lastRow * $cols + $col
                    if ($newSel -ge $count) { $newSel = $count - 1 }
                }
                $selected = $newSel
            }
            'DownArrow' {
                $newSel = $selected + $cols
                if ($newSel -ge $count) { $newSel = $selected % $cols }
                $selected = $newSel
            }
            'LeftArrow' {
                if ($selected -gt 0) { $selected-- } else { $selected = $count - 1 }
            }
            'RightArrow' {
                if ($selected -lt ($count - 1)) { $selected++ } else { $selected = 0 }
            }
            'Escape'    { return [PSCustomObject]@{ Action = 'exit'; Value = $null } }
            'D0'        { return [PSCustomObject]@{ Action = 'exit'; Value = $null } }
            'NumPad0'   { return [PSCustomObject]@{ Action = 'exit'; Value = $null } }
            'Enter'     { return [PSCustomObject]@{ Action = 'option'; Value = ($selected + 1) } }
            'T'         { return [PSCustomObject]@{ Action = 'toggle'; Value = 'traditional' } }
            'C'         { return [PSCustomObject]@{ Action = 'toggle'; Value = 'cert' } }
            default {
                $ch = $key.KeyChar.ToString()
                if ($ch -match '^[1-9]$') {
                    $n = [int]$ch
                    if ($n -ge 1 -and $n -le $count) {
                        return [PSCustomObject]@{ Action = 'option'; Value = $n }
                    }
                }
            }
        }
    }
}

# --- SELECTOR INICIAL DE MÉTODO ---

function Show-MethodSelector {
    Show-Header -Title 'GREX365' -Subtitle 'Selección inicial de método de conexión'

    Write-Centered -Text 'Antes de empezar, elige cómo te conectarás a Microsoft 365.' -Color 'White'
    Write-Host ''
    Write-Centered -Text '[A] Conexión por CERTIFICADO (App-only, desatendido)' -Color 'Cyan'
    Write-Centered -Text '    Recomendado. Requiere asistente de creación de certificado la primera vez.' -Color 'DarkGray'
    Write-Host ''
    Write-Centered -Text '[B] Conexión TRADICIONAL (device code, delegado)' -Color 'Yellow'
    Write-Centered -Text '    Usa la cuenta del operador. Necesita login interactivo cada sesión.' -Color 'DarkGray'
    Write-Host ''
    Write-Centered -Text 'Esta preferencia se guarda y puedes cambiarla luego con [T]/[C].' -Color 'DarkGray'
    Write-Host ''

    while ($true) {
        $key = [System.Console]::ReadKey($true)
        switch ($key.Key) {
            'A' { return 'cert' }
            'B' { return 'traditional' }
            'D1' { return 'cert' }
            'D2' { return 'traditional' }
            'NumPad1' { return 'cert' }
            'NumPad2' { return 'traditional' }
            'Escape' { return $null }
        }
    }
}
