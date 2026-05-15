# Menu module
# Microsoft Admin Center-style main menu with differential rendering.
#
# Performance design:
#  - Full screen is painted once on entry / resize / toggle.
#  - Keyboard navigation repaints ONLY the two affected rows
#    (previous selected + new selected) via SetCursorPosition + ANSI erase line.
#  - No Clear-Host on keystroke -> no flicker.
#  - Cursor is hidden during render -> no caret flash.
#  - Held arrow keys are drained (KeyAvailable loop) so the screen does not
#    fight against a flood of repaints.

# ANSI constants. Use $global: so Menu functions still work if a child script
# (which has its own $script: scope under StrictMode) ever invokes them.
if (-not (Get-Variable -Name GREX365_AnsiReset -Scope Global -ErrorAction SilentlyContinue)) {
    $global:GREX365_AnsiReset        = "$([char]27)[0m"
    $global:GREX365_AnsiEraseLine    = "$([char]27)[2K"
    $global:GREX365_AnsiFgCyan       = "$([char]27)[36m"
    $global:GREX365_AnsiFgBrightCyan = "$([char]27)[96m"
    $global:GREX365_AnsiFgGray       = "$([char]27)[37m"
    $global:GREX365_AnsiFgDarkGray   = "$([char]27)[90m"
    $global:GREX365_AnsiFgWhite      = "$([char]27)[97m"
}

function Format-MenuItemLine {
    param(
        [int]$Index,
        [hashtable]$Item,
        [bool]$IsSelected,
        [int]$LabelPad = 42
    )

    $label = [string]$Item.Label
    $tag   = if ($Item.ContainsKey('Tag')) { [string]$Item.Tag } else { '' }

    $indent = ' ' * ($global:GREX365_ConsolePadding * 2)
    $marker = if ($IsSelected) { '>' } else { ' ' }
    $numAnsi  = if ($IsSelected) { $global:GREX365_AnsiFgBrightCyan } else { $global:GREX365_AnsiFgDarkGray }
    $textAnsi = if ($IsSelected) { $global:GREX365_AnsiFgBrightCyan } else { $global:GREX365_AnsiFgGray }

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append($indent)
    [void]$sb.Append($global:GREX365_AnsiFgCyan).Append($marker).Append(' ').Append($global:GREX365_AnsiReset)
    [void]$sb.Append($numAnsi).AppendFormat('{0,2}   ', ($Index + 1)).Append($global:GREX365_AnsiReset)
    [void]$sb.Append($textAnsi).Append($label.PadRight($LabelPad)).Append($global:GREX365_AnsiReset)
    if ($tag) {
        [void]$sb.Append($global:GREX365_AnsiFgDarkGray).Append($tag).Append($global:GREX365_AnsiReset)
    }
    return $sb.ToString()
}

function Get-SafeCursorTop {
    try { return [Console]::CursorTop } catch { return -1 }
}

function Render-MainMenuFull {
    param(
        [Parameter(Mandatory)][array]$Items,
        [Parameter(Mandatory)][int]$SelectedIndex,
        [string]$ActiveMethod,
        [Parameter(Mandatory)]$Sections,
        [Parameter(Mandatory)][array]$SectionOrder,
        [Parameter(Mandatory)][ref]$ItemRowsOut
    )

    Show-Header -Title 'GREX365' -Subtitle 'Microsoft 365 administration toolkit' -ActiveMethod $ActiveMethod

    $rows = @{}
    $indentL1 = ' ' * $global:GREX365_ConsolePadding
    $indentL2 = ' ' * ($global:GREX365_ConsolePadding * 2)

    foreach ($section in $SectionOrder) {
        [Console]::Write($indentL1)
        [Console]::Write($global:GREX365_AnsiFgWhite + $section.ToUpperInvariant() + $global:GREX365_AnsiReset)
        [Console]::WriteLine()
        [Console]::WriteLine()

        foreach ($entry in $Sections[$section]) {
            $rows[$entry.Index] = Get-SafeCursorTop
            $line = Format-MenuItemLine -Index $entry.Index -Item $entry.Item -IsSelected ($entry.Index -eq $SelectedIndex)
            [Console]::WriteLine($line)
        }
        [Console]::WriteLine()
    }

    [Console]::Write($indentL2)
    [Console]::Write($global:GREX365_AnsiFgCyan + '> ' + $global:GREX365_AnsiReset)
    [Console]::Write($global:GREX365_AnsiFgDarkGray + ' 0   ' + $global:GREX365_AnsiReset)
    [Console]::WriteLine($global:GREX365_AnsiFgDarkGray + 'Salir' + $global:GREX365_AnsiReset)
    [Console]::WriteLine()
    Write-Rule

    $alt      = if ($ActiveMethod -eq 'cert') { 'traditional' } else { 'cert' }
    $altKey   = if ($alt -eq 'cert') { 'C' } else { 'T' }
    $altLabel = if ($alt -eq 'cert') { 'Certificate' } else { 'Device code' }

    [Console]::Write($indentL1)
    [Console]::Write($global:GREX365_AnsiFgDarkGray + "$([char]0x2191)$([char]0x2193) navegar   " + $global:GREX365_AnsiReset)
    [Console]::Write($global:GREX365_AnsiFgCyan     + 'Enter '       + $global:GREX365_AnsiReset)
    [Console]::Write($global:GREX365_AnsiFgDarkGray + 'seleccionar   ' + $global:GREX365_AnsiReset)
    [Console]::Write($global:GREX365_AnsiFgCyan     + 'Esc '         + $global:GREX365_AnsiReset)
    [Console]::Write($global:GREX365_AnsiFgDarkGray + 'salir   '     + $global:GREX365_AnsiReset)
    [Console]::Write($global:GREX365_AnsiFgCyan     + $altKey + ' '  + $global:GREX365_AnsiReset)
    [Console]::WriteLine($global:GREX365_AnsiFgDarkGray + 'cambiar a ' + $altLabel + $global:GREX365_AnsiReset)
    [Console]::WriteLine()

    $ItemRowsOut.Value = $rows
}

function Repaint-MenuItem {
    param(
        [Parameter(Mandatory)][int]$Index,
        [Parameter(Mandatory)][int]$SelectedIndex,
        [Parameter(Mandatory)][array]$Items,
        [Parameter(Mandatory)][hashtable]$Rows
    )

    if (-not $Rows.ContainsKey($Index)) { return $false }
    $row = $Rows[$Index]
    if ($row -lt 0) { return $false }
    try {
        [Console]::SetCursorPosition(0, $row)
        [Console]::Write($global:GREX365_AnsiEraseLine)
        $line = Format-MenuItemLine -Index $Index -Item $Items[$Index] -IsSelected ($Index -eq $SelectedIndex)
        [Console]::Write($line)
        return $true
    } catch {
        return $false
    }
}

function Show-MainMenu {
    param(
        [Parameter(Mandatory = $true)][array]$Items,
        [string]$ActiveMethod
    )

    $selected = 0
    $count = $Items.Count

    # Group items by section once. Stable across keystrokes.
    $sections = [ordered]@{}
    for ($i = 0; $i -lt $count; $i++) {
        $item = $Items[$i]
        $section = if ($item -is [hashtable] -and $item.ContainsKey('Section')) { $item.Section } else { 'Operaciones' }
        if (-not $sections.Contains($section)) {
            $sections[$section] = New-Object System.Collections.Generic.List[object]
        }
        [void]$sections[$section].Add([PSCustomObject]@{ Index = $i; Item = $item })
    }

    $sectionOrderBase = @('Operaciones','Configuración')
    $finalOrder = @()
    foreach ($s in $sectionOrderBase) { if ($sections.Contains($s)) { $finalOrder += $s } }
    foreach ($s in $sections.Keys)    { if ($s -notin $sectionOrderBase) { $finalOrder += $s } }

    $itemRows = @{}
    $needsFullRedraw = $true
    $lastWidth = -1
    $cursorWasVisible = $true
    try { $cursorWasVisible = [Console]::CursorVisible } catch {}

    try {
        try { [Console]::CursorVisible = $false } catch {}

        while ($true) {
            $curWidth = Get-ConsoleWidth
            if ($curWidth -ne $lastWidth) {
                $needsFullRedraw = $true
                $lastWidth = $curWidth
            }

            if ($needsFullRedraw) {
                Render-MainMenuFull `
                    -Items $Items `
                    -SelectedIndex $selected `
                    -ActiveMethod $ActiveMethod `
                    -Sections $sections `
                    -SectionOrder $finalOrder `
                    -ItemRowsOut ([ref]$itemRows)
                $needsFullRedraw = $false
            }

            $key = [Console]::ReadKey($true)

            # Drain repeats of the SAME nav direction while key is held.
            # Stop draining as soon as a non-nav key (or different direction) appears
            # so we still respond to Enter/Esc/etc. promptly.
            $navKeys = @('UpArrow','DownArrow','LeftArrow','RightArrow')
            while ($key.Key -in $navKeys -and [Console]::KeyAvailable) {
                $peek = [Console]::ReadKey($true)
                # Apply current key's movement BEFORE peeking; coalesce by accumulating.
                switch ($key.Key) {
                    'UpArrow'    { $selected = if ($selected -gt 0) { $selected - 1 } else { $count - 1 } }
                    'DownArrow'  { $selected = if ($selected -lt ($count - 1)) { $selected + 1 } else { 0 } }
                    'LeftArrow'  { $selected = if ($selected -gt 0) { $selected - 1 } else { $count - 1 } }
                    'RightArrow' { $selected = if ($selected -lt ($count - 1)) { $selected + 1 } else { 0 } }
                }
                $key = $peek
            }

            $prevSelected = $selected
            switch ($key.Key) {
                'UpArrow'    { $selected = if ($selected -gt 0) { $selected - 1 } else { $count - 1 } }
                'DownArrow'  { $selected = if ($selected -lt ($count - 1)) { $selected + 1 } else { 0 } }
                'LeftArrow'  { $selected = if ($selected -gt 0) { $selected - 1 } else { $count - 1 } }
                'RightArrow' { $selected = if ($selected -lt ($count - 1)) { $selected + 1 } else { 0 } }
                'Escape'   { return [PSCustomObject]@{ Action = 'exit';   Value = $null } }
                'D0'       { return [PSCustomObject]@{ Action = 'exit';   Value = $null } }
                'NumPad0'  { return [PSCustomObject]@{ Action = 'exit';   Value = $null } }
                'Enter'    { return [PSCustomObject]@{ Action = 'option'; Value = ($selected + 1) } }
                'T'        { return [PSCustomObject]@{ Action = 'toggle'; Value = 'traditional' } }
                'C'        { return [PSCustomObject]@{ Action = 'toggle'; Value = 'cert' } }
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

            if ($selected -ne $prevSelected) {
                $okA = Repaint-MenuItem -Index $prevSelected -SelectedIndex $selected -Items $Items -Rows $itemRows
                $okB = Repaint-MenuItem -Index $selected     -SelectedIndex $selected -Items $Items -Rows $itemRows
                if (-not ($okA -and $okB)) { $needsFullRedraw = $true }
            }
        }
    } finally {
        try { [Console]::CursorVisible = $cursorWasVisible } catch {}
    }
}

function Show-MethodSelector {
    Show-Header -Title 'GREX365' -Subtitle 'Selección inicial de método de conexión'

    Write-Indent
    Write-Host 'Selecciona el método de autenticación para Microsoft 365.' -ForegroundColor White
    Write-Host ''

    Write-Indent -Level 2
    Write-Host 'A' -NoNewline -ForegroundColor Cyan
    Write-Host '   Certificate (App-only, desatendido)' -ForegroundColor White
    Write-Indent -Level 3
    Write-Host 'Recomendado. Requiere asistente de creación en primera ejecución.' -ForegroundColor DarkGray
    Write-Host ''

    Write-Indent -Level 2
    Write-Host 'B' -NoNewline -ForegroundColor Cyan
    Write-Host '   Device code (delegated)' -ForegroundColor White
    Write-Indent -Level 3
    Write-Host 'Usa la cuenta del operador. Login interactivo cada sesión.' -ForegroundColor DarkGray
    Write-Host ''

    Write-Rule
    Write-Indent
    Write-Host 'La preferencia se guarda. Puedes cambiarla más tarde con C o T.' -ForegroundColor DarkGray
    Write-Host ''

    while ($true) {
        $key = [System.Console]::ReadKey($true)
        switch ($key.Key) {
            'A'        { return 'cert' }
            'B'        { return 'traditional' }
            'D1'       { return 'cert' }
            'D2'       { return 'traditional' }
            'NumPad1'  { return 'cert' }
            'NumPad2'  { return 'traditional' }
            'Escape'   { return $null }
        }
    }
}

