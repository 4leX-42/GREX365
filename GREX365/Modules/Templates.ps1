# Templates module
# Loads corporate message templates from GREX365/templates/*.json
# Placeholder syntax: {key} -> hashtable value via Expand-Template.
# Used by Offboarding wizard for auto-reply / handover messages.

function Get-TemplatesFolder {
    if (-not $global:GREX365_BasePath) { throw 'BasePath no inicializado.' }
    $folder = Join-Path $global:GREX365_BasePath 'templates'
    if (-not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
    return $folder
}

function Get-AvailableTemplates {
    $folder = Get-TemplatesFolder
    $files = Get-ChildItem -LiteralPath $folder -Filter *.json -File -ErrorAction SilentlyContinue
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($f in $files) {
        try {
            $raw = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
            $obj = $raw | ConvertFrom-Json -ErrorAction Stop
            $list.Add([PSCustomObject]@{
                Name        = if ($obj.name) { [string]$obj.name } else { $f.BaseName }
                File        = $f.FullName
                Description = [string]$obj.description
                Category    = if ($obj.category) { [string]$obj.category } else { 'general' }
                Lang        = if ($obj.lang) { [string]$obj.lang } else { 'es' }
                Subject     = [string]$obj.subject
                BodyText    = [string]$obj.bodyText
                BodyHtml    = [string]$obj.bodyHtml
            })
        } catch {}
    }
    return $list
}

function Get-TemplateByName {
    param([Parameter(Mandatory = $true)][string]$Name)
    $all = Get-AvailableTemplates
    $hit = $all | Where-Object { $_.Name -eq $Name } | Select-Object -First 1
    if (-not $hit) { throw "Template no encontrada: $Name" }
    return $hit
}

function Expand-Template {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text,
        [hashtable]$Values
    )
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    if (-not $Values) { return $Text }
    $out = $Text
    foreach ($k in $Values.Keys) {
        $token = '{' + $k + '}'
        $val = if ($null -eq $Values[$k]) { '' } else { [string]$Values[$k] }
        $out = $out.Replace($token, $val)
    }
    return $out
}

function Render-Template {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [hashtable]$Values
    )
    $t = Get-TemplateByName -Name $Name
    return [PSCustomObject]@{
        Name     = $t.Name
        Subject  = Expand-Template -Text $t.Subject -Values $Values
        BodyText = Expand-Template -Text $t.BodyText -Values $Values
        BodyHtml = Expand-Template -Text $t.BodyHtml -Values $Values
        Lang     = $t.Lang
    }
}

function Select-TemplateInteractive {
    param([string]$Category)

    $all = Get-AvailableTemplates
    if ($Category) { $all = @($all | Where-Object { $_.Category -eq $Category }) }
    if (-not $all -or $all.Count -eq 0) {
        Write-Log "No hay plantillas disponibles (category=$Category)" -Level WARN -Source 'Templates'
        return $null
    }

    Show-Section -Title 'Plantillas disponibles'
    $i = 0
    foreach ($t in $all) {
        $i++
        Write-Indent -Level 2
        Write-Host ('{0,2}  ' -f $i) -NoNewline -ForegroundColor DarkGray
        Write-Host $t.Name -NoNewline -ForegroundColor Cyan
        Write-Host ('  [' + $t.Category + '/' + $t.Lang + ']  ') -NoNewline -ForegroundColor DarkGray
        Write-Host $t.Description -ForegroundColor Gray
    }
    Write-Host ''

    $opt = Read-Input -Prompt 'Número de plantilla (vacío = ninguna)'
    if ([string]::IsNullOrWhiteSpace($opt)) { return $null }
    if ($opt -match '^\d+$') {
        $n = [int]$opt
        if ($n -ge 1 -and $n -le $all.Count) { return $all[$n - 1] }
    }
    return $null
}
