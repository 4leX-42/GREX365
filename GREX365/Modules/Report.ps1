# Report module
# Generates standalone HTML reports for post-operation summaries.
# Single-file output with embedded CSS, no external assets.
# Designed for archive / email / share.

function ConvertTo-HtmlSafeText {
    param([Parameter(ValueFromPipeline = $true)][AllowNull()][AllowEmptyString()]$Text)
    process {
        if ($null -eq $Text) { return '' }
        $s = [string]$Text
        return $s.Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;').Replace('"','&quot;').Replace("'",'&#39;')
    }
}

function Get-StatusCssClass {
    param([string]$Status)
    if ([string]::IsNullOrWhiteSpace($Status)) { return '' }
    $u = $Status.ToUpperInvariant()
    if ($u -match '^(AGREGADO|OK|SUCCESS|CONVERT(IDO|ED))') { return 'status-ok' }
    if ($u -match '(ERROR|FAIL|INVALID)') { return 'status-err' }
    if ($u -match '(WARN|YA_EXIST|DUPLICA|VACIO|NO_RESUEL|SIN_|OMIT)') { return 'status-warn' }
    return ''
}

function Get-ReportCss {
    return @'
<style>
  :root { color-scheme: light; }
  * { box-sizing: border-box; }
  body { font-family: -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif; background:#f6f8fa; color:#1f2328; margin:0; padding:24px; }
  .container { max-width: 1280px; margin: 0 auto; }
  h1 { margin: 0 0 4px 0; color:#0078d4; font-size: 22px; }
  .subtitle { color:#57606a; font-size: 13px; margin-bottom: 20px; }
  .card { background:#fff; border:1px solid #d0d7de; border-radius:8px; padding:16px 20px; margin-bottom:16px; box-shadow: 0 1px 2px rgba(31,35,40,0.04); }
  .meta-grid { display: grid; grid-template-columns: minmax(140px,200px) 1fr; gap: 6px 16px; font-size: 13px; }
  .meta-key { color:#57606a; font-weight: 600; }
  .meta-val { color:#1f2328; word-break: break-word; }
  .meta-val code { background:#f6f8fa; padding:2px 6px; border-radius:4px; font-size:12px; }
  .summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px,1fr)); gap: 12px; }
  .stat { background:#fff; border:1px solid #d0d7de; border-radius:6px; padding: 12px 14px; }
  .stat-key { color:#57606a; font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; font-weight: 600; }
  .stat-val { font-size: 24px; font-weight: 600; color:#1f2328; margin-top: 4px; }
  .stat.ok    .stat-val { color:#1a7f37; }
  .stat.warn  .stat-val { color:#9a6700; }
  .stat.err   .stat-val { color:#cf222e; }
  table { width:100%; border-collapse: collapse; background:#fff; border:1px solid #d0d7de; border-radius:8px; overflow: hidden; }
  th { background:#f6f8fa; text-align:left; padding:10px 12px; border-bottom:1px solid #d0d7de; font-size:12px; text-transform: uppercase; letter-spacing: 0.3px; color:#57606a; }
  td { padding:9px 12px; border-bottom:1px solid #eaeef2; font-size:13px; vertical-align: top; }
  tr:last-child td { border-bottom: none; }
  tr:hover td { background:#f6f8fa; }
  .status-ok   { color:#1a7f37; font-weight:600; }
  .status-warn { color:#9a6700; font-weight:600; }
  .status-err  { color:#cf222e; font-weight:600; }
  .footer { margin-top: 20px; padding-top: 12px; border-top: 1px solid #eaeef2; color:#57606a; font-size:12px; text-align:center; }
  .badge { display:inline-block; padding:2px 8px; border-radius:10px; font-size:11px; font-weight:600; }
  .badge-ok   { background:#dafbe1; color:#1a7f37; }
  .badge-warn { background:#fff8c5; color:#9a6700; }
  .badge-err  { background:#ffebe9; color:#cf222e; }
</style>
'@
}

function New-HtmlReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$Operation,
        [datetime]$StartTime = (Get-Date),
        [datetime]$EndTime   = (Get-Date),
        [hashtable]$Summary,
        [array]$Items,
        [string]$Tenant,
        [string]$TenantId,
        [string]$Account,
        [string]$CorrelationId,
        [string]$LogFile,
        [string[]]$StatusFieldNames = @('Estado','Status','Result','Resultado'),
        [string[]]$WarnFields,
        [string[]]$ErrorFields,
        [string[]]$OkFields,
        [Parameter(Mandatory = $true)][string]$OutputPath
    )

    $duration = $EndTime - $StartTime
    $css = Get-ReportCss

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine('<!DOCTYPE html>')
    [void]$sb.AppendLine('<html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">')
    [void]$sb.Append('<title>').Append((ConvertTo-HtmlSafeText $Title)).AppendLine('</title>')
    [void]$sb.AppendLine($css)
    [void]$sb.AppendLine('</head><body><div class="container">')

    [void]$sb.Append('<h1>').Append((ConvertTo-HtmlSafeText $Title)).AppendLine('</h1>')
    [void]$sb.Append('<div class="subtitle">').Append((ConvertTo-HtmlSafeText $Operation)).AppendLine('</div>')

    [void]$sb.AppendLine('<div class="card"><div class="meta-grid">')
    $metaRows = [ordered]@{
        'Operación'      = $Operation
        'Tenant'         = $Tenant
        'Tenant ID'      = $TenantId
        'Cuenta'         = $Account
        'Correlation ID' = $CorrelationId
        'Inicio'         = $StartTime.ToString('yyyy-MM-dd HH:mm:ss')
        'Fin'            = $EndTime.ToString('yyyy-MM-dd HH:mm:ss')
        'Duración'       = ('{0:N2}s' -f $duration.TotalSeconds)
        'Host'           = ('{0} · {1}' -f $env:COMPUTERNAME, $env:USERNAME)
        'Log'            = $LogFile
    }
    foreach ($k in $metaRows.Keys) {
        $v = $metaRows[$k]
        if ([string]::IsNullOrWhiteSpace([string]$v)) { continue }
        [void]$sb.Append('<div class="meta-key">').Append((ConvertTo-HtmlSafeText $k)).Append('</div>')
        if ($k -eq 'Correlation ID') {
            [void]$sb.Append('<div class="meta-val"><code>').Append((ConvertTo-HtmlSafeText $v)).Append('</code></div>')
        } else {
            [void]$sb.Append('<div class="meta-val">').Append((ConvertTo-HtmlSafeText $v)).Append('</div>')
        }
    }
    [void]$sb.AppendLine('</div></div>')

    if ($Summary -and $Summary.Count -gt 0) {
        [void]$sb.AppendLine('<div class="card"><div class="summary">')
        foreach ($k in $Summary.Keys) {
            $cls = 'stat'
            if ($OkFields    -and ($k -in $OkFields))    { $cls = 'stat ok' }
            elseif ($ErrorFields -and ($k -in $ErrorFields)) { $cls = 'stat err' }
            elseif ($WarnFields  -and ($k -in $WarnFields))  { $cls = 'stat warn' }
            [void]$sb.Append('<div class="').Append($cls).Append('">')
            [void]$sb.Append('<div class="stat-key">').Append((ConvertTo-HtmlSafeText $k)).Append('</div>')
            [void]$sb.Append('<div class="stat-val">').Append((ConvertTo-HtmlSafeText $Summary[$k])).Append('</div>')
            [void]$sb.AppendLine('</div>')
        }
        [void]$sb.AppendLine('</div></div>')
    }

    if ($Items -and $Items.Count -gt 0) {
        $first = $Items[0]
        $cols = @($first.PSObject.Properties.Name)

        [void]$sb.AppendLine('<div class="card" style="padding:0;">')
        [void]$sb.AppendLine('<table><thead><tr>')
        foreach ($c in $cols) {
            [void]$sb.Append('<th>').Append((ConvertTo-HtmlSafeText $c)).Append('</th>')
        }
        [void]$sb.AppendLine('</tr></thead><tbody>')

        foreach ($item in $Items) {
            [void]$sb.AppendLine('<tr>')
            foreach ($c in $cols) {
                $val = $item.$c
                $cls = ''
                if ($c -in $StatusFieldNames) {
                    $cls = Get-StatusCssClass -Status ([string]$val)
                }
                $attr = if ($cls) { ' class="' + $cls + '"' } else { '' }
                [void]$sb.Append('<td').Append($attr).Append('>').Append((ConvertTo-HtmlSafeText $val)).Append('</td>')
            }
            [void]$sb.AppendLine('</tr>')
        }
        [void]$sb.AppendLine('</tbody></table></div>')
    }

    [void]$sb.Append('<div class="footer">GREX365 · informe generado ').Append((Get-Date -Format 'yyyy-MM-dd HH:mm:ss')).AppendLine('</div>')
    [void]$sb.AppendLine('</div></body></html>')

    $folder = Split-Path -Parent $OutputPath
    if ($folder -and -not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
    Set-Content -LiteralPath $OutputPath -Value $sb.ToString() -Encoding UTF8
    return $OutputPath
}

# Convenience wrapper: scripts can call this without wiring tenant / correlationId / log
# manually. Reuses current audit context + session state. Returns path of HTML written.
function Publish-OperationReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$Operation,
        [datetime]$StartTime = (Get-Date),
        [datetime]$EndTime   = (Get-Date),
        [hashtable]$Summary,
        [array]$Items,
        [string]$OutputFolder,
        [string]$LogFile,
        [string[]]$WarnFields,
        [string[]]$ErrorFields,
        [string[]]$OkFields
    )

    $tenant = $null; $tenantId = $null; $account = $null
    try {
        $s = Get-SessionState
        $tenant   = [string]$s.TenantDomain
        $tenantId = [string]$s.TenantId
        $account  = [string]$s.Account
    } catch {}

    $correlationId = $null
    try { $correlationId = Get-AuditCorrelationId } catch {}

    if (-not $OutputFolder) {
        $OutputFolder = Join-Path $global:GREX365_BasePath 'logs\reports'
    }
    $safeOp = ($Operation -replace '[^a-zA-Z0-9_-]', '_')
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $file = Join-Path $OutputFolder ("$safeOp" + "_$stamp.html")

    return (New-HtmlReport `
        -Title $Title `
        -Operation $Operation `
        -StartTime $StartTime `
        -EndTime $EndTime `
        -Summary $Summary `
        -Items $Items `
        -Tenant $tenant `
        -TenantId $tenantId `
        -Account $account `
        -CorrelationId $correlationId `
        -LogFile $LogFile `
        -WarnFields $WarnFields `
        -ErrorFields $ErrorFields `
        -OkFields $OkFields `
        -OutputPath $file)
}
