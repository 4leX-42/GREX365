using System.Globalization;
using System.Text;
using Grex365.Core.Models;

namespace Grex365.Core.Audit;

public static class AuditReportHtmlBuilder
{
    public static string Build(
        IEnumerable<AuditFinding> findings,
        AuditReportContext context)
    {
        var list = (findings ?? Array.Empty<AuditFinding>()).ToList();

        var errors = list.Where(f => IsSeverity(f, "ERROR")).ToList();
        var warns = list.Where(f => IsSeverity(f, "WARN")).ToList();
        var infos = list.Where(f => IsSeverity(f, "INFO")).ToList();

        var sb = new StringBuilder(capacity: 4096);
        sb.Append("<!DOCTYPE html>\n<html lang=\"es\"><head><meta charset=\"utf-8\">");
        sb.Append("<title>").Append(Esc(context.Title)).Append("</title>");
        sb.Append("<style>").Append(InlineCss).Append("</style>");
        sb.Append("</head><body>");

        sb.Append("<header><h1>").Append(Esc(context.Title)).Append("</h1>");
        sb.Append("<div class=\"meta\">");
        sb.Append("<span><strong>Generado:</strong> ").Append(Esc(context.GeneratedAt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture))).Append("</span>");
        if (!string.IsNullOrWhiteSpace(context.TenantDomain))
        {
            sb.Append("<span><strong>Tenant:</strong> ").Append(Esc(context.TenantDomain!)).Append("</span>");
        }
        if (!string.IsNullOrWhiteSpace(context.GeneratedBy))
        {
            sb.Append("<span><strong>Operador:</strong> ").Append(Esc(context.GeneratedBy!)).Append("</span>");
        }
        sb.Append("</div></header>");

        sb.Append("<section class=\"summary\">");
        AppendPill(sb, "ERROR", errors.Count, "error");
        AppendPill(sb, "WARN", warns.Count, "warn");
        AppendPill(sb, "INFO", infos.Count, "info");
        AppendPill(sb, "TOTAL", list.Count, "total");
        sb.Append("</section>");

        if (list.Count == 0)
        {
            sb.Append("<section class=\"empty\">Sin hallazgos registrados.</section>");
        }
        else
        {
            AppendTable(sb, "Errores", errors, "error");
            AppendTable(sb, "Advertencias", warns, "warn");
            AppendTable(sb, "Informativos", infos, "info");
        }

        sb.Append("<footer>GREX365 · informe de auditoría</footer>");
        sb.Append("</body></html>");
        return sb.ToString();
    }

    private static void AppendTable(StringBuilder sb, string heading, IReadOnlyList<AuditFinding> rows, string cssClass)
    {
        if (rows.Count == 0)
        {
            return;
        }
        sb.Append("<section class=\"group ").Append(cssClass).Append("\">");
        sb.Append("<h2>").Append(Esc(heading)).Append(" <span class=\"count\">(").Append(rows.Count).Append(")</span></h2>");
        sb.Append("<table><thead><tr><th>Categoría</th><th>Identidad</th><th>Detalle</th></tr></thead><tbody>");
        foreach (var f in rows)
        {
            sb.Append("<tr>");
            sb.Append("<td>").Append(Esc(f.Category)).Append("</td>");
            sb.Append("<td>").Append(Esc(f.Identity)).Append("</td>");
            sb.Append("<td>").Append(Esc(f.Detail)).Append("</td>");
            sb.Append("</tr>");
        }
        sb.Append("</tbody></table></section>");
    }

    private static void AppendPill(StringBuilder sb, string label, int count, string cssClass)
    {
        sb.Append("<div class=\"pill ").Append(cssClass).Append("\">");
        sb.Append("<span class=\"label\">").Append(Esc(label)).Append("</span>");
        sb.Append("<span class=\"count\">").Append(count).Append("</span>");
        sb.Append("</div>");
    }

    private static bool IsSeverity(AuditFinding f, string target) =>
        string.Equals(f.Severity, target, StringComparison.OrdinalIgnoreCase);

    public static string Esc(string? raw)
    {
        if (string.IsNullOrEmpty(raw))
        {
            return string.Empty;
        }
        var sb = new StringBuilder(raw.Length);
        foreach (var ch in raw)
        {
            switch (ch)
            {
                case '<': sb.Append("&lt;"); break;
                case '>': sb.Append("&gt;"); break;
                case '&': sb.Append("&amp;"); break;
                case '"': sb.Append("&quot;"); break;
                case '\'': sb.Append("&#39;"); break;
                default: sb.Append(ch); break;
            }
        }
        return sb.ToString();
    }

    private const string InlineCss = @"
:root { color-scheme: light dark; }
* { box-sizing: border-box; }
body { font-family: 'Segoe UI', system-ui, -apple-system, sans-serif; margin: 0; padding: 0; background: #f6f7fb; color: #1c1f24; }
header { padding: 32px 40px 16px; border-bottom: 1px solid #e4e6ea; background: #fff; }
header h1 { margin: 0 0 8px; font-size: 24px; font-weight: 600; }
.meta { display: flex; flex-wrap: wrap; gap: 18px; font-size: 13px; color: #57606a; }
.summary { display: flex; flex-wrap: wrap; gap: 12px; padding: 20px 40px; background: #fff; border-bottom: 1px solid #e4e6ea; }
.pill { display: inline-flex; align-items: center; gap: 10px; padding: 10px 16px; border-radius: 8px; border: 1px solid #d0d7de; background: #fff; }
.pill .label { font-size: 11px; font-weight: 700; letter-spacing: 0.5px; }
.pill .count { font-size: 18px; font-weight: 700; }
.pill.error { border-color: #f97066; background: #fff4f3; color: #b42318; }
.pill.warn { border-color: #f7b955; background: #fffaf0; color: #93540a; }
.pill.info { border-color: #76a9fa; background: #f0f7ff; color: #0b4f9c; }
.pill.total { border-color: #b1bac4; background: #f6f7fb; color: #24292f; }
.group { padding: 20px 40px; }
.group h2 { font-size: 16px; font-weight: 600; margin: 0 0 12px; display: flex; align-items: center; gap: 6px; }
.group h2 .count { font-weight: 500; color: #57606a; font-size: 13px; }
.group.error h2 { color: #b42318; }
.group.warn h2 { color: #93540a; }
.group.info h2 { color: #0b4f9c; }
table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 8px; overflow: hidden; border: 1px solid #e4e6ea; }
thead { background: #f6f7fb; }
th { text-align: left; padding: 10px 12px; font-size: 12px; font-weight: 600; color: #57606a; text-transform: uppercase; letter-spacing: 0.4px; }
td { padding: 10px 12px; font-size: 13px; border-top: 1px solid #eef0f3; vertical-align: top; }
tr:hover td { background: #fafbfc; }
.empty { padding: 40px; text-align: center; color: #57606a; font-style: italic; }
footer { padding: 18px 40px; text-align: center; font-size: 11px; color: #8b95a1; border-top: 1px solid #e4e6ea; background: #fff; }
@media (prefers-color-scheme: dark) {
    body { background: #1a1d23; color: #e6e8eb; }
    header, .summary, table, footer { background: #21252b; border-color: #2c313a; }
    .pill { background: #21252b; border-color: #2c313a; color: #e6e8eb; }
    .pill.error { background: #2c1d1e; }
    .pill.warn { background: #2d2618; }
    .pill.info { background: #1a2333; }
    thead { background: #1f242b; }
    th { color: #aab2bd; }
    td { border-top-color: #2c313a; }
    tr:hover td { background: #262b33; }
    .meta { color: #aab2bd; }
}
";
}

public sealed record AuditReportContext(
    string Title,
    DateTime GeneratedAt,
    string? TenantDomain = null,
    string? GeneratedBy = null);
