using System.Globalization;
using System.Text;
using System.Text.Json;
using Grex365.Core.Models;

namespace Grex365.Core.Offboarding;

/// <summary>
/// Pure (IO-free) serialisers that turn offboarding run results into an auditable artifact
/// — JSON for machine processing, CSV (one row per step) for handing to HR / a ticket.
/// Kept side-effect-free so it is fully unit-testable; callers own where the bytes go.
/// </summary>
public static class OffboardingReport
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    public static string ToJson(IEnumerable<OffboardingResult> results)
    {
        ArgumentNullException.ThrowIfNull(results);
        var payload = results.Select(r => new
        {
            upn = r.Upn,
            success = r.Success,
            dryRun = r.DryRun,
            startedAt = r.StartedAt,
            endedAt = r.EndedAt,
            steps = r.Steps.Select(s => new
            {
                name = s.Name,
                status = s.Status,
                detail = s.Detail,
                at = s.At,
            }),
        });
        return JsonSerializer.Serialize(payload, JsonOptions);
    }

    public static string ToCsv(IEnumerable<OffboardingResult> results)
    {
        ArgumentNullException.ThrowIfNull(results);
        var sb = new StringBuilder();
        sb.AppendLine("upn,success,dryRun,startedAt,endedAt,step,status,detail,at");
        foreach (var r in results)
        {
            var started = Iso(r.StartedAt);
            var ended = Iso(r.EndedAt);
            // Emit a row per step; a run with no steps still gets a single summary row so it
            // isn't silently dropped from the export.
            if (r.Steps.Count == 0)
            {
                sb.AppendLine(Row(r.Upn, r.Success, r.DryRun, started, ended, "", "", "", ""));
                continue;
            }
            foreach (var s in r.Steps)
            {
                sb.AppendLine(Row(r.Upn, r.Success, r.DryRun, started, ended, s.Name, s.Status, s.Detail, Iso(s.At)));
            }
        }
        return sb.ToString();
    }

    private static string Row(string upn, bool success, bool dry, string started, string ended,
        string step, string status, string detail, string at) =>
        string.Join(",", new[]
        {
            Esc(upn), success ? "true" : "false", dry ? "true" : "false", Esc(started), Esc(ended),
            Esc(step), Esc(status), Esc(detail), Esc(at),
        });

    private static string Iso(DateTimeOffset? value) =>
        value?.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) ?? string.Empty;

    // RFC-4180 escaping: wrap in quotes and double embedded quotes when the field contains a
    // comma, quote, or newline.
    private static string Esc(string? value)
    {
        var v = value ?? string.Empty;
        if (v.IndexOfAny(new[] { ',', '"', '\n', '\r' }) < 0)
        {
            return v;
        }
        return "\"" + v.Replace("\"", "\"\"") + "\"";
    }
}
