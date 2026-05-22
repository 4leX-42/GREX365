using System.Text.Json;
using Grex365.Core.Models;

namespace Grex365.Core.Audit;

public static class AuditReportJsonBuilder
{
    private static readonly JsonSerializerOptions DefaultOptions = new()
    {
        WriteIndented = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public static string Build(
        IEnumerable<AuditFinding> findings,
        AuditReportContext context,
        JsonSerializerOptions? options = null)
    {
        var list = (findings ?? Array.Empty<AuditFinding>()).ToList();
        var errors = list.Count(f => IsSeverity(f, "ERROR"));
        var warns = list.Count(f => IsSeverity(f, "WARN"));
        var infos = list.Count(f => IsSeverity(f, "INFO"));

        var report = new AuditReportEnvelope(
            Schema: SchemaVersion,
            Title: context.Title,
            GeneratedAt: context.GeneratedAt,
            TenantDomain: context.TenantDomain,
            GeneratedBy: context.GeneratedBy,
            Counts: new AuditReportCounts(errors, warns, infos, list.Count),
            Findings: list);

        return JsonSerializer.Serialize(report, options ?? DefaultOptions);
    }

    public static AuditReportEnvelope? Parse(string json, JsonSerializerOptions? options = null)
    {
        if (string.IsNullOrWhiteSpace(json))
        {
            return null;
        }
        return JsonSerializer.Deserialize<AuditReportEnvelope>(json, options);
    }

    public const string SchemaVersion = "grex365.audit.v1";

    private static bool IsSeverity(AuditFinding f, string target) =>
        string.Equals(f.Severity, target, StringComparison.OrdinalIgnoreCase);
}

public sealed record AuditReportEnvelope(
    string Schema,
    string Title,
    DateTime GeneratedAt,
    string? TenantDomain,
    string? GeneratedBy,
    AuditReportCounts Counts,
    IReadOnlyList<AuditFinding> Findings);

public sealed record AuditReportCounts(
    int Errors,
    int Warnings,
    int Info,
    int Total);
