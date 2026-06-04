using System.Text.Json;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.PowerShell;

// Lists Exchange Online transport (mail flow) rules. Runs Get-TransportRule in an external pwsh
// host: the in-proc RunspacePool trips the EXO V3 "HttpResponseMessage does not contain
// GetResponseHeader" bug when reading rule PSObject properties.
public sealed class MailFlowRulesService : IMailFlowRulesService
{
    private readonly IExternalExoRunner _runner;

    public MailFlowRulesService(IExternalExoRunner runner)
    {
        _runner = runner;
    }

    public async Task<IReadOnlyList<TransportRuleSummary>> GetRulesAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var body = $$"""
            $rules = @(Get-TransportRule -ErrorAction Stop)
            $out = New-Object System.Collections.Generic.List[object]
            foreach ($r in $rules) {
                $out.Add([PSCustomObject]@{
                    Name        = [string]$r.Name
                    State       = [string]$r.State
                    Priority    = [int]$r.Priority
                    Mode        = [string]$r.Mode
                    Description = [string]$r.Description
                })
            }
            Write-Output ('{{ExternalExoRunner.JsonMarker}}' + ($out | ConvertTo-Json -Compress -AsArray))
            """;
        var json = await _runner.RunAsync(body, progress, cancellationToken).ConfigureAwait(false);

        return ParseRules(json)
            .OrderBy(r => r.Priority)
            .ThenBy(r => r.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static IReadOnlyList<TransportRuleSummary> ParseRules(string? json)
    {
        var list = new List<TransportRuleSummary>();
        if (string.IsNullOrWhiteSpace(json)) return list;
        using var doc = JsonDocument.Parse(json);
        if (doc.RootElement.ValueKind != JsonValueKind.Array) return list;
        foreach (var el in doc.RootElement.EnumerateArray())
        {
            string? S(string n) => el.TryGetProperty(n, out var v) && v.ValueKind == JsonValueKind.String && v.GetString() is { Length: > 0 } s ? s : null;
            int I(string n) => el.TryGetProperty(n, out var v) && v.ValueKind == JsonValueKind.Number && v.TryGetInt32(out var i) ? i : 0;
            list.Add(new TransportRuleSummary(
                Name: S("Name") ?? "(sin nombre)",
                State: S("State") ?? "(?)",
                Priority: I("Priority"),
                Mode: S("Mode") ?? "(?)",
                Description: S("Description")));
        }
        return list;
    }
}
