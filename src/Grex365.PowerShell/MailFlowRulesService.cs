using System.Globalization;
using System.Management.Automation;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.PowerShell;

public sealed class MailFlowRulesService : IMailFlowRulesService
{
    private readonly IPowerShellRunner _runner;

    public MailFlowRulesService(IPowerShellRunner runner)
    {
        _runner = runner;
    }

    public async Task<IReadOnlyList<TransportRuleSummary>> GetRulesAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        // Bug ExchangeOnlineManagement (PS7 + EXO 3.x): si los properties se acceden lazy en
        // un ForEach-Object posterior, EXO llama internamente HttpWebResponse.GetResponseHeader
        // sobre un HttpResponseMessage -> NullReference/MethodNotFound.
        // Workaround: forzar evaluación eager con Select-Object proyectando las propiedades
        // directamente. Devuelve PSObjects con propiedades como strings/ints.
        const string script = """
            param()
            Get-TransportRule -ErrorAction Stop |
                Select-Object @{N='Name';E={[string]$_.Name}},
                              @{N='State';E={[string]$_.State}},
                              @{N='Priority';E={[int]$_.Priority}},
                              @{N='Mode';E={[string]$_.Mode}},
                              @{N='Description';E={[string]$_.Description}}
            """;

        var result = await _runner.RunAsync(
            script,
            parameters: null,
            progress,
            cancellationToken).ConfigureAwait(false);

        if (!result.Success)
        {
            throw new InvalidOperationException("Get-TransportRule falló: " + string.Join("; ", result.Errors));
        }

        var rules = new List<TransportRuleSummary>(result.Output.Count);
        foreach (var obj in result.Output)
        {
            if (obj is not PSObject ps)
            {
                continue;
            }
            rules.Add(new TransportRuleSummary(
                Name: GetString(ps, "Name") ?? "(sin nombre)",
                State: GetString(ps, "State") ?? "(?)",
                Priority: GetInt(ps, "Priority"),
                Mode: GetString(ps, "Mode") ?? "(?)",
                Description: GetString(ps, "Description")));
        }

        return rules
            .OrderBy(r => r.Priority)
            .ThenBy(r => r.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string? GetString(PSObject obj, string property)
    {
        var value = obj.Properties[property]?.Value;
        if (value is null)
        {
            return null;
        }
        var s = Convert.ToString(value, CultureInfo.InvariantCulture);
        return string.IsNullOrWhiteSpace(s) ? null : s;
    }

    private static int GetInt(PSObject obj, string property)
    {
        var value = obj.Properties[property]?.Value;
        return value switch
        {
            int i => i,
            long l => (int)l,
            null => 0,
            _ => int.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), out var parsed) ? parsed : 0,
        };
    }
}
