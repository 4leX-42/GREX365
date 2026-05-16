using Grex365.Core.Abstractions;
using Grex365.Core.Groups;
using Grex365.Core.Models;

namespace Grex365.PowerShell;

public sealed class DistributionListsService : IDistributionListsService
{
    private readonly IPowerShellRunner _runner;

    public DistributionListsService(IPowerShellRunner runner)
    {
        _runner = runner;
    }

    public async Task<IReadOnlyList<BulkGroupResult>> CreateFromRowsAsync(
        IReadOnlyList<BulkGroupRow> rows,
        string domain,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var results = new List<BulkGroupResult>();
        var cleanDomain = (domain ?? string.Empty).TrimStart('@').Trim();
        if (string.IsNullOrEmpty(cleanDomain))
        {
            throw new ArgumentException("Dominio requerido.", nameof(domain));
        }

        var groups = rows.GroupBy(r => r.GroupName, StringComparer.OrdinalIgnoreCase);
        foreach (var grp in groups)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var groupName = grp.Key.Trim();
            var groupEmail = $"{groupName}@{cleanDomain}";

            bool exists;
            try
            {
                exists = await ExistsAsync(groupEmail, progress, cancellationToken).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                results.Add(new BulkGroupResult(groupName, groupEmail, "Error", null, "Lookup fallido: " + ex.Message));
                continue;
            }

            if (!exists)
            {
                try
                {
                    await CreateAsync(groupName, groupEmail, progress, cancellationToken).ConfigureAwait(false);
                    results.Add(new BulkGroupResult(groupName, groupEmail, "Created", null, "DL creada"));
                }
                catch (Exception ex)
                {
                    results.Add(new BulkGroupResult(groupName, groupEmail, "Error", null, "Creación fallida: " + ex.Message));
                    continue;
                }
            }
            else
            {
                results.Add(new BulkGroupResult(groupName, groupEmail, "Skipped", null, "Ya existía"));
            }

            HashSet<string> existingMembers;
            try
            {
                existingMembers = await ListMembersAsync(groupEmail, cancellationToken).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                results.Add(new BulkGroupResult(groupName, groupEmail, "Error", null, "Get-members fallido: " + ex.Message));
                continue;
            }

            foreach (var row in grp)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var email = row.Email.Trim();
                if (!BulkGroupRowPreprocessor.IsEmail(email))
                {
                    results.Add(new BulkGroupResult(groupName, groupEmail, "Error", email, "Email inválido"));
                    continue;
                }
                if (existingMembers.Contains(email))
                {
                    results.Add(new BulkGroupResult(groupName, groupEmail, "MemberSkipped", email, "Ya pertenece"));
                    continue;
                }
                try
                {
                    await AddMemberAsync(groupEmail, email, progress, cancellationToken).ConfigureAwait(false);
                    existingMembers.Add(email);
                    results.Add(new BulkGroupResult(groupName, groupEmail, "MemberAdded", email, "OK"));
                }
                catch (Exception ex)
                {
                    results.Add(new BulkGroupResult(groupName, groupEmail, "Error", email, "Add fallido: " + ex.Message));
                }
            }
        }

        return results;
    }

    private async Task<bool> ExistsAsync(string groupEmail, IProgress<LogEntry>? progress, CancellationToken ct)
    {
        const string script = """
            $g = Get-DistributionGroup -Identity $Identity -ErrorAction SilentlyContinue
            [PSCustomObject]@{ Found = [bool]$g }
            """;
        var result = await _runner.RunAsync(
            script,
            new Dictionary<string, object?> { ["Identity"] = groupEmail },
            progress,
            ct).ConfigureAwait(false);
        if (!result.Success || result.Output.Count == 0)
        {
            return false;
        }
        return ReadBool(result.Output[0], "Found");
    }

    private async Task CreateAsync(string groupName, string groupEmail, IProgress<LogEntry>? progress, CancellationToken ct)
    {
        const string script = """
            New-DistributionGroup -Name $Name -PrimarySmtpAddress $Smtp -Type Distribution -ErrorAction Stop | Out-Null
            Write-Information "DL creada: $Smtp"
            """;
        var result = await _runner.RunAsync(
            script,
            new Dictionary<string, object?>
            {
                ["Name"] = groupName,
                ["Smtp"] = groupEmail
            },
            progress,
            ct).ConfigureAwait(false);
        if (!result.Success)
        {
            throw new InvalidOperationException(string.Join("; ", result.Errors));
        }
    }

    private async Task<HashSet<string>> ListMembersAsync(string groupEmail, CancellationToken ct)
    {
        const string script = """
            $members = @(Get-DistributionGroupMember -Identity $Identity -ResultSize Unlimited -ErrorAction SilentlyContinue)
            foreach ($m in $members) {
                if ($m.PrimarySmtpAddress) {
                    [PSCustomObject]@{ Smtp = [string]$m.PrimarySmtpAddress }
                }
            }
            """;
        var result = await _runner.RunAsync(
            script,
            new Dictionary<string, object?> { ["Identity"] = groupEmail },
            progress: null,
            ct).ConfigureAwait(false);

        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (!result.Success)
        {
            return set;
        }
        foreach (var raw in result.Output)
        {
            var smtp = ReadString(raw, "Smtp");
            if (!string.IsNullOrWhiteSpace(smtp))
            {
                set.Add(smtp);
            }
        }
        return set;
    }

    private async Task AddMemberAsync(string groupEmail, string member, IProgress<LogEntry>? progress, CancellationToken ct)
    {
        const string script = """
            Add-DistributionGroupMember -Identity $Identity -Member $Member -ErrorAction Stop | Out-Null
            """;
        var result = await _runner.RunAsync(
            script,
            new Dictionary<string, object?>
            {
                ["Identity"] = groupEmail,
                ["Member"] = member
            },
            progress,
            ct).ConfigureAwait(false);
        if (!result.Success)
        {
            throw new InvalidOperationException(string.Join("; ", result.Errors));
        }
    }

    private static bool ReadBool(object? raw, string prop)
    {
        if (raw is System.Management.Automation.PSObject ps)
        {
            var v = ps.Properties[prop]?.Value;
            return v is bool b ? b : bool.TryParse(v?.ToString(), out var parsed) && parsed;
        }
        var t = raw?.GetType();
        var pv = t?.GetProperty(prop)?.GetValue(raw);
        return pv is bool bb ? bb : bool.TryParse(pv?.ToString(), out var p) && p;
    }

    private static string ReadString(object? raw, string prop)
    {
        if (raw is System.Management.Automation.PSObject ps)
        {
            return ps.Properties[prop]?.Value?.ToString() ?? string.Empty;
        }
        var t = raw?.GetType();
        return t?.GetProperty(prop)?.GetValue(raw)?.ToString() ?? string.Empty;
    }
}
