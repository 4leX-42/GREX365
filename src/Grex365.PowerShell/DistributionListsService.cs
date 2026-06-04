using System.Text.Json;
using Grex365.Core.Abstractions;
using Grex365.Core.Groups;
using Grex365.Core.Models;

namespace Grex365.PowerShell;

// Bulk distribution-list creation (exists -> create -> list members -> add members). The EXO
// cmdlets run in an external pwsh host via IExternalExoRunner; the in-proc RunspacePool trips the
// EXO V3 "HttpResponseMessage does not contain GetResponseHeader" bug. The orchestration and
// per-row error accounting stay in C# (testable). CSV-supplied names/emails are interpolated with
// ExternalExoRunner.Lit (single-quote escaping) — never raw — to stay injection-safe.
public sealed class DistributionListsService : IDistributionListsService
{
    private readonly IExternalExoRunner _runner;

    public DistributionListsService(IExternalExoRunner runner)
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
            var rawName = grp.Key.Trim();
            string groupName;
            string groupEmail;
            if (rawName.Contains('@'))
            {
                groupEmail = rawName;
                groupName = rawName.Split('@', 2)[0];
            }
            else
            {
                groupName = rawName;
                groupEmail = $"{rawName}@{cleanDomain}";
            }

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

            // Best-effort: a members-listing failure leaves an empty set, so add proceeds.
            var existingMembers = await ListMembersAsync(groupEmail, cancellationToken).ConfigureAwait(false);

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
        var body = $$"""
            $g = Get-DistributionGroup -Identity {{ExternalExoRunner.Lit(groupEmail)}} -ErrorAction SilentlyContinue
            Write-Output ('{{ExternalExoRunner.JsonMarker}}' + (([PSCustomObject]@{ Found = [bool]$g }) | ConvertTo-Json -Compress))
            """;
        var json = await _runner.RunAsync(body, progress, ct).ConfigureAwait(false);
        return ParseFound(json);
    }

    private async Task CreateAsync(string groupName, string groupEmail, IProgress<LogEntry>? progress, CancellationToken ct)
    {
        var body = $$"""
            New-DistributionGroup -Name {{ExternalExoRunner.Lit(groupName)}} -PrimarySmtpAddress {{ExternalExoRunner.Lit(groupEmail)}} -Type Distribution -ErrorAction Stop | Out-Null
            Write-Output ('{{ExternalExoRunner.JsonMarker}}' + (([PSCustomObject]@{ Note = 'DL creada' }) | ConvertTo-Json -Compress))
            """;
        // RunAsync throws InvalidOperationException with the real EXO message on failure.
        await _runner.RunAsync(body, progress, ct).ConfigureAwait(false);
    }

    private async Task<HashSet<string>> ListMembersAsync(string groupEmail, CancellationToken ct)
    {
        var body = $$"""
            $members = @(Get-DistributionGroupMember -Identity {{ExternalExoRunner.Lit(groupEmail)}} -ResultSize Unlimited -ErrorAction SilentlyContinue)
            $out = New-Object System.Collections.Generic.List[object]
            foreach ($m in $members) {
                if ($m.PrimarySmtpAddress) { $out.Add([PSCustomObject]@{ Smtp = [string]$m.PrimarySmtpAddress }) }
            }
            Write-Output ('{{ExternalExoRunner.JsonMarker}}' + ($out | ConvertTo-Json -Compress -AsArray))
            """;
        try
        {
            var json = await _runner.RunAsync(body, progress: null, ct).ConfigureAwait(false);
            return ParseMembers(json);
        }
        catch
        {
            // Best-effort, mirrors the legacy SilentlyContinue behaviour: on failure return empty.
            return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }
    }

    private async Task AddMemberAsync(string groupEmail, string member, IProgress<LogEntry>? progress, CancellationToken ct)
    {
        var body = $$"""
            Add-DistributionGroupMember -Identity {{ExternalExoRunner.Lit(groupEmail)}} -Member {{ExternalExoRunner.Lit(member)}} -ErrorAction Stop | Out-Null
            Write-Output ('{{ExternalExoRunner.JsonMarker}}' + (([PSCustomObject]@{ Note = 'ok' }) | ConvertTo-Json -Compress))
            """;
        await _runner.RunAsync(body, progress, ct).ConfigureAwait(false);
    }

    private static bool ParseFound(string? json)
    {
        if (string.IsNullOrWhiteSpace(json)) return false;
        using var doc = JsonDocument.Parse(json);
        return doc.RootElement.ValueKind == JsonValueKind.Object
            && doc.RootElement.TryGetProperty("Found", out var v)
            && v.ValueKind == JsonValueKind.True;
    }

    private static HashSet<string> ParseMembers(string? json)
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(json)) return set;
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        void Add(JsonElement el)
        {
            if (el.ValueKind == JsonValueKind.Object && el.TryGetProperty("Smtp", out var v)
                && v.ValueKind == JsonValueKind.String && v.GetString() is { Length: > 0 } s)
            {
                set.Add(s);
            }
        }

        if (root.ValueKind == JsonValueKind.Array)
        {
            foreach (var el in root.EnumerateArray()) Add(el);
        }
        else if (root.ValueKind == JsonValueKind.Object)
        {
            Add(root); // single-element array collapsed to a lone object by ConvertTo-Json
        }
        return set;
    }
}
