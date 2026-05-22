using Grex365.Core.Abstractions;
using Grex365.Core.Audit;
using Grex365.Core.Models;
using System.Management.Automation;

namespace Grex365.PowerShell;

public sealed class ExoForwardingAuditService : IExoForwardingAuditService
{
    private readonly IPowerShellRunner _runner;

    public ExoForwardingAuditService(IPowerShellRunner runner)
    {
        _runner = runner;
    }

    public async Task<IReadOnlyList<AuditFinding>> ScanExternalForwardingAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        progress?.Report(LogEntry.Info("ExoAudit", "Get-AcceptedDomain..."));

        const string domainsScript = """
            param()
            Get-AcceptedDomain -ErrorAction Stop |
                Select-Object -ExpandProperty DomainName
            """;
        var domainsResult = await _runner.RunAsync(domainsScript, parameters: null, progress, cancellationToken).ConfigureAwait(false);
        if (!domainsResult.Success)
        {
            throw new InvalidOperationException("Get-AcceptedDomain falló: " + string.Join("; ", domainsResult.Errors));
        }

        var acceptedDomains = domainsResult.Output
            .Select(o => o?.ToString())
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .Select(s => s!)
            .ToList();

        progress?.Report(LogEntry.Info("ExoAudit",
            $"Aceptados {acceptedDomains.Count} dominios. Escaneando buzones con forwarding..."));

        const string mailboxesScript = """
            param()
            Get-Mailbox -ResultSize Unlimited -ErrorAction Stop |
                Where-Object { $_.ForwardingSmtpAddress -or $_.ForwardingAddress } |
                Select-Object UserPrincipalName, ForwardingSmtpAddress, ForwardingAddress
            """;
        var mbResult = await _runner.RunAsync(mailboxesScript, parameters: null, progress, cancellationToken).ConfigureAwait(false);
        if (!mbResult.Success)
        {
            throw new InvalidOperationException("Get-Mailbox falló: " + string.Join("; ", mbResult.Errors));
        }

        var rows = mbResult.Output
            .OfType<PSObject>()
            .Select(o => new MailboxForwardingRow(
                UserPrincipalName: o.Properties["UserPrincipalName"]?.Value?.ToString() ?? string.Empty,
                ForwardingSmtpAddress: o.Properties["ForwardingSmtpAddress"]?.Value?.ToString(),
                ForwardingAddress: o.Properties["ForwardingAddress"]?.Value?.ToString()))
            .ToList();

        progress?.Report(LogEntry.Info("ExoAudit",
            $"Buzones con forwarding: {rows.Count}. Aplicando filtro de dominios..."));

        var findings = MailboxForwardingAnalyzer.Analyze(rows, acceptedDomains);
        progress?.Report(LogEntry.Ok("ExoAudit",
            $"{findings.Count} forwards externos detectados sobre {rows.Count} buzones."));
        return findings;
    }

    public async Task<IReadOnlyList<AuditFinding>> ScanInboxRulesAsync(
        int maxMailboxes = 200,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        if (maxMailboxes < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(maxMailboxes), "Debe ser >= 1.");
        }

        progress?.Report(LogEntry.Info("ExoAudit", "Get-AcceptedDomain..."));
        const string domainsScript = """
            param()
            Get-AcceptedDomain -ErrorAction Stop |
                Select-Object -ExpandProperty DomainName
            """;
        var domainsResult = await _runner.RunAsync(domainsScript, parameters: null, progress, cancellationToken).ConfigureAwait(false);
        if (!domainsResult.Success)
        {
            throw new InvalidOperationException("Get-AcceptedDomain falló: " + string.Join("; ", domainsResult.Errors));
        }
        var acceptedDomains = domainsResult.Output
            .Select(o => o?.ToString())
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .Select(s => s!)
            .ToList();

        progress?.Report(LogEntry.Info("ExoAudit",
            $"Get-Mailbox -ResultSize {maxMailboxes} para inbox rules..."));

        const string mailboxesScript = """
            param([int]$Top)
            Get-Mailbox -ResultSize $Top -RecipientTypeDetails UserMailbox -ErrorAction Stop |
                Select-Object -ExpandProperty UserPrincipalName
            """;
        var mbResult = await _runner.RunAsync(
            mailboxesScript,
            new Dictionary<string, object?> { ["Top"] = maxMailboxes },
            progress,
            cancellationToken).ConfigureAwait(false);
        if (!mbResult.Success)
        {
            throw new InvalidOperationException("Get-Mailbox falló: " + string.Join("; ", mbResult.Errors));
        }

        var upns = mbResult.Output
            .Select(o => o?.ToString())
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .Select(s => s!)
            .ToList();

        progress?.Report(LogEntry.Info("ExoAudit",
            $"{upns.Count} buzones; iterando Get-InboxRule (puede tardar)..."));

        var allRules = new List<InboxRuleRow>();
        var done = 0;
        foreach (var upn in upns)
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                var rules = await FetchRulesForMailboxAsync(upn, progress, cancellationToken).ConfigureAwait(false);
                allRules.AddRange(rules);
            }
            catch (Exception ex)
            {
                progress?.Report(LogEntry.Warn("ExoAudit", $"Get-InboxRule {upn}: {ex.Message}"));
            }
            done++;
            if (done % 25 == 0)
            {
                progress?.Report(LogEntry.Info("ExoAudit", $"Progreso: {done}/{upns.Count}"));
            }
        }

        var findings = InboxRuleAnalyzer.Analyze(allRules, acceptedDomains);
        progress?.Report(LogEntry.Ok("ExoAudit",
            $"{findings.Count} reglas sospechosas sobre {allRules.Count} reglas / {upns.Count} buzones."));
        return findings;
    }

    private async Task<IReadOnlyList<InboxRuleRow>> FetchRulesForMailboxAsync(
        string upn,
        IProgress<LogEntry>? progress,
        CancellationToken cancellationToken)
    {
        const string script = """
            param([string]$Identity)
            Get-InboxRule -Mailbox $Identity -ErrorAction Stop |
                Select-Object Name, Enabled, DeleteMessage, MoveToFolder,
                              @{N='ForwardTo';E={ if ($_.ForwardTo) { @($_.ForwardTo | ForEach-Object { $_.ToString() }) } else { @() } }},
                              @{N='ForwardAsAttachmentTo';E={ if ($_.ForwardAsAttachmentTo) { @($_.ForwardAsAttachmentTo | ForEach-Object { $_.ToString() }) } else { @() } }},
                              @{N='RedirectTo';E={ if ($_.RedirectTo) { @($_.RedirectTo | ForEach-Object { $_.ToString() }) } else { @() } }},
                              @{N='SubjectContainsWords';E={ if ($_.SubjectContainsWords) { @($_.SubjectContainsWords) } else { @() } }},
                              @{N='BodyContainsWords';E={ if ($_.BodyContainsWords) { @($_.BodyContainsWords) } else { @() } }}
            """;
        var result = await _runner.RunAsync(
            script,
            new Dictionary<string, object?> { ["Identity"] = upn },
            progress,
            cancellationToken).ConfigureAwait(false);
        if (!result.Success)
        {
            throw new InvalidOperationException(string.Join("; ", result.Errors));
        }

        return result.Output
            .OfType<System.Management.Automation.PSObject>()
            .Select(o => new InboxRuleRow(
                MailboxUpn: upn,
                RuleName: o.Properties["Name"]?.Value?.ToString() ?? "(sin nombre)",
                Enabled: o.Properties["Enabled"]?.Value is bool b && b,
                DeleteMessage: o.Properties["DeleteMessage"]?.Value is bool d && d,
                MoveToFolder: o.Properties["MoveToFolder"]?.Value?.ToString(),
                ForwardTo: ToList(o.Properties["ForwardTo"]?.Value),
                ForwardAsAttachmentTo: ToList(o.Properties["ForwardAsAttachmentTo"]?.Value),
                RedirectTo: ToList(o.Properties["RedirectTo"]?.Value),
                SubjectContainsWords: ToList(o.Properties["SubjectContainsWords"]?.Value),
                BodyContainsWords: ToList(o.Properties["BodyContainsWords"]?.Value)))
            .ToList();
    }

    public async Task<(Grex365.Core.Audit.TransportRulesSummary Summary, IReadOnlyList<AuditFinding> Findings)>
        ScanTransportRulesAsync(
            IProgress<LogEntry>? progress = null,
            CancellationToken cancellationToken = default)
    {
        progress?.Report(LogEntry.Info("ExoAudit", "Get-AcceptedDomain..."));
        const string domainsScript = """
            param()
            Get-AcceptedDomain -ErrorAction Stop |
                Select-Object -ExpandProperty DomainName
            """;
        var domainsResult = await _runner.RunAsync(domainsScript, parameters: null, progress, cancellationToken).ConfigureAwait(false);
        if (!domainsResult.Success)
        {
            throw new InvalidOperationException("Get-AcceptedDomain falló: " + string.Join("; ", domainsResult.Errors));
        }
        var acceptedDomains = domainsResult.Output
            .Select(o => o?.ToString())
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .Select(s => s!)
            .ToList();

        progress?.Report(LogEntry.Info("ExoAudit",
            $"Aceptados {acceptedDomains.Count} dominios. Get-TransportRule completo..."));

        const string rulesScript = """
            param()
            Get-TransportRule -ErrorAction Stop |
                Select-Object Name, State, Priority, Mode, Description,
                              @{N='ForwardTo';E={ if ($_.ForwardTo) { @($_.ForwardTo | ForEach-Object { $_.ToString() }) } else { @() } }},
                              @{N='BlindCopyTo';E={ if ($_.BlindCopyTo) { @($_.BlindCopyTo | ForEach-Object { $_.ToString() }) } else { @() } }},
                              @{N='RedirectMessageTo';E={ if ($_.RedirectMessageTo) { @($_.RedirectMessageTo | ForEach-Object { $_.ToString() }) } else { @() } }},
                              RouteMessageOutboundConnector, DeleteMessage, SentToScope, FromScope
            """;
        var rulesResult = await _runner.RunAsync(rulesScript, parameters: null, progress, cancellationToken).ConfigureAwait(false);
        if (!rulesResult.Success)
        {
            throw new InvalidOperationException("Get-TransportRule falló: " + string.Join("; ", rulesResult.Errors));
        }

        var snapshots = rulesResult.Output
            .OfType<PSObject>()
            .Select(o => new Grex365.Core.Audit.TransportRuleSnapshot(
                Name: o.Properties["Name"]?.Value?.ToString() ?? string.Empty,
                State: o.Properties["State"]?.Value?.ToString() ?? string.Empty,
                Priority: o.Properties["Priority"]?.Value is int pi ? pi : 0,
                Mode: o.Properties["Mode"]?.Value?.ToString() ?? string.Empty,
                Description: o.Properties["Description"]?.Value?.ToString(),
                ForwardTo: ToList(o.Properties["ForwardTo"]?.Value),
                BlindCopyTo: ToList(o.Properties["BlindCopyTo"]?.Value),
                RedirectMessageTo: ToList(o.Properties["RedirectMessageTo"]?.Value),
                RouteMessageOutboundConnector: o.Properties["RouteMessageOutboundConnector"]?.Value?.ToString(),
                DeleteMessage: o.Properties["DeleteMessage"]?.Value is bool dm && dm,
                SentToScope: o.Properties["SentToScope"]?.Value?.ToString(),
                FromScope: o.Properties["FromScope"]?.Value?.ToString()))
            .ToList();

        var (summary, findings) = Grex365.Core.Audit.TransportRuleAuditAnalyzer.Analyze(snapshots, acceptedDomains);
        progress?.Report(LogEntry.Ok("ExoAudit",
            $"Transport rules: {summary.Total} totales · {summary.Enabled} enabled · " +
            $"fwd-ext={summary.WithExternalForward} bcc-ext={summary.WithExternalBcc} redir-ext={summary.WithExternalRedirect} · " +
            $"{findings.Count} hallazgos"));
        return (summary, findings);
    }

    private static IReadOnlyList<string> ToList(object? value)
    {
        if (value is null) return Array.Empty<string>();
        if (value is System.Collections.IEnumerable enumerable and not string)
        {
            var list = new List<string>();
            foreach (var item in enumerable)
            {
                if (item is null) continue;
                var s = item.ToString();
                if (!string.IsNullOrWhiteSpace(s)) list.Add(s!);
            }
            return list;
        }
        var single = value.ToString();
        return string.IsNullOrWhiteSpace(single) ? Array.Empty<string>() : new[] { single! };
    }
}
