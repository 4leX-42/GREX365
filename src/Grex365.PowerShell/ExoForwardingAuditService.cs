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
}
