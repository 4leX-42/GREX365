using System.Globalization;
using Grex365.Core.Abstractions;
using Grex365.Core.Mailboxes;
using Grex365.Core.Models;

namespace Grex365.PowerShell;

public sealed class MailboxRulesService : IMailboxRulesService
{
    private readonly IPowerShellRunner _runner;

    public MailboxRulesService(IPowerShellRunner runner)
    {
        _runner = runner;
    }

    public async Task<AutoReplyConfig?> GetAutoReplyAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        const string script = """
            $cfg = Get-MailboxAutoReplyConfiguration -Identity $Identity -ErrorAction Stop
            [PSCustomObject]@{
                State            = [string]$cfg.AutoReplyState
                InternalMessage  = [string]$cfg.InternalMessage
                ExternalMessage  = [string]$cfg.ExternalMessage
                StartTime        = [string]$cfg.StartTime
                EndTime          = [string]$cfg.EndTime
            }
            """;

        var result = await _runner.RunAsync(
            script,
            new Dictionary<string, object?> { ["Identity"] = identity },
            progress,
            cancellationToken).ConfigureAwait(false);

        if (!result.Success)
        {
            throw new InvalidOperationException("Get-MailboxAutoReplyConfiguration falló: " + string.Join("; ", result.Errors));
        }
        if (result.Output.Count == 0)
        {
            return null;
        }
        return MapAutoReply(result.Output[0]);
    }

    public async Task SetAutoReplyAsync(
        string identity,
        AutoReplyConfig config,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var errors = MailboxRulesValidator.ValidateAutoReply(config);
        if (errors.Count > 0)
        {
            throw new ArgumentException("AutoReply inválido: " + string.Join("; ", errors));
        }

        const string script = """
            $params = @{
                Identity        = $Identity
                AutoReplyState  = $State
                ErrorAction     = 'Stop'
            }
            if ($InternalMessage) { $params['InternalMessage'] = $InternalMessage }
            if ($ExternalMessage) { $params['ExternalMessage'] = $ExternalMessage }
            if ($StartTime)       { $params['StartTime']       = [datetime]::Parse($StartTime, [System.Globalization.CultureInfo]::InvariantCulture) }
            if ($EndTime)         { $params['EndTime']         = [datetime]::Parse($EndTime,   [System.Globalization.CultureInfo]::InvariantCulture) }
            Set-MailboxAutoReplyConfiguration @params | Out-Null
            Write-Information "AutoReply configurado: $State"
            """;

        var parameters = new Dictionary<string, object?>
        {
            ["Identity"] = identity,
            ["State"] = config.State.ToString(),
            ["InternalMessage"] = config.InternalMessage,
            ["ExternalMessage"] = config.ExternalMessage,
            ["StartTime"] = config.StartTime?.ToString("o", CultureInfo.InvariantCulture),
            ["EndTime"] = config.EndTime?.ToString("o", CultureInfo.InvariantCulture)
        };

        var result = await _runner.RunAsync(script, parameters, progress, cancellationToken).ConfigureAwait(false);
        if (!result.Success)
        {
            throw new InvalidOperationException("Set-MailboxAutoReplyConfiguration falló: " + string.Join("; ", result.Errors));
        }
        progress?.Report(LogEntry.Ok("Mailbox", $"AutoReply {config.State} en {identity}"));
    }

    public async Task<ForwardingConfig?> GetForwardingAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        const string script = """
            $m = Get-Mailbox -Identity $Identity -ErrorAction Stop
            [PSCustomObject]@{
                ForwardingAddress              = [string]$m.ForwardingAddress
                ForwardingSmtpAddress          = [string]$m.ForwardingSmtpAddress
                DeliverToMailboxAndForward     = [bool]$m.DeliverToMailboxAndForward
            }
            """;

        var result = await _runner.RunAsync(
            script,
            new Dictionary<string, object?> { ["Identity"] = identity },
            progress,
            cancellationToken).ConfigureAwait(false);

        if (!result.Success)
        {
            throw new InvalidOperationException("Get-Mailbox (forwarding) falló: " + string.Join("; ", result.Errors));
        }
        if (result.Output.Count == 0)
        {
            return null;
        }
        return MapForwarding(result.Output[0]);
    }

    public async Task SetForwardingAsync(
        string identity,
        string forwardingSmtpAddress,
        bool deliverToMailboxAndForward,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var errors = MailboxRulesValidator.ValidateForwarding(forwardingSmtpAddress);
        if (errors.Count > 0)
        {
            throw new ArgumentException("Forwarding inválido: " + string.Join("; ", errors));
        }

        const string script = """
            Set-Mailbox -Identity $Identity `
                -ForwardingSmtpAddress $Smtp `
                -DeliverToMailboxAndForward:$Deliver `
                -ErrorAction Stop | Out-Null
            Write-Information "Forwarding configurado: $Smtp (deliver=$Deliver)"
            """;

        var result = await _runner.RunAsync(
            script,
            new Dictionary<string, object?>
            {
                ["Identity"] = identity,
                ["Smtp"] = forwardingSmtpAddress,
                ["Deliver"] = deliverToMailboxAndForward
            },
            progress,
            cancellationToken).ConfigureAwait(false);
        if (!result.Success)
        {
            throw new InvalidOperationException("Set-Mailbox (forwarding) falló: " + string.Join("; ", result.Errors));
        }
        progress?.Report(LogEntry.Ok("Mailbox", $"Forwarding {identity} → {forwardingSmtpAddress}"));
    }

    public async Task ClearForwardingAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        const string script = """
            Set-Mailbox -Identity $Identity `
                -ForwardingAddress $null `
                -ForwardingSmtpAddress $null `
                -DeliverToMailboxAndForward:$false `
                -ErrorAction Stop | Out-Null
            Write-Information "Forwarding limpiado en $Identity"
            """;

        var result = await _runner.RunAsync(
            script,
            new Dictionary<string, object?> { ["Identity"] = identity },
            progress,
            cancellationToken).ConfigureAwait(false);
        if (!result.Success)
        {
            throw new InvalidOperationException("Clear forwarding falló: " + string.Join("; ", result.Errors));
        }
        progress?.Report(LogEntry.Ok("Mailbox", $"Forwarding limpiado en {identity}"));
    }

    private static AutoReplyConfig MapAutoReply(object? raw)
    {
        string? Get(string name)
        {
            if (raw is System.Management.Automation.PSObject ps)
            {
                return ps.Properties[name]?.Value?.ToString();
            }
            var t = raw?.GetType();
            return t?.GetProperty(name)?.GetValue(raw)?.ToString();
        }

        var stateStr = Get("State") ?? "Disabled";
        var state = stateStr switch
        {
            "Enabled" => AutoReplyState.Enabled,
            "Scheduled" => AutoReplyState.Scheduled,
            _ => AutoReplyState.Disabled
        };
        DateTime? ParseDate(string name)
        {
            var s = Get(name);
            if (string.IsNullOrWhiteSpace(s)) return null;
            return DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var dt) ? dt : null;
        }

        return new AutoReplyConfig(
            State: state,
            InternalMessage: Get("InternalMessage"),
            ExternalMessage: Get("ExternalMessage"),
            StartTime: ParseDate("StartTime"),
            EndTime: ParseDate("EndTime"));
    }

    private static ForwardingConfig MapForwarding(object? raw)
    {
        string? GetString(string name)
        {
            if (raw is System.Management.Automation.PSObject ps)
            {
                return ps.Properties[name]?.Value?.ToString();
            }
            var t = raw?.GetType();
            return t?.GetProperty(name)?.GetValue(raw)?.ToString();
        }

        var deliverStr = GetString("DeliverToMailboxAndForward") ?? "False";
        bool.TryParse(deliverStr, out var deliver);
        return new ForwardingConfig(
            ForwardingAddress: GetString("ForwardingAddress"),
            ForwardingSmtpAddress: GetString("ForwardingSmtpAddress"),
            DeliverToMailboxAndForward: deliver);
    }
}
