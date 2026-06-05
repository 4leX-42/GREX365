using System.Text.Json;
using Grex365.Core.Abstractions;
using Grex365.Core.Audit;
using Grex365.Core.Models;

namespace Grex365.PowerShell;

// Read-only Exchange Online security audits (external forwarding, inbox rules, transport rules,
// shared-mailbox sign-in). Each scan runs as ONE external pwsh invocation (connect once, gather,
// emit JSON) instead of the in-proc RunspacePool, which trips the EXO V3 "HttpResponseMessage
// does not contain GetResponseHeader" bug when reading mailbox/rule PSObjects. The pure analyzers
// (already unit-tested) consume the parsed rows.
//
// NOTE: bodies emit JSON via the pipeline form ($list | ConvertTo-Json -AsArray) and concatenate
// the fragments as strings. Wrapping the List[object] in @(...) — or assembling a combined
// [PSCustomObject]@{ ... } — throws "Argument types do not match" when the items carry
// array-typed properties (EXO module ETS quirk, reproduced live 2026-06-05).
public sealed class ExoForwardingAuditService : IExoForwardingAuditService
{
    private readonly IExternalExoRunner _runner;
    private readonly IExchangeConnection? _exchange;

    public ExoForwardingAuditService(IExternalExoRunner runner, IExchangeConnection? exchange = null)
    {
        _runner = runner;
        _exchange = exchange;
    }

    // Fail fast with an actionable message if there's no Exchange Online session yet, instead of
    // letting the external connect-by-cert surface a cryptic error later.
    private void EnsureExchangeConnected()
    {
        if (_exchange is not null && !_exchange.IsConnected)
        {
            throw new InvalidOperationException(
                "No hay conexión con Exchange Online. Conéctate (certificado) en la pestaña " +
                "Conexión antes de ejecutar auditorías de correo.");
        }
    }

    public async Task<IReadOnlyList<AuditFinding>> ScanExternalForwardingAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        EnsureExchangeConnected();
        progress?.Report(LogEntry.Info("ExoAudit", "Get-AcceptedDomain + buzones con forwarding..."));

        var body = $$"""
            $domains = @(Get-AcceptedDomain -ErrorAction Stop | Select-Object -ExpandProperty DomainName)
            $mbx = @(Get-Mailbox -ResultSize Unlimited -ErrorAction Stop |
                Where-Object { $_.ForwardingSmtpAddress -or $_.ForwardingAddress })
            $rows = New-Object System.Collections.Generic.List[object]
            foreach ($m in $mbx) {
                $rows.Add([PSCustomObject]@{
                    UserPrincipalName     = [string]$m.UserPrincipalName
                    ForwardingSmtpAddress = [string]$m.ForwardingSmtpAddress
                    ForwardingAddress     = [string]$m.ForwardingAddress
                })
            }
            $jd = if ($domains.Count -gt 0) { $domains | ConvertTo-Json -Compress -AsArray } else { '[]' }
            $jr = if ($rows.Count -gt 0) { $rows | ConvertTo-Json -Compress -Depth 6 -AsArray } else { '[]' }
            Write-Output ('{{ExternalExoRunner.JsonMarker}}' + ('{"Domains":' + $jd + ',"Rows":' + $jr + '}'))
            """;

        var json = await _runner.RunAsync(body, progress, cancellationToken).ConfigureAwait(false);
        using var doc = ParseDoc(json);
        var root = doc?.RootElement;
        var domains = ReadStringList(root, "Domains");
        var rows = EnumerateObjects(root, "Rows")
            .Select(o => new MailboxForwardingRow(
                UserPrincipalName: ReadString(o, "UserPrincipalName") ?? string.Empty,
                ForwardingSmtpAddress: ReadString(o, "ForwardingSmtpAddress"),
                ForwardingAddress: ReadString(o, "ForwardingAddress")))
            .ToList();

        var findings = MailboxForwardingAnalyzer.Analyze(rows, domains);
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

        EnsureExchangeConnected();
        progress?.Report(LogEntry.Info("ExoAudit",
            $"Get-AcceptedDomain + Get-Mailbox (tope {maxMailboxes}) + Get-InboxRule (puede tardar)..."));

        // One invocation: connect once, enumerate mailboxes, iterate Get-InboxRule per mailbox.
        // Non-JSON Write-Output lines are forwarded as progress by the runner.
        var body = $$"""
            $top = {{maxMailboxes}}
            $domains = @(Get-AcceptedDomain -ErrorAction Stop | Select-Object -ExpandProperty DomainName)
            $upns = @(Get-Mailbox -ResultSize $top -RecipientTypeDetails UserMailbox -ErrorAction Stop |
                Select-Object -ExpandProperty UserPrincipalName)
            Write-Output ("$($upns.Count) buzones; iterando Get-InboxRule...")
            $rules = New-Object System.Collections.Generic.List[object]
            $done = 0
            foreach ($upn in $upns) {
                try {
                    $r = @(Get-InboxRule -Mailbox $upn -ErrorAction Stop)
                    foreach ($rule in $r) {
                        $rules.Add([PSCustomObject]@{
                            MailboxUpn            = [string]$upn
                            Name                  = [string]$rule.Name
                            Enabled               = [bool]$rule.Enabled
                            DeleteMessage         = [bool]$rule.DeleteMessage
                            MoveToFolder          = [string]$rule.MoveToFolder
                            ForwardTo             = @( if ($rule.ForwardTo) { $rule.ForwardTo | ForEach-Object { $_.ToString() } } )
                            ForwardAsAttachmentTo = @( if ($rule.ForwardAsAttachmentTo) { $rule.ForwardAsAttachmentTo | ForEach-Object { $_.ToString() } } )
                            RedirectTo            = @( if ($rule.RedirectTo) { $rule.RedirectTo | ForEach-Object { $_.ToString() } } )
                            SubjectContainsWords  = @( if ($rule.SubjectContainsWords) { $rule.SubjectContainsWords } )
                            BodyContainsWords     = @( if ($rule.BodyContainsWords) { $rule.BodyContainsWords } )
                        })
                    }
                } catch {
                    Write-Output ("Get-InboxRule $upn: " + $_.Exception.Message)
                }
                $done++
                if ($done % 25 -eq 0) { Write-Output ("Progreso: $done/$($upns.Count)") }
            }
            $jd = if ($domains.Count -gt 0) { $domains | ConvertTo-Json -Compress -AsArray } else { '[]' }
            $jr = if ($rules.Count -gt 0) { $rules | ConvertTo-Json -Compress -Depth 8 -AsArray } else { '[]' }
            Write-Output ('{{ExternalExoRunner.JsonMarker}}' + ('{"Domains":' + $jd + ',"Rules":' + $jr + '}'))
            """;

        var json = await _runner.RunAsync(body, progress, cancellationToken).ConfigureAwait(false);
        using var doc = ParseDoc(json);
        var root = doc?.RootElement;
        var domains = ReadStringList(root, "Domains");
        var allRules = EnumerateObjects(root, "Rules")
            .Select(o => new InboxRuleRow(
                MailboxUpn: ReadString(o, "MailboxUpn") ?? string.Empty,
                RuleName: ReadString(o, "Name") ?? "(sin nombre)",
                Enabled: ReadBool(o, "Enabled"),
                DeleteMessage: ReadBool(o, "DeleteMessage"),
                MoveToFolder: ReadString(o, "MoveToFolder"),
                ForwardTo: ReadStringList(o, "ForwardTo"),
                ForwardAsAttachmentTo: ReadStringList(o, "ForwardAsAttachmentTo"),
                RedirectTo: ReadStringList(o, "RedirectTo"),
                SubjectContainsWords: ReadStringList(o, "SubjectContainsWords"),
                BodyContainsWords: ReadStringList(o, "BodyContainsWords")))
            .ToList();

        var findings = InboxRuleAnalyzer.Analyze(allRules, domains);
        progress?.Report(LogEntry.Ok("ExoAudit",
            $"{findings.Count} reglas sospechosas sobre {allRules.Count} reglas."));
        return findings;
    }

    public async Task<(TransportRulesSummary Summary, IReadOnlyList<AuditFinding> Findings)>
        ScanTransportRulesAsync(
            IProgress<LogEntry>? progress = null,
            CancellationToken cancellationToken = default)
    {
        EnsureExchangeConnected();
        progress?.Report(LogEntry.Info("ExoAudit", "Get-AcceptedDomain + Get-TransportRule completo..."));

        var body = $$"""
            $domains = @(Get-AcceptedDomain -ErrorAction Stop | Select-Object -ExpandProperty DomainName)
            $rules = @(Get-TransportRule -ErrorAction Stop)
            $out = New-Object System.Collections.Generic.List[object]
            foreach ($r in $rules) {
                $out.Add([PSCustomObject]@{
                    Name                          = [string]$r.Name
                    State                         = [string]$r.State
                    Priority                      = [int]$r.Priority
                    Mode                          = [string]$r.Mode
                    Description                   = [string]$r.Description
                    ForwardTo                     = @( if ($r.ForwardTo) { $r.ForwardTo | ForEach-Object { $_.ToString() } } )
                    BlindCopyTo                   = @( if ($r.BlindCopyTo) { $r.BlindCopyTo | ForEach-Object { $_.ToString() } } )
                    RedirectMessageTo             = @( if ($r.RedirectMessageTo) { $r.RedirectMessageTo | ForEach-Object { $_.ToString() } } )
                    RouteMessageOutboundConnector = [string]$r.RouteMessageOutboundConnector
                    DeleteMessage                 = [bool]$r.DeleteMessage
                    SentToScope                   = [string]$r.SentToScope
                    FromScope                     = [string]$r.FromScope
                })
            }
            $jd = if ($domains.Count -gt 0) { $domains | ConvertTo-Json -Compress -AsArray } else { '[]' }
            $jr = if ($out.Count -gt 0) { $out | ConvertTo-Json -Compress -Depth 8 -AsArray } else { '[]' }
            Write-Output ('{{ExternalExoRunner.JsonMarker}}' + ('{"Domains":' + $jd + ',"Rules":' + $jr + '}'))
            """;

        var json = await _runner.RunAsync(body, progress, cancellationToken).ConfigureAwait(false);
        using var doc = ParseDoc(json);
        var root = doc?.RootElement;
        var domains = ReadStringList(root, "Domains");
        var snapshots = EnumerateObjects(root, "Rules")
            .Select(o => new TransportRuleSnapshot(
                Name: ReadString(o, "Name") ?? string.Empty,
                State: ReadString(o, "State") ?? string.Empty,
                Priority: ReadInt(o, "Priority"),
                Mode: ReadString(o, "Mode") ?? string.Empty,
                Description: ReadString(o, "Description"),
                ForwardTo: ReadStringList(o, "ForwardTo"),
                BlindCopyTo: ReadStringList(o, "BlindCopyTo"),
                RedirectMessageTo: ReadStringList(o, "RedirectMessageTo"),
                RouteMessageOutboundConnector: ReadString(o, "RouteMessageOutboundConnector"),
                DeleteMessage: ReadBool(o, "DeleteMessage"),
                SentToScope: ReadString(o, "SentToScope"),
                FromScope: ReadString(o, "FromScope")))
            .ToList();

        var (summary, findings) = TransportRuleAuditAnalyzer.Analyze(snapshots, domains);
        progress?.Report(LogEntry.Ok("ExoAudit",
            $"Transport rules: {summary.Total} totales · {summary.Enabled} enabled · " +
            $"fwd-ext={summary.WithExternalForward} bcc-ext={summary.WithExternalBcc} redir-ext={summary.WithExternalRedirect} · " +
            $"{findings.Count} hallazgos"));
        return (summary, findings);
    }

    public async Task<(SharedMailboxSignInSummary Summary, IReadOnlyList<AuditFinding> Findings)>
        ScanSharedMailboxSignInAsync(
            IProgress<LogEntry>? progress = null,
            CancellationToken cancellationToken = default)
    {
        EnsureExchangeConnected();
        progress?.Report(LogEntry.Info("ExoAudit", "Enumerando shared mailboxes y estado AccountDisabled..."));

        var body = $$"""
            $shared = @(Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -ErrorAction Stop)
            $rows = New-Object System.Collections.Generic.List[object]
            foreach ($m in $shared) {
                $upn = [string]$m.UserPrincipalName
                $disabled = $null
                try {
                    $u = Get-User -Identity $upn -ErrorAction Stop
                    $disabled = [bool]$u.AccountDisabled
                } catch {
                    $disabled = $null
                }
                $rows.Add([PSCustomObject]@{
                    UserPrincipalName = $upn
                    DisplayName       = [string]$m.DisplayName
                    AccountDisabled   = $disabled
                })
            }
            $jr = if ($rows.Count -gt 0) { $rows | ConvertTo-Json -Compress -Depth 5 -AsArray } else { '[]' }
            Write-Output ('{{ExternalExoRunner.JsonMarker}}' + $jr)
            """;

        var json = await _runner.RunAsync(body, progress, cancellationToken).ConfigureAwait(false);
        using var doc = ParseDoc(json);
        var rows = EnumerateArray(doc?.RootElement)
            .Select(o => new SharedMailboxSignInRow(
                UserPrincipalName: ReadString(o, "UserPrincipalName") ?? string.Empty,
                DisplayName: ReadString(o, "DisplayName"),
                AccountDisabled: ReadNullableBool(o, "AccountDisabled")))
            .ToList();

        var (summary, findings) = SharedMailboxSignInAnalyzer.Analyze(rows);
        progress?.Report(LogEntry.Ok("ExoAudit",
            $"Shared mailboxes: {summary.Total} totales · " +
            $"sign-in enabled={summary.SignInEnabled} (WARN) · disabled={summary.SignInDisabled} · " +
            $"unknown={summary.Unknown}"));
        return (summary, findings);
    }

    // ---- JSON helpers: tolerant of ConvertTo-Json quirks (single-element arrays collapse to a
    //      scalar/object; absent properties; empty strings). ----

    private static JsonDocument? ParseDoc(string? json) =>
        string.IsNullOrWhiteSpace(json) ? null : JsonDocument.Parse(json);

    private static string? ReadString(JsonElement el, string name) =>
        el.ValueKind == JsonValueKind.Object && el.TryGetProperty(name, out var v)
            && v.ValueKind == JsonValueKind.String && v.GetString() is { Length: > 0 } s
            ? s : null;

    private static int ReadInt(JsonElement el, string name) =>
        el.ValueKind == JsonValueKind.Object && el.TryGetProperty(name, out var v)
            && v.ValueKind == JsonValueKind.Number && v.TryGetInt32(out var i) ? i : 0;

    private static bool ReadBool(JsonElement el, string name) =>
        el.ValueKind == JsonValueKind.Object && el.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.True;

    private static bool? ReadNullableBool(JsonElement el, string name) =>
        el.ValueKind == JsonValueKind.Object && el.TryGetProperty(name, out var v)
            ? v.ValueKind switch { JsonValueKind.True => true, JsonValueKind.False => false, _ => (bool?)null }
            : null;

    private static IReadOnlyList<string> ReadStringList(JsonElement? parent, string name)
    {
        if (parent is not { ValueKind: JsonValueKind.Object } p || !p.TryGetProperty(name, out var v))
        {
            return Array.Empty<string>();
        }
        var list = new List<string>();
        if (v.ValueKind == JsonValueKind.Array)
        {
            foreach (var e in v.EnumerateArray())
            {
                if (e.ValueKind == JsonValueKind.String && e.GetString() is { Length: > 0 } s) list.Add(s);
            }
        }
        else if (v.ValueKind == JsonValueKind.String && v.GetString() is { Length: > 0 } single)
        {
            list.Add(single); // single-element array collapsed to a scalar by ConvertTo-Json
        }
        return list;
    }

    // Property that should be an array of objects; ConvertTo-Json may collapse a single element
    // to a lone object. Yields each object element either way.
    private static IEnumerable<JsonElement> EnumerateObjects(JsonElement? parent, string name)
    {
        if (parent is not { ValueKind: JsonValueKind.Object } p || !p.TryGetProperty(name, out var v))
        {
            yield break;
        }
        foreach (var e in EnumerateArray(v))
        {
            yield return e;
        }
    }

    private static IEnumerable<JsonElement> EnumerateArray(JsonElement? value)
    {
        if (value is not { } v) yield break;
        if (v.ValueKind == JsonValueKind.Array)
        {
            foreach (var e in v.EnumerateArray())
            {
                if (e.ValueKind == JsonValueKind.Object) yield return e;
            }
        }
        else if (v.ValueKind == JsonValueKind.Object)
        {
            yield return v;
        }
    }
}
