using System.Diagnostics;
using System.Text;
using System.Text.Json;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.PowerShell;

public sealed class ExternalExoOps : IExternalExoOps
{
    private const string JsonMarker = "###GREX-JSON###";
    private const string ErrMarker = "###GREX-ERR###";
    private readonly ICertConfigStore _certStore;

    public ExternalExoOps(ICertConfigStore certStore)
    {
        _certStore = certStore;
    }

    public async Task<MailboxInfo?> GetMailboxFactsAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var cfg = await RequireConfigAsync(cancellationToken).ConfigureAwait(false);
        var id = Lit(identity);
        var body = $$"""
            $id = {{id}}
            $m = Get-Mailbox -Identity $id -ErrorAction Stop
            $bytes = $null
            try {
                $s = Get-EXOMailboxStatistics -Identity $id -Properties TotalItemSize -ErrorAction Stop
                $t = [string]$s.TotalItemSize
                if ($t -match '\(([\d,]+) bytes\)') { $bytes = [int64]($matches[1] -replace ',','') }
            } catch { }
            $holds = 0; if ($m.InPlaceHolds) { $holds = @($m.InPlaceHolds).Count }
            $arch = $false
            if ($m.ArchiveStatus -and [string]$m.ArchiveStatus -ne 'None') { $arch = $true }
            $o = [PSCustomObject]@{
                Identity             = [string]$m.Identity
                DisplayName          = [string]$m.DisplayName
                PrimarySmtpAddress   = [string]$m.PrimarySmtpAddress
                RecipientTypeDetails = [string]$m.RecipientTypeDetails
                LitigationHoldEnabled= [bool]$m.LitigationHoldEnabled
                InPlaceHoldCount     = [int]$holds
                ArchiveEnabled       = [bool]$arch
                TotalItemBytes       = $bytes
            }
            Write-Output ('{{JsonMarker}}' + ($o | ConvertTo-Json -Compress))
            """;

        var json = await RunAsync(cfg, body, progress, cancellationToken).ConfigureAwait(false);
        return Parse(json);
    }

    public async Task<MailboxInfo?> ConvertToSharedAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var cfg = await RequireConfigAsync(cancellationToken).ConfigureAwait(false);
        var id = Lit(identity);
        var body = $$"""
            $id = {{id}}
            $cur = Get-Mailbox -Identity $id -ErrorAction Stop
            if ($cur.RecipientTypeDetails -ne 'SharedMailbox') {
                Set-Mailbox -Identity $id -Type Shared -ErrorAction Stop
                Write-Output 'Set-Mailbox -Type Shared aplicado; esperando propagacion...'
            }
            $final = [string]$cur.RecipientTypeDetails
            $deadline = (Get-Date).AddSeconds(150)
            while ($final -ne 'SharedMailbox' -and (Get-Date) -lt $deadline) {
                Start-Sleep -Seconds 8
                try { $final = [string](Get-Mailbox -Identity $id -ErrorAction Stop).RecipientTypeDetails } catch { }
            }
            $m = Get-Mailbox -Identity $id -ErrorAction Stop
            $o = [PSCustomObject]@{
                Identity             = [string]$m.Identity
                DisplayName          = [string]$m.DisplayName
                PrimarySmtpAddress   = [string]$m.PrimarySmtpAddress
                RecipientTypeDetails = [string]$m.RecipientTypeDetails
            }
            Write-Output ('{{JsonMarker}}' + ($o | ConvertTo-Json -Compress))
            """;

        var json = await RunAsync(cfg, body, progress, cancellationToken).ConfigureAwait(false);
        return Parse(json);
    }

    public async Task SetAutoReplyAsync(
        string identity,
        string message,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var cfg = await RequireConfigAsync(cancellationToken).ConfigureAwait(false);
        var body = $$"""
            $id = {{Lit(identity)}}
            $msg = {{Lit(message)}}
            Set-MailboxAutoReplyConfiguration -Identity $id -AutoReplyState Enabled -ExternalAudience All -InternalMessage $msg -ExternalMessage $msg -ErrorAction Stop
            Write-Output ('{{JsonMarker}}' + (([PSCustomObject]@{ Note = 'auto-reply Enabled' }) | ConvertTo-Json -Compress))
            """;
        await RunAsync(cfg, body, progress, cancellationToken).ConfigureAwait(false);
    }

    public async Task SetForwardingAsync(
        string identity,
        string forwardTo,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var cfg = await RequireConfigAsync(cancellationToken).ConfigureAwait(false);
        var body = $$"""
            $id = {{Lit(identity)}}
            Set-Mailbox -Identity $id -ForwardingSmtpAddress {{Lit(forwardTo)}} -DeliverToMailboxAndForward $true -ErrorAction Stop
            Write-Output ('{{JsonMarker}}' + (([PSCustomObject]@{ Note = 'forward configurado' }) | ConvertTo-Json -Compress))
            """;
        await RunAsync(cfg, body, progress, cancellationToken).ConfigureAwait(false);
    }

    public async Task<string> HideFromGalAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var cfg = await RequireConfigAsync(cancellationToken).ConfigureAwait(false);
        // Hybrid objects synced from on-prem AD can't be modified in EXO — detect that specific
        // failure and report it as a SKIP-with-guidance instead of a hard error (legacy parity).
        var body = $$"""
            $id = {{Lit(identity)}}
            try {
                Set-Mailbox -Identity $id -HiddenFromAddressListsEnabled $true -ErrorAction Stop
                $note = 'HiddenFromAddressListsEnabled=true'
            } catch {
                $m = $_.Exception.Message
                if ($m -match 'sincroniz|on-prem|on premises|write scope|organizaci.n local|cannot be performed.*synchron') {
                    $note = 'objeto hibrido sincronizado desde AD on-prem: aplicar msExchHideFromAddressLists=TRUE en AD local'
                } else { throw }
            }
            Write-Output ('{{JsonMarker}}' + (([PSCustomObject]@{ Note = $note }) | ConvertTo-Json -Compress))
            """;
        var json = await RunAsync(cfg, body, progress, cancellationToken).ConfigureAwait(false);
        return ParseNote(json) ?? "aplicado";
    }

    public async Task<string> GrantDelegateAsync(
        string mailbox,
        string delegateUpn,
        bool sendAs,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var cfg = await RequireConfigAsync(cancellationToken).ConfigureAwait(false);
        var sendAsLit = sendAs ? "$true" : "$false";
        var body = $$"""
            $id = {{Lit(mailbox)}}
            $d = {{Lit(delegateUpn)}}
            Add-MailboxPermission -Identity $id -User $d -AccessRights FullAccess -InheritanceType All -AutoMapping:$false -Confirm:$false -ErrorAction Stop | Out-Null
            $note = 'FullAccess'
            if ({{sendAsLit}}) {
                Add-RecipientPermission -Identity $id -Trustee $d -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                $note = 'FullAccess + SendAs'
            }
            Write-Output ('{{JsonMarker}}' + (([PSCustomObject]@{ Note = ($note + ' -> ' + $d) }) | ConvertTo-Json -Compress))
            """;
        var json = await RunAsync(cfg, body, progress, cancellationToken).ConfigureAwait(false);
        return ParseNote(json) ?? "delegado";
    }

    private static string? ParseNote(string? json)
    {
        if (string.IsNullOrWhiteSpace(json)) return null;
        using var doc = JsonDocument.Parse(json);
        return doc.RootElement.TryGetProperty("Note", out var v) && v.ValueKind == JsonValueKind.String
            ? v.GetString()
            : null;
    }

    private async Task<CertConfig> RequireConfigAsync(CancellationToken ct)
    {
        var cfg = await _certStore.LoadAsync(ct).ConfigureAwait(false);
        if (cfg is null || string.IsNullOrWhiteSpace(cfg.CertThumbprint))
        {
            throw new InvalidOperationException(
                "No hay configuración de certificado para Exchange Online. Conéctate por certificado primero.");
        }
        return cfg;
    }

    // Wraps the body in connect/disconnect and runs it in an external pwsh process.
    private async Task<string?> RunAsync(
        CertConfig cfg,
        string body,
        IProgress<LogEntry>? progress,
        CancellationToken cancellationToken)
    {
        var full = $$"""
            $ErrorActionPreference = 'Stop'
            $ProgressPreference = 'SilentlyContinue'
            $InformationPreference = 'SilentlyContinue'
            $WarningPreference = 'SilentlyContinue'
            Import-Module ExchangeOnlineManagement -ErrorAction Stop
            Connect-ExchangeOnline -AppId {{Lit(cfg.AppId)}} -CertificateThumbprint {{Lit(cfg.CertThumbprint)}} -Organization {{Lit(cfg.Organization)}} -ShowBanner:$false -InformationAction SilentlyContinue -ErrorAction Stop | Out-Null
            try {
            {{body}}
            }
            catch {
                Write-Output ('{{ErrMarker}}' + $_.Exception.Message)
            }
            finally {
                Disconnect-ExchangeOnline -Confirm:$false -InformationAction SilentlyContinue -ErrorAction SilentlyContinue *> $null
            }
            """;

        var encoded = Convert.ToBase64String(Encoding.Unicode.GetBytes(full));

        var psi = new ProcessStartInfo
        {
            FileName = ExeResolver.ResolvePwsh(),
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
        };
        psi.ArgumentList.Add("-NoLogo");
        psi.ArgumentList.Add("-NoProfile");
        psi.ArgumentList.Add("-NonInteractive");
        psi.ArgumentList.Add("-EncodedCommand");
        psi.ArgumentList.Add(encoded);

        using var proc = new Process { StartInfo = psi };
        string? jsonLine = null;
        string? scriptError = null;
        var errors = new StringBuilder();

        proc.OutputDataReceived += (_, e) =>
        {
            if (string.IsNullOrEmpty(e.Data)) return;
            // The script wraps the body in try/catch and emits the real exception message on a
            // marked line — capture it so failures surface the actual EXO error instead of a
            // bare "pwsh exit 1".
            var errIdx = e.Data.IndexOf(ErrMarker, StringComparison.Ordinal);
            if (errIdx >= 0)
            {
                scriptError = e.Data[(errIdx + ErrMarker.Length)..];
                return;
            }
            var idx = e.Data.IndexOf(JsonMarker, StringComparison.Ordinal);
            if (idx >= 0)
            {
                jsonLine = e.Data[(idx + JsonMarker.Length)..];
            }
            else
            {
                progress?.Report(LogEntry.Info("EXO", e.Data));
            }
        };
        proc.ErrorDataReceived += (_, e) =>
        {
            if (string.IsNullOrEmpty(e.Data)) return;
            // The EXO module serialises non-text stream records to stderr as CLIXML noise
            // ("#< CLIXML", "<Objs ...>"). It is not an error — drop it so the live log
            // stays clean instead of showing scary XML warnings.
            var t = e.Data.TrimStart();
            if (t.StartsWith("#< CLIXML", StringComparison.Ordinal) ||
                t.StartsWith("<Objs", StringComparison.Ordinal) ||
                t.StartsWith("<", StringComparison.Ordinal))
            {
                return;
            }
            errors.AppendLine(e.Data);
            progress?.Report(LogEntry.Warn("EXO", e.Data));
        };

        proc.Start();
        proc.BeginOutputReadLine();
        proc.BeginErrorReadLine();

        await using (cancellationToken.Register(() =>
        {
            try { if (!proc.HasExited) proc.Kill(entireProcessTree: true); } catch { }
        }).ConfigureAwait(false))
        {
            await proc.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
        }

        if (scriptError is not null)
        {
            throw new InvalidOperationException("Operación Exchange Online falló: " + scriptError.Trim());
        }
        if (proc.ExitCode != 0 && jsonLine is null)
        {
            throw new InvalidOperationException(
                "Operación Exchange Online falló: " + (errors.Length > 0 ? errors.ToString().Trim() : $"pwsh exit {proc.ExitCode}"));
        }
        return jsonLine;
    }

    private static MailboxInfo? Parse(string? json)
    {
        if (string.IsNullOrWhiteSpace(json)) return null;
        using var doc = JsonDocument.Parse(json);
        var r = doc.RootElement;
        string S(string n) => r.TryGetProperty(n, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString()! : string.Empty;
        bool B(string n) => r.TryGetProperty(n, out var v) && (v.ValueKind == JsonValueKind.True || (v.ValueKind == JsonValueKind.String && bool.TryParse(v.GetString(), out var b) && b));
        int I(string n) => r.TryGetProperty(n, out var v) && v.TryGetInt32(out var i) ? i : 0;
        long? L(string n) => r.TryGetProperty(n, out var v) && v.ValueKind == JsonValueKind.Number && v.TryGetInt64(out var l) ? l : null;

        return new MailboxInfo(
            Identity: S("Identity"),
            DisplayName: S("DisplayName"),
            PrimarySmtpAddress: S("PrimarySmtpAddress"),
            RecipientTypeDetails: S("RecipientTypeDetails"),
            LitigationHoldEnabled: B("LitigationHoldEnabled"),
            InPlaceHoldCount: I("InPlaceHoldCount"),
            ArchiveEnabled: B("ArchiveEnabled"),
            TotalItemBytes: L("TotalItemBytes"));
    }

    // PowerShell single-quoted literal (doubles embedded quotes) — safe for injection.
    private static string Lit(string? value) => "'" + (value ?? string.Empty).Replace("'", "''") + "'";
}
