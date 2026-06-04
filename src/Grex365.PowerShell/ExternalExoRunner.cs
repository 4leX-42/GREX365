using System.Diagnostics;
using System.Text;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.PowerShell;

// Single source of truth for the external pwsh.exe Exchange Online host. Bodies emit their result
// after JsonMarker; failures surface the real EXO message via ErrMarker. See IExternalExoRunner.
public sealed class ExternalExoRunner : IExternalExoRunner
{
    public const string JsonMarker = "###GREX-JSON###";
    public const string ErrMarker = "###GREX-ERR###";

    private readonly ICertConfigStore _certStore;

    public ExternalExoRunner(ICertConfigStore certStore)
    {
        _certStore = certStore;
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
    public async Task<string?> RunAsync(
        string body,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var cfg = await RequireConfigAsync(cancellationToken).ConfigureAwait(false);
        var full = $$"""
            $ErrorActionPreference = 'Stop'
            $ProgressPreference = 'SilentlyContinue'
            $InformationPreference = 'SilentlyContinue'
            $WarningPreference = 'SilentlyContinue'
            [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
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
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
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

    // PowerShell single-quoted literal (doubles embedded quotes) — safe for injection.
    internal static string Lit(string? value) => "'" + (value ?? string.Empty).Replace("'", "''") + "'";
}
