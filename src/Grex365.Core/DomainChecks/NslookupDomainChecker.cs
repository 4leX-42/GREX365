using System.Diagnostics;
using System.Text.RegularExpressions;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.Core.DomainChecks;

public sealed class NslookupDomainChecker : IDomainChecker
{
    public async Task<DomainCheck> CheckAsync(string domain, IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default)
    {
        var d = (domain ?? string.Empty).Trim();
        if (string.IsNullOrEmpty(d))
        {
            throw new ArgumentException("Dominio vacío.", nameof(domain));
        }

        var records = new List<DnsRecord>();

        progress?.Report(LogEntry.Info("DNS", $"Consultando MX para {d}..."));
        records.Add(await LookupAsync("MX", d, cancellationToken).ConfigureAwait(false));

        progress?.Report(LogEntry.Info("DNS", $"Consultando TXT para {d} (SPF/DMARC)..."));
        var txt = await LookupAsync("TXT", d, cancellationToken).ConfigureAwait(false);
        records.Add(txt);

        // SPF check
        var spf = txt.Value.Split('\n').FirstOrDefault(l => l.Contains("v=spf1", StringComparison.OrdinalIgnoreCase));
        records.Add(spf is not null
            ? new DnsRecord("SPF", "OK", spf.Trim())
            : new DnsRecord("SPF", "MISSING", "No se encontró v=spf1"));

        progress?.Report(LogEntry.Info("DNS", $"Consultando _dmarc.{d}..."));
        var dmarc = await LookupAsync("TXT", "_dmarc." + d, cancellationToken).ConfigureAwait(false);
        records.Add(dmarc.Status == "OK"
            ? new DnsRecord("DMARC", "OK", dmarc.Value)
            : new DnsRecord("DMARC", "MISSING", "Sin registro _dmarc"));

        return new DomainCheck(d, records);
    }

    private static async Task<DnsRecord> LookupAsync(string type, string target, CancellationToken ct)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "nslookup",
                Arguments = $"-type={type} {target}",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            using var proc = Process.Start(psi);
            if (proc is null)
            {
                return new DnsRecord(type, "ERROR", "No se pudo iniciar nslookup");
            }

            using var reg = ct.Register(() =>
            {
                try { proc.Kill(true); } catch { /* ignore */ }
            });

            var stdout = await proc.StandardOutput.ReadToEndAsync(ct).ConfigureAwait(false);
            await proc.WaitForExitAsync(ct).ConfigureAwait(false);

            var clean = CleanOutput(stdout);
            if (string.IsNullOrWhiteSpace(clean) || clean.Contains("can't find", StringComparison.OrdinalIgnoreCase))
            {
                return new DnsRecord(type, "MISSING", clean);
            }
            return new DnsRecord(type, "OK", clean);
        }
        catch (Exception ex)
        {
            return new DnsRecord(type, "ERROR", ex.Message);
        }
    }

    private static string CleanOutput(string raw)
    {
        var lines = raw.Split('\n')
            .Select(l => l.TrimEnd('\r'))
            .Where(l => !string.IsNullOrWhiteSpace(l))
            .Where(l => !l.StartsWith("Servidor", StringComparison.OrdinalIgnoreCase))
            .Where(l => !l.StartsWith("Server:", StringComparison.OrdinalIgnoreCase))
            .Where(l => !l.StartsWith("Address:", StringComparison.OrdinalIgnoreCase))
            .Where(l => !l.StartsWith("Direccion:", StringComparison.OrdinalIgnoreCase))
            .Where(l => !Regex.IsMatch(l, "^[A-Z][a-z]+:?\\s*$"))
            .ToList();
        return string.Join("\n", lines);
    }
}
