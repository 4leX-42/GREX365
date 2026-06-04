using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

// Runs a single Exchange Online script body in an external pwsh.exe process: it loads the app
// certificate config, connects (Connect-ExchangeOnline by cert), runs the body wrapped in
// try/catch/finally(disconnect), and returns the JSON line the body emitted after the JSON marker
// (or null). Centralises the external-host plumbing so every EXO consumer shares one validated
// path. The in-process RunspacePool is unreliable for EXO V3 (the "HttpResponseMessage does not
// contain GetResponseHeader" failure), so the modern REST cmdlets run in a full PowerShell host.
//
// The body must emit its result as: Write-Output (JsonMarker + (... | ConvertTo-Json -Compress)).
// On failure the runner throws InvalidOperationException with the real EXO error message.
public interface IExternalExoRunner
{
    Task<string?> RunAsync(
        string body,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);
}
