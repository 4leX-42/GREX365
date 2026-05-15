using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

public interface IPowerShellRunner : IAsyncDisposable
{
    Task<PowerShellResult> RunAsync(
        string script,
        IDictionary<string, object?>? parameters = null,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);
}

public sealed record PowerShellResult(
    bool Success,
    IReadOnlyList<object?> Output,
    IReadOnlyList<string> Errors);
