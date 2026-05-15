using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

public interface IExchangeConnection
{
    bool IsConnected { get; }

    string? TenantId { get; }

    string? Organization { get; }

    Task ConnectByCertificateAsync(
        CertConfig config,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    Task<bool> CheckLiveAsync(CancellationToken cancellationToken = default);

    Task DisconnectAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);
}
