using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

public sealed record ExoModuleStatus(bool Installed, string? Version, string? Detail);

public interface IExchangeConnection
{
    bool IsConnected { get; }

    string? TenantId { get; }

    string? Organization { get; }

    Task ConnectByCertificateAsync(
        CertConfig config,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    Task<ExoModuleStatus> ProbeModuleAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    Task<ExoModuleStatus> InstallModuleAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    Task<bool> CheckLiveAsync(CancellationToken cancellationToken = default);

    Task DisconnectAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);
}
