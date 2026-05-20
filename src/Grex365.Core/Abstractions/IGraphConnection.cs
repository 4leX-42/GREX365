using Grex365.Core.Models;
using Microsoft.Graph;

namespace Grex365.Core.Abstractions;

public sealed record DeviceCodePrompt(string UserCode, string VerificationUri, string Message, DateTimeOffset ExpiresOn);

public interface IGraphConnection
{
    bool IsConnected { get; }

    GraphServiceClient? Client { get; }

    string? TenantId { get; }

    string? Account { get; }

    Task ConnectByCertificateAsync(
        CertConfig config,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    Task ConnectByDeviceCodeAsync(
        string? tenantHint,
        Func<DeviceCodePrompt, CancellationToken, Task> codeCallback,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    Task<bool> CheckLiveAsync(CancellationToken cancellationToken = default);

    Task DisconnectAsync(CancellationToken cancellationToken = default);
}
