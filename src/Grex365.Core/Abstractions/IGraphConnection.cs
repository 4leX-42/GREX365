using Grex365.Core.Models;
using Microsoft.Graph;

namespace Grex365.Core.Abstractions;

public interface IGraphConnection
{
    bool IsConnected { get; }

    GraphServiceClient? Client { get; }

    Task ConnectByCertificateAsync(
        CertConfig config,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    Task DisconnectAsync(CancellationToken cancellationToken = default);
}
