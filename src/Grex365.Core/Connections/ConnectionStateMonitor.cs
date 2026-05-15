using System.ComponentModel;
using System.Runtime.CompilerServices;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.Core.Connections;

public sealed class ConnectionStateMonitor : IConnectionStateMonitor
{
    private readonly IGraphConnection _graph;
    private readonly IExchangeConnection _exchange;
    private readonly TimeSpan _pollInterval;
    private CancellationTokenSource? _cts;
    private Task? _loop;
    private ConnectionState _current = ConnectionState.Disconnected;

    public ConnectionStateMonitor(IGraphConnection graph, IExchangeConnection exchange)
        : this(graph, exchange, TimeSpan.FromSeconds(1))
    {
    }

    public ConnectionStateMonitor(IGraphConnection graph, IExchangeConnection exchange, TimeSpan pollInterval)
    {
        _graph = graph;
        _exchange = exchange;
        _pollInterval = pollInterval;
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public ConnectionState Current
    {
        get => _current;
        private set
        {
            if (_current == value)
            {
                return;
            }
            _current = value;
            OnPropertyChanged();
        }
    }

    public void Start()
    {
        if (_loop is not null)
        {
            return;
        }
        _cts = new CancellationTokenSource();
        _loop = Task.Run(() => PollLoopAsync(_cts.Token));
    }

    public void Stop()
    {
        _cts?.Cancel();
    }

    private async Task PollLoopAsync(CancellationToken ct)
    {
        while (!ct.IsCancellationRequested)
        {
            try
            {
                Current = new ConnectionState(
                    GraphConnected: _graph.IsConnected,
                    ExchangeConnected: _exchange.IsConnected,
                    TenantId: null,
                    TenantDomain: null,
                    Account: null);
            }
            catch
            {
                // swallow; next tick retries.
            }

            try
            {
                await Task.Delay(_pollInterval, ct).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                break;
            }
        }
    }

    private void OnPropertyChanged([CallerMemberName] string? property = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property));
    }

    public async ValueTask DisposeAsync()
    {
        Stop();
        if (_loop is not null)
        {
            try
            {
                await _loop.ConfigureAwait(false);
            }
            catch
            {
                // ignore
            }
        }
        _cts?.Dispose();
    }
}
