using Grex365.Core.Abstractions;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.Extensibility;

namespace Grex365.App.Services;

public sealed class ApplicationInsightsTelemetry : ITelemetry, IDisposable
{
    private readonly TelemetryClient _client;
    private readonly TelemetryConfiguration _config;

    public ApplicationInsightsTelemetry(string connectionString)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(connectionString);
        _config = TelemetryConfiguration.CreateDefault();
        _config.ConnectionString = connectionString;
        _client = new TelemetryClient(_config);
        _client.Context.GlobalProperties["app"] = "Grex365";
        _client.Context.GlobalProperties["host"] = Environment.MachineName;
        _client.Context.User.AuthenticatedUserId = Environment.UserName;
    }

    public bool IsEnabled => true;

    public void TrackEvent(string name, IDictionary<string, string>? properties = null)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return;
        }
        _client.TrackEvent(name, properties);
    }

    public void TrackException(Exception exception, IDictionary<string, string>? properties = null)
    {
        ArgumentNullException.ThrowIfNull(exception);
        _client.TrackException(exception, properties);
    }

    public void Flush() => _client.Flush();

    public void Dispose()
    {
        try { _client.Flush(); } catch { /* swallow on shutdown */ }
        _config.Dispose();
    }
}
