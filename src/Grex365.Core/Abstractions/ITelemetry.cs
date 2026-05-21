namespace Grex365.Core.Abstractions;

public interface ITelemetry
{
    bool IsEnabled { get; }

    void TrackEvent(string name, IDictionary<string, string>? properties = null);

    void TrackException(Exception exception, IDictionary<string, string>? properties = null);

    void Flush();
}

public sealed class NullTelemetry : ITelemetry
{
    public bool IsEnabled => false;

    public void TrackEvent(string name, IDictionary<string, string>? properties = null) { }

    public void TrackException(Exception exception, IDictionary<string, string>? properties = null) { }

    public void Flush() { }
}
