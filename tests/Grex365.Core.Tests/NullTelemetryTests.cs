using FluentAssertions;
using Grex365.Core.Abstractions;

namespace Grex365.Core.Tests;

public class NullTelemetryTests
{
    [Fact]
    public void IsEnabled_False()
    {
        new NullTelemetry().IsEnabled.Should().BeFalse();
    }

    [Fact]
    public void AllMethods_DoNotThrow_WithNullArgs()
    {
        var t = new NullTelemetry();
        var act = () =>
        {
            t.TrackEvent("evt");
            t.TrackEvent("evt", null);
            t.TrackException(new InvalidOperationException("x"));
            t.TrackException(new InvalidOperationException("x"), null);
            t.Flush();
        };
        act.Should().NotThrow();
    }

    [Fact]
    public void TrackEvent_WithProperties_NoThrow()
    {
        var t = new NullTelemetry();
        var props = new Dictionary<string, string> { ["k"] = "v" };
        var act = () => t.TrackEvent("evt", props);
        act.Should().NotThrow();
    }
}
