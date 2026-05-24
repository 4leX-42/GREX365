using FluentAssertions;
using Grex365.App.Services;
using Grex365.Core.Models;
using Xunit;

namespace Grex365.App.Tests;

public class UiLogSinkNotificationTests
{
    private sealed class FakeNotifier : INotifier
    {
        public List<(string Title, string Message, LogSeverity Severity)> Calls { get; } = new();
        public void Notify(string title, string message, LogSeverity severity) =>
            Calls.Add((title, message, severity));
    }

    private static UiLogSink MakeSink(FakeNotifier? notifier = null) => new(notifier);

    // Note: UiLogSink's Progress<T> requires UI synchronization context. In a unit test
    // we cannot reliably trigger OnEntry via Progress.Report; instead we cover the policy
    // matrix by direct call when possible. For now we exercise via reflection of the
    // private ShouldNotify gate to avoid SynchronizationContext flakes.

    private static bool InvokeShouldNotify(UiLogSink sink, LogEntry entry)
    {
        var method = typeof(UiLogSink).GetMethod("ShouldNotify",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        return (bool)method!.Invoke(sink, new object[] { entry })!;
    }

    [Fact]
    public void ShouldNotify_NullNotifier_ReturnsFalse()
    {
        var sink = MakeSink(notifier: null);
        var entry = LogEntry.Error("X", "boom");

        InvokeShouldNotify(sink, entry).Should().BeFalse();
    }

    [Theory]
    [InlineData("AutoConnect")]
    [InlineData("Connect")]
    [InlineData("ConnectionMonitor")]
    [InlineData("TenantLock")]
    [InlineData("Settings")]
    [InlineData("EXO")]
    [InlineData("Graph")]
    [InlineData("autoconnect")] // case-insensitive
    [InlineData("CONNECT")]
    public void ShouldNotify_SilentSources_ReturnsFalse(string source)
    {
        var sink = MakeSink(new FakeNotifier());
        var error = LogEntry.Error(source, "boom");
        var warn = LogEntry.Warn(source, "careful");

        InvokeShouldNotify(sink, error).Should().BeFalse(because: $"{source} is in silent set");
        InvokeShouldNotify(sink, warn).Should().BeFalse();
    }

    [Theory]
    [InlineData("Users")]
    [InlineData("Groups")]
    [InlineData("Audit")]
    [InlineData("Offboarding")]
    public void ShouldNotify_NonSilentSource_Warning_ReturnsTrue(string source)
    {
        var sink = MakeSink(new FakeNotifier());
        var entry = LogEntry.Warn(source, "rbac denied");

        InvokeShouldNotify(sink, entry).Should().BeTrue();
    }

    [Theory]
    [InlineData("Users")]
    [InlineData("Groups")]
    public void ShouldNotify_NonSilentSource_Error_ReturnsTrue(string source)
    {
        var sink = MakeSink(new FakeNotifier());
        var entry = LogEntry.Error(source, "service error");

        InvokeShouldNotify(sink, entry).Should().BeTrue();
    }

    [Theory]
    [InlineData(LogSeverity.Info)]
    [InlineData(LogSeverity.Debug)]
    public void ShouldNotify_InfoOrDebug_ReturnsFalse(LogSeverity severity)
    {
        var sink = MakeSink(new FakeNotifier());
        var entry = new LogEntry(DateTime.UtcNow, severity, "Users", "msg", null);

        InvokeShouldNotify(sink, entry).Should().BeFalse(because: "info/debug stays in log panel only");
    }

    [Fact]
    public void ShouldNotify_Ok_NonSilent_ReturnsTrue_ButOkPathSkipsNotifyInOnEntry()
    {
        // Ok severity passes the source-and-severity gate, but the switch in OnEntry
        // explicitly skips notifier call for Ok (success goes to status bar / log
        // panel). The gate itself answers "is this a candidate" — Ok happens to also
        // be valid for explicit caller-driven notify if anyone needs it. Verify gate
        // logic only here.
        var sink = MakeSink(new FakeNotifier());
        var entry = LogEntry.Ok("Users", "saved");

        InvokeShouldNotify(sink, entry).Should().BeTrue();
    }
}
