using FluentAssertions;
using Grex365.App.Services;
using Grex365.App.ViewModels;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Grex365.Core.Plugins;
using Moq;
using Serilog.Core;
using Serilog.Events;

namespace Grex365.App.Tests;

public class UpdateFeedResolverTests
{
    [Theory]
    [InlineData(null, UpdateFeedKind.None)]
    [InlineData("", UpdateFeedKind.None)]
    [InlineData("   ", UpdateFeedKind.None)]
    [InlineData("no-es-url", UpdateFeedKind.None)]
    [InlineData("ftp://server/feed", UpdateFeedKind.None)]
    [InlineData("https://github.com/org/grex365-releases", UpdateFeedKind.GitHub)]
    [InlineData("https://GITHUB.com/org/repo", UpdateFeedKind.GitHub)]
    [InlineData("https://enterprise.github.com/org/repo", UpdateFeedKind.GitHub)]
    [InlineData("https://share.interno.local/grex365/releases", UpdateFeedKind.Http)]
    [InlineData("http://10.0.0.5/updates", UpdateFeedKind.Http)]
    [InlineData("  https://github.com/o/r  ", UpdateFeedKind.GitHub)]
    [InlineData(@"C:\feeds\grex365", UpdateFeedKind.LocalPath)]
    [InlineData(@"\\fileserver\soft\grex365", UpdateFeedKind.LocalPath)]
    [InlineData("file://fileserver/soft/grex365", UpdateFeedKind.LocalPath)]
    public void Classify_RoutesByUrl(string? url, UpdateFeedKind expected) =>
        UpdateFeedResolver.Classify(url).Should().Be(expected);

    [Theory]
    [InlineData(@"C:\feeds\grex365", @"C:\feeds\grex365")]
    [InlineData(@"\\fileserver\soft\grex365", @"\\fileserver\soft\grex365")]
    [InlineData("file:///C:/feeds/grex365", @"C:\feeds\grex365")]
    public void ToLocalDirectory_NormalizesFileUris(string url, string expected) =>
        UpdateFeedResolver.ToLocalDirectory(url).Should().Be(expected);
}

// SettingsViewModel update commands: status messages per check outcome, apply gating, feed
// URL persistence. IUpdateService is mocked — no Velopack/network involved.
public class SettingsViewModelUpdatesTests
{
    private sealed class FakeLogSink : IUiLogSink
    {
        public System.Collections.ObjectModel.ObservableCollection<LogEntry> Entries { get; } = new();
        public IProgress<LogEntry> Progress { get; } = new Progress<LogEntry>(_ => { });
        public void Clear() { }
    }

    private static (SettingsViewModel Vm, Mock<IPreferencesStore> Prefs) Build(IUpdateService? updates)
    {
        var prefs = new Mock<IPreferencesStore>();
        prefs.Setup(p => p.LoadAsync(It.IsAny<CancellationToken>())).ReturnsAsync(new UserPreferences());
        var certs = new Mock<ICertConfigStore>();
        certs.Setup(c => c.LoadAsync(It.IsAny<CancellationToken>())).ReturnsAsync((CertConfig?)null);
        var validator = new Mock<ICertValidator>();
        validator.Setup(v => v.Validate(It.IsAny<CertConfig>()))
            .Returns(new CertValidationResult(CertValidationStatus.MissingConfig, "—"));
        var vm = new SettingsViewModel(
            prefs.Object,
            certs.Object,
            validator.Object,
            new FakeLogSink(),
            new PluginLoadReport(new List<DiscoveredPlugin>(), new List<PluginLoadFailure>(), new List<DisabledPlugin>()),
            new LoggingLevelSwitch(LogEventLevel.Information),
            updates);
        return (vm, prefs);
    }

    [Fact]
    public async Task CheckUpdates_NoService_ReportsNotWired()
    {
        var (vm, _) = Build(updates: null);

        await vm.CheckUpdatesCommand.ExecuteAsync(null);

        vm.UpdateStatus.Should().Be(L10n.Get("Settings.Updates.NotWired"));
        vm.UpdateAvailable.Should().BeFalse();
    }

    [Theory]
    [InlineData(UpdateCheckStatus.FeedNotConfigured, "Settings.Updates.NoFeed")]
    [InlineData(UpdateCheckStatus.NotInstalled, "Settings.Updates.NotInstalled")]
    [InlineData(UpdateCheckStatus.UpToDate, "Settings.Updates.UpToDate")]
    public async Task CheckUpdates_StatusMapsToMessage(UpdateCheckStatus status, string expectedKey)
    {
        var updates = new Mock<IUpdateService>();
        updates.Setup(u => u.CheckAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new UpdateCheckResult(status));
        var (vm, _) = Build(updates.Object);

        await vm.CheckUpdatesCommand.ExecuteAsync(null);

        vm.UpdateStatus.Should().Be(L10n.Get(expectedKey));
        vm.UpdateAvailable.Should().BeFalse();
        vm.IsCheckingUpdates.Should().BeFalse();
    }

    [Fact]
    public async Task CheckUpdates_Available_SetsFlagAndVersion()
    {
        var updates = new Mock<IUpdateService>();
        updates.Setup(u => u.CheckAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new UpdateCheckResult(UpdateCheckStatus.UpdateAvailable, "2.1.0"));
        var (vm, _) = Build(updates.Object);

        await vm.CheckUpdatesCommand.ExecuteAsync(null);

        vm.UpdateAvailable.Should().BeTrue();
        vm.UpdateStatus.Should().Contain("2.1.0");
    }

    [Fact]
    public async Task CheckUpdates_Error_ShowsDetail()
    {
        var updates = new Mock<IUpdateService>();
        updates.Setup(u => u.CheckAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new UpdateCheckResult(UpdateCheckStatus.Error, Detail: "DNS down"));
        var (vm, _) = Build(updates.Object);

        await vm.CheckUpdatesCommand.ExecuteAsync(null);

        vm.UpdateAvailable.Should().BeFalse();
        vm.UpdateStatus.Should().Contain("DNS down");
    }

    [Fact]
    public async Task ApplyUpdate_WithoutPendingUpdate_NoOps()
    {
        var updates = new Mock<IUpdateService>(MockBehavior.Strict); // ApplyAsync must NOT be hit
        updates.Setup(u => u.CheckAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new UpdateCheckResult(UpdateCheckStatus.UpToDate));
        var (vm, _) = Build(updates.Object);
        await vm.CheckUpdatesCommand.ExecuteAsync(null);

        await vm.ApplyUpdateCommand.ExecuteAsync(null);

        updates.Verify(u => u.ApplyAsync(It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task ApplyUpdate_Failure_ReportsError()
    {
        var updates = new Mock<IUpdateService>();
        updates.Setup(u => u.CheckAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new UpdateCheckResult(UpdateCheckStatus.UpdateAvailable, "2.1.0"));
        updates.Setup(u => u.ApplyAsync(It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("descarga rota"));
        var (vm, _) = Build(updates.Object);
        await vm.CheckUpdatesCommand.ExecuteAsync(null);

        await vm.ApplyUpdateCommand.ExecuteAsync(null);

        vm.UpdateStatus.Should().Contain("descarga rota");
    }

    [Fact]
    public async Task Save_PersistsTrimmedFeedUrl_NullWhenBlank()
    {
        // Capture the VALUE at save time — the mock hands back the same instance on every load,
        // so storing references would see the second save's mutation.
        var saved = new List<string?>();
        var (vm, prefs) = Build(updates: null);
        prefs.Setup(p => p.SaveAsync(It.IsAny<UserPreferences>(), It.IsAny<CancellationToken>()))
            .Callback<UserPreferences, CancellationToken>((p, _) => saved.Add(p.UpdateFeedUrl))
            .Returns(Task.CompletedTask);

        vm.UpdateFeedUrl = "  https://github.com/org/repo  ";
        await vm.SaveCommand.ExecuteAsync(null);
        vm.UpdateFeedUrl = "   ";
        await vm.SaveCommand.ExecuteAsync(null);

        saved.Should().HaveCount(2);
        saved[0].Should().Be("https://github.com/org/repo");
        saved[1].Should().BeNull();
    }
}
