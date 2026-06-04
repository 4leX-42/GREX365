using FluentAssertions;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Grex365.PowerShell;
using Moq;

namespace Grex365.Core.Tests;

// MailboxRulesService is now a thin wrapper over IExternalExoOps (external pwsh) — the in-proc
// RunspacePool path tripped the GetResponseHeader bug. Pure validation stays in the wrapper and
// must short-circuit before any EXO call; everything else delegates to the external host.
public class MailboxRulesServiceTests
{
    private static (MailboxRulesService Sut, Mock<IExternalExoOps> Exo) Make()
    {
        var exo = new Mock<IExternalExoOps>();
        return (new MailboxRulesService(exo.Object), exo);
    }

    // ----- auto-reply: validation short-circuits, no EXO call -----
    [Fact]
    public async Task SetAutoReply_Invalid_Throws_NoExoCall()
    {
        var (sut, exo) = Make();
        var bad = new AutoReplyConfig(AutoReplyState.Enabled, null, null, null, null); // no message

        var act = () => sut.SetAutoReplyAsync("u@a", bad);

        await act.Should().ThrowAsync<ArgumentException>();
        exo.Verify(e => e.SetAutoReplyConfigAsync(It.IsAny<string>(), It.IsAny<AutoReplyConfig>(),
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task SetAutoReply_Valid_DelegatesToExternalExo()
    {
        var (sut, exo) = Make();
        var cfg = new AutoReplyConfig(AutoReplyState.Enabled, "fuera", null, null, null);

        await sut.SetAutoReplyAsync("u@a", cfg);

        exo.Verify(e => e.SetAutoReplyConfigAsync("u@a", cfg,
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task GetAutoReply_DelegatesToExternalExo()
    {
        var (sut, exo) = Make();
        exo.Setup(e => e.GetAutoReplyAsync("u@a", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new AutoReplyConfig(AutoReplyState.Enabled, "x", null, null, null));

        var r = await sut.GetAutoReplyAsync("u@a");

        r!.State.Should().Be(AutoReplyState.Enabled);
        exo.Verify(e => e.GetAutoReplyAsync("u@a", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    // ----- forwarding: validation short-circuits, no EXO call -----
    [Fact]
    public async Task SetForwarding_InvalidSmtp_Throws_NoExoCall()
    {
        var (sut, exo) = Make();

        var act = () => sut.SetForwardingAsync("u@a", "notanemail", true);

        await act.Should().ThrowAsync<ArgumentException>();
        exo.Verify(e => e.ConfigureForwardingAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<bool>(),
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task SetForwarding_Valid_DelegatesWithDeliverFlag()
    {
        var (sut, exo) = Make();

        await sut.SetForwardingAsync("u@a", "ext@b.com", deliverToMailboxAndForward: true);

        exo.Verify(e => e.ConfigureForwardingAsync("u@a", "ext@b.com", true,
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task ClearForwarding_DelegatesToExternalExo()
    {
        var (sut, exo) = Make();

        await sut.ClearForwardingAsync("u@a");

        exo.Verify(e => e.ClearForwardingAsync("u@a",
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task GetForwarding_DelegatesToExternalExo()
    {
        var (sut, exo) = Make();
        exo.Setup(e => e.GetForwardingAsync("u@a", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new ForwardingConfig(null, "ext@b.com", true));

        var r = await sut.GetForwardingAsync("u@a");

        r!.ForwardingSmtpAddress.Should().Be("ext@b.com");
        exo.Verify(e => e.GetForwardingAsync("u@a", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    // ----- calendar: validation short-circuits, no EXO call -----
    [Theory]
    [InlineData("", "jane@a", "Reviewer")]
    [InlineData("u@a", "", "Reviewer")]
    [InlineData("u@a", "jane@a", "Bogus")]
    [InlineData("u@a", "jane@a", "")]
    public async Task ApplyCalendarPermission_Invalid_Throws_NoExoCall(string id, string principal, string rights)
    {
        var (sut, exo) = Make();

        var act = () => sut.ApplyCalendarPermissionAsync(id, principal, rights);

        await act.Should().ThrowAsync<ArgumentException>();
        exo.Verify(e => e.ApplyCalendarPermissionAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>(),
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task ApplyCalendarPermission_Valid_DelegatesToExternalExo()
    {
        var (sut, exo) = Make();

        await sut.ApplyCalendarPermissionAsync("u@a", "jane@a", CalendarAccessRights.Reviewer);

        exo.Verify(e => e.ApplyCalendarPermissionAsync("u@a", "jane@a", "Reviewer",
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task RemoveCalendarPermission_EmptyPrincipal_Throws_NoExoCall()
    {
        var (sut, exo) = Make();

        var act = () => sut.RemoveCalendarPermissionAsync("u@a", "");

        await act.Should().ThrowAsync<ArgumentException>();
        exo.Verify(e => e.RemoveCalendarPermissionAsync(It.IsAny<string>(), It.IsAny<string>(),
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task RemoveCalendarPermission_Valid_DelegatesToExternalExo()
    {
        var (sut, exo) = Make();

        await sut.RemoveCalendarPermissionAsync("u@a", "jane@a");

        exo.Verify(e => e.RemoveCalendarPermissionAsync("u@a", "jane@a",
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task GetCalendarPermissions_DelegatesToExternalExo()
    {
        var (sut, exo) = Make();
        exo.Setup(e => e.GetCalendarPermissionsAsync("u@a", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new List<CalendarPermissionEntry> { new("jane@a", "Reviewer") });

        var r = await sut.GetCalendarPermissionsAsync("u@a");

        r.Should().HaveCount(1);
        exo.Verify(e => e.GetCalendarPermissionsAsync("u@a", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }
}
