using FluentAssertions;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Grex365.Core.Offboarding;
using Moq;

namespace Grex365.Core.Tests;

public class OffboardingServiceTests
{
    private static UserSummary SampleUser() =>
        new("uid", "Jane Doe", "jane@a", "jane@a", true, false, 2, null);

    private static Mock<IUsersService> UsersOk()
    {
        var m = new Mock<IUsersService>();
        m.Setup(u => u.GetByIdAsync(It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(SampleUser());
        return m;
    }

    private static Mock<ISharedMailboxService> MailboxOk()
    {
        var m = new Mock<ISharedMailboxService>();
        m.Setup(s => s.ConvertToSharedAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new MailboxInfo("u@a", "U", "u@a", "SharedMailbox"));
        return m;
    }

    [Fact]
    public async Task EmptyUpn_ReturnsError()
    {
        var sut = new OffboardingService(UsersOk().Object, MailboxOk().Object);
        var r = await sut.RunAsync("", new OffboardingOptions(true, true, true));
        r.Success.Should().BeFalse();
        r.Steps.Should().Contain(s => s.Status == "ERROR");
    }

    [Fact]
    public async Task UserNotFound_ReturnsError()
    {
        var users = new Mock<IUsersService>();
        users.Setup(u => u.GetByIdAsync(It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((UserSummary?)null);
        var sut = new OffboardingService(users.Object, MailboxOk().Object);
        var r = await sut.RunAsync("ghost@a", new OffboardingOptions(true, true, false));
        r.Success.Should().BeFalse();
        r.Steps.Should().ContainSingle(s => s.Status == "ERROR" && s.Name.Contains("usuario"));
    }

    [Fact]
    public async Task AllStepsEnabled_AllRun()
    {
        var users = UsersOk();
        var mbx = MailboxOk();
        var sut = new OffboardingService(users.Object, mbx.Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(true, true, true));

        r.Success.Should().BeTrue();
        r.Steps.Should().HaveCount(4); // find + disable + licenses + mailbox
        users.Verify(u => u.SetAccountEnabledAsync("uid", false, It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
        users.Verify(u => u.RemoveAllLicensesAsync("uid", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
        mbx.Verify(s => s.ConvertToSharedAsync("jane@a", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task DisableOnly_SkipsOthers()
    {
        var users = UsersOk();
        var mbx = MailboxOk();
        var sut = new OffboardingService(users.Object, mbx.Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(DisableAccount: true, RemoveLicenses: false, ConvertMailboxToShared: false));

        r.Steps.Should().HaveCount(2); // find + disable
        users.Verify(u => u.SetAccountEnabledAsync(It.IsAny<string>(), false, It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
        users.Verify(u => u.RemoveAllLicensesAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
        mbx.Verify(s => s.ConvertToSharedAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task DisableFails_OtherStepsStillRunAndSuccessFalse()
    {
        var users = UsersOk();
        users.Setup(u => u.SetAccountEnabledAsync(It.IsAny<string>(), It.IsAny<bool>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("forbidden"));
        var sut = new OffboardingService(users.Object, MailboxOk().Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(true, true, true));

        r.Success.Should().BeFalse();
        r.Steps.Should().Contain(s => s.Name.Contains("Deshabilitar") && s.Status == "ERROR");
        users.Verify(u => u.RemoveAllLicensesAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task MailboxFails_ReportsErrorButDoesNotThrow()
    {
        var users = UsersOk();
        var mbx = new Mock<ISharedMailboxService>();
        mbx.Setup(s => s.ConvertToSharedAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("EXO down"));

        var sut = new OffboardingService(users.Object, mbx.Object);
        var r = await sut.RunAsync("jane@a", new OffboardingOptions(false, false, true));

        r.Success.Should().BeFalse();
        r.Steps.Should().Contain(s => s.Name.StartsWith("Mailbox") && s.Status == "ERROR" && s.Detail.Contains("EXO down"));
    }
}
