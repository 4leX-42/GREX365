using FluentAssertions;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Grex365.PowerShell;
using Moq;

namespace Grex365.Core.Tests;

public class SharedMailboxServiceTests
{
    private static (SharedMailboxService Sut, Mock<IExternalExoOps> Exo) Make()
    {
        var exo = new Mock<IExternalExoOps>();
        return (new SharedMailboxService(exo.Object), exo);
    }

    private static void VerifyNoApply(Mock<IExternalExoOps> exo) =>
        exo.Verify(e => e.ApplyPermissionAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>(),
            It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);

    // ----- ApplyPermission validation stays in the wrapper (no EXO call) -----
    [Theory]
    [InlineData("", "FullAccess", "mbx", "principal")]
    [InlineData("invalid", "FullAccess", "mbx", "principal")]
    public async Task ApplyPermission_InvalidAction_ReturnsInvalido(string action, string perm, string mbx, string prn)
    {
        var (sut, exo) = Make();
        var r = await sut.ApplyPermissionAsync(action, perm, mbx, prn);
        r.Status.Should().Be("INVALIDO");
        VerifyNoApply(exo);
    }

    [Fact]
    public async Task ApplyPermission_InvalidPermission_ReturnsInvalido()
    {
        var (sut, exo) = Make();
        var r = await sut.ApplyPermissionAsync("add", "Bogus", "mbx@a", "prn@a");
        r.Status.Should().Be("INVALIDO");
        VerifyNoApply(exo);
    }

    [Theory]
    [InlineData("", "p@a")]
    [InlineData("m@a", "")]
    [InlineData(" ", "p@a")]
    public async Task ApplyPermission_EmptyMailboxOrPrincipal_ReturnsInvalido(string mbx, string prn)
    {
        var (sut, exo) = Make();
        var r = await sut.ApplyPermissionAsync("add", "FullAccess", mbx, prn);
        r.Status.Should().Be("INVALIDO");
        VerifyNoApply(exo);
    }

    [Fact]
    public async Task ApplyPermission_Valid_DelegatesToExternalExo_WithLowercasedAction()
    {
        var (sut, exo) = Make();
        exo.Setup(e => e.ApplyPermissionAsync("add", "FullAccess", "mbx@a", "prn@a", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new MailboxPermissionResult("add", "FullAccess", "mbx@a", "prn@a", "OK", "Aplicado"));

        var r = await sut.ApplyPermissionAsync("Add", "FullAccess", "mbx@a", "prn@a");

        r.Status.Should().Be("OK");
        exo.Verify(e => e.ApplyPermissionAsync("add", "FullAccess", "mbx@a", "prn@a", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    // ----- delegation to the external EXO host -----
    [Fact]
    public async Task GetMailbox_DelegatesToExternalExo()
    {
        var (sut, exo) = Make();
        exo.Setup(e => e.GetMailboxFactsAsync("u@a", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new MailboxInfo("u@a", "User A", "u@a", "SharedMailbox"));

        var info = await sut.GetMailboxAsync("u@a");

        info!.RecipientTypeDetails.Should().Be("SharedMailbox");
        exo.Verify(e => e.GetMailboxFactsAsync("u@a", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task ConvertToShared_DelegatesToExternalExo()
    {
        var (sut, exo) = Make();
        await sut.ConvertToSharedAsync("u@a");
        exo.Verify(e => e.ConvertToSharedAsync("u@a", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task ConvertToRegular_DelegatesToExternalExo()
    {
        var (sut, exo) = Make();
        await sut.ConvertToRegularAsync("u@a");
        exo.Verify(e => e.ConvertToRegularAsync("u@a", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task GetPermissions_DelegatesToExternalExo()
    {
        var (sut, exo) = Make();
        exo.Setup(e => e.GetPermissionsAsync("m@a", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new List<MailboxPermissionEntry> { new("FullAccess", "x@a", "FullAccess") });

        var perms = await sut.GetPermissionsAsync("m@a");

        perms.Should().HaveCount(1);
        exo.Verify(e => e.GetPermissionsAsync("m@a", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    // ----- pure cmdlet mapping (now in ExternalExoOps.BuildPermissionCmdlet) -----
    [Theory]
    [InlineData("add", "FullAccess", "Add-MailboxPermission")]
    [InlineData("remove", "FullAccess", "Remove-MailboxPermission")]
    [InlineData("add", "SendAs", "Add-RecipientPermission")]
    [InlineData("remove", "SendAs", "Remove-RecipientPermission")]
    [InlineData("add", "SendOnBehalf", "GrantSendOnBehalfTo")]
    [InlineData("remove", "SendOnBehalf", "GrantSendOnBehalfTo")]
    public void BuildPermissionCmdlet_MapsExpectedCmdlet(string action, string perm, string expected)
    {
        ExternalExoOps.BuildPermissionCmdlet(action, perm).Should().Contain(expected);
    }

    [Fact]
    public void BuildPermissionCmdlet_Unsupported_ReturnsNull()
    {
        ExternalExoOps.BuildPermissionCmdlet("add", "Bogus").Should().BeNull();
    }
}
