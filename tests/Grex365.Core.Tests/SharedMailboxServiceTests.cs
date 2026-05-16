using FluentAssertions;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Grex365.PowerShell;
using Moq;

namespace Grex365.Core.Tests;

public class SharedMailboxServiceTests
{
    private static PowerShellResult OkEmpty() =>
        new(Success: true, Output: Array.Empty<object?>(), Errors: Array.Empty<string>());

    private static PowerShellResult Failed(params string[] errors) =>
        new(Success: false, Output: Array.Empty<object?>(), Errors: errors);

    private static Mock<IPowerShellRunner> Runner(PowerShellResult result)
    {
        var mock = new Mock<IPowerShellRunner>();
        mock.Setup(r => r.RunAsync(
                It.IsAny<string>(),
                It.IsAny<IDictionary<string, object?>>(),
                It.IsAny<IProgress<LogEntry>>(),
                It.IsAny<CancellationToken>()))
            .ReturnsAsync(result);
        return mock;
    }

    [Theory]
    [InlineData("", "FullAccess", "mbx", "principal")]
    [InlineData("invalid", "FullAccess", "mbx", "principal")]
    public async Task ApplyPermission_InvalidAction_ReturnsInvalido(string action, string perm, string mbx, string prn)
    {
        var sut = new SharedMailboxService(Runner(OkEmpty()).Object);
        var r = await sut.ApplyPermissionAsync(action, perm, mbx, prn);
        r.Status.Should().Be("INVALIDO");
    }

    [Fact]
    public async Task ApplyPermission_InvalidPermission_ReturnsInvalido()
    {
        var sut = new SharedMailboxService(Runner(OkEmpty()).Object);
        var r = await sut.ApplyPermissionAsync("add", "Bogus", "mbx@a", "prn@a");
        r.Status.Should().Be("INVALIDO");
    }

    [Theory]
    [InlineData("", "p@a")]
    [InlineData("m@a", "")]
    [InlineData(" ", "p@a")]
    public async Task ApplyPermission_EmptyMailboxOrPrincipal_ReturnsInvalido(string mbx, string prn)
    {
        var sut = new SharedMailboxService(Runner(OkEmpty()).Object);
        var r = await sut.ApplyPermissionAsync("add", "FullAccess", mbx, prn);
        r.Status.Should().Be("INVALIDO");
    }

    [Theory]
    [InlineData("add", "FullAccess", "Add-MailboxPermission")]
    [InlineData("remove", "FullAccess", "Remove-MailboxPermission")]
    [InlineData("add", "SendAs", "Add-RecipientPermission")]
    [InlineData("remove", "SendAs", "Remove-RecipientPermission")]
    [InlineData("add", "SendOnBehalf", "GrantSendOnBehalfTo")]
    [InlineData("remove", "SendOnBehalf", "GrantSendOnBehalfTo")]
    public async Task ApplyPermission_BuildsExpectedCmdlet(string action, string perm, string expectedCmdlet)
    {
        var runner = Runner(OkEmpty());
        var sut = new SharedMailboxService(runner.Object);

        var r = await sut.ApplyPermissionAsync(action, perm, "mbx@a", "prn@a");

        r.Status.Should().Be("OK");
        runner.Verify(x => x.RunAsync(
            It.Is<string>(s => s.Contains(expectedCmdlet)),
            It.Is<IDictionary<string, object?>>(p => (string)p["Mailbox"]! == "mbx@a" && (string)p["Principal"]! == "prn@a"),
            It.IsAny<IProgress<LogEntry>>(),
            It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public async Task ApplyPermission_RunnerFails_ReturnsError()
    {
        var sut = new SharedMailboxService(Runner(Failed("ouch")).Object);
        var r = await sut.ApplyPermissionAsync("add", "FullAccess", "m@a", "p@a");
        r.Status.Should().Be("ERROR");
        r.Detail.Should().Contain("ouch");
    }

    [Fact]
    public async Task GetMailbox_NoOutput_ReturnsNull()
    {
        var sut = new SharedMailboxService(Runner(OkEmpty()).Object);
        var info = await sut.GetMailboxAsync("user@a");
        info.Should().BeNull();
    }

    [Fact]
    public async Task GetMailbox_WithOutput_MapsFields()
    {
        var fake = new FakeMailbox("u@a", "User A", "u@a", "SharedMailbox");
        var result = new PowerShellResult(true, new object?[] { fake }, Array.Empty<string>());
        var sut = new SharedMailboxService(Runner(result).Object);

        var info = await sut.GetMailboxAsync("u@a");

        info.Should().NotBeNull();
        info!.Identity.Should().Be("u@a");
        info.DisplayName.Should().Be("User A");
        info.PrimarySmtpAddress.Should().Be("u@a");
        info.RecipientTypeDetails.Should().Be("SharedMailbox");
    }

    [Fact]
    public async Task ConvertToRegular_RunnerFails_Throws()
    {
        var sut = new SharedMailboxService(Runner(Failed("fail")).Object);
        Func<Task> act = () => sut.ConvertToRegularAsync("u@a");
        await act.Should().ThrowAsync<InvalidOperationException>();
    }

    [Fact]
    public async Task ConvertToShared_RunnerFails_Throws()
    {
        var sut = new SharedMailboxService(Runner(Failed("EXO blocked")).Object);
        Func<Task> act = () => sut.ConvertToSharedAsync("u@a");
        await act.Should().ThrowAsync<InvalidOperationException>();
    }

    [Fact]
    public async Task GetPermissions_Empty_ReturnsEmptyList()
    {
        var sut = new SharedMailboxService(Runner(OkEmpty()).Object);
        var perms = await sut.GetPermissionsAsync("m@a");
        perms.Should().BeEmpty();
    }

    [Fact]
    public async Task GetPermissions_MapsEntries()
    {
        var fake1 = new FakePermission("FullAccess", "user1@a", "FullAccess");
        var fake2 = new FakePermission("SendAs",     "user2@a", "SendAs");
        var fake3 = new FakePermission("SendOnBehalf", "user3@a", "From Set-Mailbox");
        var result = new Grex365.Core.Abstractions.PowerShellResult(true,
            new object?[] { fake1, fake2, fake3 },
            Array.Empty<string>());
        var sut = new SharedMailboxService(Runner(result).Object);

        var perms = await sut.GetPermissionsAsync("m@a");
        perms.Should().HaveCount(3);
        perms[0].Permission.Should().Be("FullAccess");
        perms[1].Permission.Should().Be("SendAs");
        perms[2].Permission.Should().Be("SendOnBehalf");
        perms[2].Principal.Should().Be("user3@a");
    }

    [Fact]
    public async Task GetPermissions_RunnerFails_Throws()
    {
        var sut = new SharedMailboxService(Runner(Failed("get failed")).Object);
        Func<Task> act = () => sut.GetPermissionsAsync("m@a");
        await act.Should().ThrowAsync<InvalidOperationException>();
    }

    private sealed record FakeMailbox(string Identity, string DisplayName, string PrimarySmtpAddress, string RecipientTypeDetails);
    private sealed record FakePermission(string Permission, string Principal, string Detail);
}
