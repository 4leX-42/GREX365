using FluentAssertions;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Grex365.PowerShell;
using Moq;

namespace Grex365.Core.Tests;

// ExoForwardingAuditService now runs each scan as one external pwsh invocation returning JSON.
// These tests mock IExternalExoRunner with canned JSON and assert the parse -> pure-analyzer path
// (the integration risk lives in the JSON shapes, not the analyzers, which are tested separately).
public class ExoForwardingAuditServiceTests
{
    private static (ExoForwardingAuditService Sut, Mock<IExternalExoRunner> Runner) Make(string? json, bool connected = true)
    {
        var runner = new Mock<IExternalExoRunner>();
        runner.Setup(r => r.RunAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(json);
        var exo = new Mock<IExchangeConnection>();
        exo.SetupGet(e => e.IsConnected).Returns(connected);
        return (new ExoForwardingAuditService(runner.Object, exo.Object), runner);
    }

    // ----- external forwarding -----
    [Fact]
    public async Task ScanExternalForwarding_FlagsForwardToUnacceptedDomain()
    {
        var json = """
            {"Domains":["contoso.com"],"Rows":[{"UserPrincipalName":"u@contoso.com","ForwardingSmtpAddress":"thief@evil.com","ForwardingAddress":""}]}
            """;
        var (sut, _) = Make(json);

        var findings = await sut.ScanExternalForwardingAsync();

        findings.Should().NotBeEmpty();
    }

    [Fact]
    public async Task ScanExternalForwarding_TolerantOfSingleElementCollapse()
    {
        // ConvertTo-Json collapses single-element arrays: Domains scalar, Rows a lone object.
        var json = """
            {"Domains":"contoso.com","Rows":{"UserPrincipalName":"u@contoso.com","ForwardingSmtpAddress":"thief@evil.com"}}
            """;
        var (sut, _) = Make(json);

        var findings = await sut.ScanExternalForwardingAsync();

        findings.Should().NotBeEmpty();
    }

    [Fact]
    public async Task ScanExternalForwarding_AllInternal_NoFindings()
    {
        var json = """
            {"Domains":["contoso.com"],"Rows":[{"UserPrincipalName":"u@contoso.com","ForwardingSmtpAddress":"boss@contoso.com"}]}
            """;
        var (sut, _) = Make(json);

        var findings = await sut.ScanExternalForwardingAsync();

        findings.Should().BeEmpty();
    }

    // ----- inbox rules -----
    [Fact]
    public async Task ScanInboxRules_ZeroMax_Throws()
    {
        var (sut, runner) = Make("{}");
        var act = () => sut.ScanInboxRulesAsync(0);
        await act.Should().ThrowAsync<ArgumentOutOfRangeException>();
        runner.Verify(r => r.RunAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task ScanInboxRules_FlagsEnabledDeleteRule()
    {
        var json = """
            {"Domains":["contoso.com"],"Rules":[{"MailboxUpn":"u@contoso.com","Name":"nuke","Enabled":true,"DeleteMessage":true,"ForwardTo":[],"RedirectTo":[]}]}
            """;
        var (sut, _) = Make(json);

        var findings = await sut.ScanInboxRulesAsync(50);

        findings.Should().NotBeEmpty();
    }

    [Fact]
    public async Task ScanInboxRules_NullJson_Empty()
    {
        var (sut, _) = Make(null);
        var findings = await sut.ScanInboxRulesAsync(50);
        findings.Should().BeEmpty();
    }

    // ----- transport rules -----
    [Fact]
    public async Task ScanTransportRules_FlagsEnabledExternalForward_AndCounts()
    {
        var json = """
            {"Domains":["contoso.com"],"Rules":[{"Name":"exfil","State":"Enabled","Priority":0,"Mode":"Enforce","ForwardTo":["thief@evil.com"],"BlindCopyTo":[],"RedirectMessageTo":[],"DeleteMessage":false}]}
            """;
        var (sut, _) = Make(json);

        var (summary, findings) = await sut.ScanTransportRulesAsync();

        summary.Total.Should().Be(1);
        summary.Enabled.Should().Be(1);
        summary.WithExternalForward.Should().Be(1);
        findings.Should().NotBeEmpty();
    }

    [Fact]
    public async Task ScanTransportRules_NullJson_EmptySummary()
    {
        var (sut, _) = Make(null);
        var (summary, findings) = await sut.ScanTransportRulesAsync();
        summary.Total.Should().Be(0);
        findings.Should().BeEmpty();
    }

    // ----- shared mailbox sign-in -----
    [Fact]
    public async Task ScanSharedMailboxSignIn_FlagsSignInEnabled()
    {
        var json = """
            [{"UserPrincipalName":"shared@contoso.com","DisplayName":"Shared","AccountDisabled":false}]
            """;
        var (sut, _) = Make(json);

        var (summary, findings) = await sut.ScanSharedMailboxSignInAsync();

        summary.Total.Should().Be(1);
        summary.SignInEnabled.Should().Be(1);
        findings.Should().NotBeEmpty();
    }

    [Fact]
    public async Task ScanSharedMailboxSignIn_TolerantOfSingleObject_AndDisabledOk()
    {
        // Single shared mailbox → ConvertTo-Json may emit a lone object, not an array.
        var json = """
            {"UserPrincipalName":"shared@contoso.com","DisplayName":"Shared","AccountDisabled":true}
            """;
        var (sut, _) = Make(json);

        var (summary, findings) = await sut.ScanSharedMailboxSignInAsync();

        summary.Total.Should().Be(1);
        summary.SignInDisabled.Should().Be(1);
        findings.Should().BeEmpty();
    }

    [Fact]
    public async Task ScanSharedMailboxSignIn_UnknownWhenAccountDisabledNull()
    {
        var json = """
            [{"UserPrincipalName":"shared@contoso.com","DisplayName":"Shared","AccountDisabled":null}]
            """;
        var (sut, _) = Make(json);

        var (summary, _) = await sut.ScanSharedMailboxSignInAsync();

        summary.Total.Should().Be(1);
        summary.Unknown.Should().Be(1);
    }

    // ----- connection gate -----
    [Fact]
    public async Task Scan_WhenExchangeNotConnected_Throws_NoRunnerCall()
    {
        var (sut, runner) = Make("{}", connected: false);

        var act = () => sut.ScanExternalForwardingAsync();

        await act.Should().ThrowAsync<InvalidOperationException>();
        runner.Verify(r => r.RunAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }
}
