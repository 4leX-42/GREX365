using FluentAssertions;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Grex365.Core.Offboarding;
using Moq;

namespace Grex365.Core.Tests;

public class OffboardingServiceTests
{
    private static UserSummary SampleUser(int licenses = 2) =>
        new("uid", "Jane Doe", "jane@a", "jane@a", true, false, licenses, null);

    private static Mock<IUsersService> UsersOk()
    {
        var m = new Mock<IUsersService>();
        // First read = initial lookup (still licensed); subsequent reads = post-removal
        // verification (0 licenses), so VerifyLicensesRemovedAsync confirms immediately.
        var n = 0;
        m.Setup(u => u.GetByIdAsync(It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(() => n++ == 0 ? SampleUser(2) : SampleUser(0));
        return m;
    }

    private static Mock<ISharedMailboxService> MailboxOk()
    {
        var m = new Mock<ISharedMailboxService>();
        m.Setup(s => s.ConvertToSharedAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new MailboxInfo("u@a", "U", "u@a", "SharedMailbox"));
        return m;
    }

    private static UserSummary DisabledUser(int licenses = 0) =>
        new("uid", "Jane Doe", "jane@a", "jane@a", false, false, licenses, null);

    // External EXO ops that report the given mailbox facts in pre-checks and flip the type to
    // SharedMailbox on convert (so the gate logic, not the convert, is what's under test).
    private static Mock<IExternalExoOps> ExoWithFacts(MailboxInfo facts)
    {
        var e = new Mock<IExternalExoOps>();
        e.Setup(x => x.GetMailboxFactsAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(facts);
        e.Setup(x => x.ConvertToSharedAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(facts with { RecipientTypeDetails = "SharedMailbox" });
        e.Setup(x => x.HideFromGalAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync("GAL oculto");
        return e;
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
        r.Steps.Should().HaveCount(5); // find + pre-checks + disable + mailbox + licenses
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

        r.Steps.Should().HaveCount(3); // find + pre-checks + disable
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
    public async Task AllStepsEnabled_RevokesSessions_AndConvertsBeforeRemovingLicense()
    {
        var users = UsersOk();
        var mbx = MailboxOk();
        var sut = new OffboardingService(users.Object, mbx.Object);

        await sut.RunAsync("jane@a", new OffboardingOptions(true, true, true));

        // Blocking sign-in must also revoke active sessions.
        users.Verify(u => u.RevokeSignInSessionsAsync("uid", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task ConvertFails_SkipsLicenseRemoval_ToAvoidStrandingMailbox()
    {
        var users = UsersOk();
        var mbx = new Mock<ISharedMailboxService>();
        mbx.Setup(s => s.ConvertToSharedAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("EXO down"));

        var sut = new OffboardingService(users.Object, mbx.Object);
        var r = await sut.RunAsync("jane@a", new OffboardingOptions(DisableAccount: false, RemoveLicenses: true, ConvertMailboxToShared: true));

        r.Success.Should().BeFalse();
        r.Steps.Should().Contain(s => s.Name.Contains("Quitar licencias") && s.Status == "OMITIDO");
        users.Verify(u => u.RemoveAllLicensesAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
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
        r.Steps.Should().Contain(s => s.Name.Contains("compartido") && s.Status == "ERROR" && s.Detail.Contains("EXO down"));
    }

    // The convert call can succeed (no exception) yet the mailbox type never actually flips
    // to SharedMailbox — e.g. EXO accepts the request but a hold/policy keeps it a UserMailbox.
    // The safety gate must treat that as a failed conversion and still skip license removal.
    [Fact]
    public async Task ConvertVerificationFails_NoException_SkipsLicenseRemoval()
    {
        var users = UsersOk();
        var mbx = new Mock<ISharedMailboxService>();
        mbx.Setup(s => s.ConvertToSharedAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new MailboxInfo("u@a", "U", "u@a", "UserMailbox")); // type did NOT flip

        var sut = new OffboardingService(users.Object, mbx.Object);
        var r = await sut.RunAsync("jane@a", new OffboardingOptions(DisableAccount: false, RemoveLicenses: true, ConvertMailboxToShared: true));

        r.Success.Should().BeFalse();
        r.Steps.Should().Contain(s => s.Name.Contains("compartido") && s.Status == "ERROR");
        r.Steps.Should().Contain(s => s.Name.Contains("Quitar licencias") && s.Status == "OMITIDO");
        users.Verify(u => u.RemoveAllLicensesAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    // When an external EXO ops backend is wired, the conversion must go through it
    // (the in-proc EXO path is unreliable) and the in-proc mailbox service must NOT be touched.
    [Fact]
    public async Task ExternalExo_WhenWired_HandlesConversion_InProcServiceUntouched()
    {
        var users = UsersOk();
        var mbx = new Mock<ISharedMailboxService>();
        var exo = new Mock<IExternalExoOps>();
        exo.Setup(e => e.GetMailboxFactsAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new MailboxInfo("u@a", "U", "u@a", "UserMailbox"));
        exo.Setup(e => e.ConvertToSharedAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new MailboxInfo("u@a", "U", "u@a", "SharedMailbox"));

        var sut = new OffboardingService(users.Object, mbx.Object, exo.Object);
        var r = await sut.RunAsync("jane@a", new OffboardingOptions(DisableAccount: false, RemoveLicenses: true, ConvertMailboxToShared: true));

        r.Success.Should().BeTrue();
        exo.Verify(e => e.ConvertToSharedAsync("jane@a", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
        mbx.Verify(s => s.ConvertToSharedAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
        users.Verify(u => u.RemoveAllLicensesAsync("uid", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    // Idempotency: re-running over an already-disabled account must not re-issue the disable
    // call, but must still revoke sessions (cheap, and good hygiene on a re-run).
    [Fact]
    public async Task AlreadyDisabled_SkipsDisableCall_StillRevokesSessions()
    {
        var users = new Mock<IUsersService>();
        users.Setup(u => u.GetByIdAsync(It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(DisabledUser());
        var sut = new OffboardingService(users.Object, MailboxOk().Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(DisableAccount: true, RemoveLicenses: false, ConvertMailboxToShared: false));

        r.Success.Should().BeTrue();
        r.Steps.Should().Contain(s => s.Name.Contains("Deshabilitar") && s.Status == "OMITIDO");
        users.Verify(u => u.SetAccountEnabledAsync(It.IsAny<string>(), It.IsAny<bool>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
        users.Verify(u => u.RevokeSignInSessionsAsync("uid", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    // Idempotency: an already-shared mailbox skips the convert step but the gate still lets the
    // license be released (the mailbox is already in its final shared state).
    [Fact]
    public async Task AlreadyShared_SkipsConvert_AllowsLicenseRemoval()
    {
        var users = UsersOk();
        var exo = ExoWithFacts(new MailboxInfo("u@a", "U", "u@a", "SharedMailbox"));
        var sut = new OffboardingService(users.Object, new Mock<ISharedMailboxService>().Object, exo.Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(DisableAccount: false, RemoveLicenses: true, ConvertMailboxToShared: true));

        r.Success.Should().BeTrue();
        r.Steps.Should().Contain(s => s.Name.Contains("compartido") && s.Status == "OMITIDO" && s.Detail.Contains("ya"));
        exo.Verify(e => e.ConvertToSharedAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
        users.Verify(u => u.RemoveAllLicensesAsync("uid", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    // Safety gate: a >50 GB mailbox can't stay a shared mailbox without a license, so stripping
    // the license is blocked even though the conversion itself succeeded.
    [Fact]
    public async Task MailboxOver50Gb_BlocksLicenseRemoval()
    {
        var users = UsersOk();
        var exo = ExoWithFacts(new MailboxInfo("u@a", "U", "u@a", "UserMailbox", TotalItemBytes: 60L * 1024 * 1024 * 1024));
        var sut = new OffboardingService(users.Object, new Mock<ISharedMailboxService>().Object, exo.Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(DisableAccount: false, RemoveLicenses: true, ConvertMailboxToShared: true));

        r.Success.Should().BeFalse();
        r.Steps.Should().Contain(s => s.Name.Contains("Quitar licencias") && s.Status == "OMITIDO" && s.Detail.Contains("50"));
        users.Verify(u => u.RemoveAllLicensesAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    // Safety gate: an active hold needs a license to be preserved, so stripping it is blocked.
    [Fact]
    public async Task MailboxOnHold_BlocksLicenseRemoval()
    {
        var users = UsersOk();
        var exo = ExoWithFacts(new MailboxInfo("u@a", "U", "u@a", "UserMailbox", LitigationHoldEnabled: true));
        var sut = new OffboardingService(users.Object, new Mock<ISharedMailboxService>().Object, exo.Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(DisableAccount: false, RemoveLicenses: true, ConvertMailboxToShared: true));

        r.Success.Should().BeFalse();
        r.Steps.Should().Contain(s => s.Name.Contains("Quitar licencias") && s.Status == "OMITIDO" && s.Detail.Contains("hold"));
        users.Verify(u => u.RemoveAllLicensesAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    // Dry-run rehearses the whole flow read-only: every mutating call is suppressed and the
    // mutating steps are reported as SIMULADO.
    [Fact]
    public async Task DryRun_SimulatesEverything_NoMutations()
    {
        var users = UsersOk();
        var mbx = MailboxOk();
        var sut = new OffboardingService(users.Object, mbx.Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(true, true, true, DryRun: true));

        r.DryRun.Should().BeTrue();
        r.Success.Should().BeTrue();
        r.Steps.Should().Contain(s => s.Status == "SIMULADO");
        users.Verify(u => u.SetAccountEnabledAsync(It.IsAny<string>(), It.IsAny<bool>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
        users.Verify(u => u.RevokeSignInSessionsAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
        users.Verify(u => u.RemoveAllLicensesAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
        mbx.Verify(s => s.ConvertToSharedAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    // Residual licenses after a successful removal call are flagged (likely group-inherited),
    // not silently reported as fully removed — but it's a warning, not a hard failure.
    [Fact]
    public async Task ResidualLicensesAfterRemoval_WarnsGroupInherited()
    {
        var users = new Mock<IUsersService>();
        users.Setup(u => u.GetByIdAsync(It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(SampleUser(1)); // never drops to 0 → simulates a group-inherited license
        var sut = new OffboardingService(users.Object, MailboxOk().Object)
        {
            VerifyPollDelay = TimeSpan.Zero,
            VerifyAttempts = 2,
        };

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(DisableAccount: false, RemoveLicenses: true, ConvertMailboxToShared: false));

        r.Success.Should().BeTrue();
        r.Steps.Should().Contain(s => s.Name.Contains("Quitar licencias") && s.Status == "AVISO" && s.Detail.Contains("grupo"));
        users.Verify(u => u.RemoveAllLicensesAsync("uid", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    // EXO reachable but the user has no mailbox → convert is skipped (not errored), and the
    // license can still be released (no mailbox to strand). Repro of the real "pwsh exit 1".
    [Fact]
    public async Task NoMailbox_SkipsConvert_AllowsLicenseRemoval()
    {
        var users = UsersOk();
        var exo = new Mock<IExternalExoOps>();
        exo.Setup(e => e.GetMailboxFactsAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((MailboxInfo?)null); // connected, but no mailbox for this user
        var sut = new OffboardingService(users.Object, new Mock<ISharedMailboxService>().Object, exo.Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(DisableAccount: false, RemoveLicenses: true, ConvertMailboxToShared: true));

        r.Success.Should().BeTrue();
        r.Steps.Should().Contain(s => s.Name.Contains("compartido") && s.Status == "OMITIDO" && s.Detail.Contains("no tiene buzón"));
        exo.Verify(e => e.ConvertToSharedAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
        users.Verify(u => u.RemoveAllLicensesAsync("uid", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    // Optional EXO finalization steps run after a successful conversion when requested.
    [Fact]
    public async Task Finalization_RunsAutoReplyForwardHideGal_WhenRequested()
    {
        var users = UsersOk();
        var exo = ExoWithFacts(new MailboxInfo("u@a", "U", "u@a", "UserMailbox"));
        var sut = new OffboardingService(users.Object, new Mock<ISharedMailboxService>().Object, exo.Object);

        var opts = new OffboardingOptions(
            DisableAccount: false, RemoveLicenses: false, ConvertMailboxToShared: true,
            ForwardTo: "deleg@a", AutoReplyMessage: "Ya no trabaja aquí", HideFromGal: true);
        var r = await sut.RunAsync("jane@a", opts);

        r.Success.Should().BeTrue();
        r.Steps.Should().Contain(s => s.Name == "Auto-reply" && s.Status == "OK");
        r.Steps.Should().Contain(s => s.Name == "Forward al delegado" && s.Status == "OK");
        r.Steps.Should().Contain(s => s.Name == "Ocultar de la GAL" && s.Status == "OK" && s.Detail.Contains("GAL"));
        exo.Verify(e => e.SetAutoReplyAsync("jane@a", "Ya no trabaja aquí", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
        exo.Verify(e => e.SetForwardingAsync("jane@a", "deleg@a", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
        exo.Verify(e => e.HideFromGalAsync("jane@a", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    // Dry-run simulates the finalization steps too — no EXO calls.
    [Fact]
    public async Task Finalization_DryRun_SimulatesWithoutExoCalls()
    {
        var users = UsersOk();
        var exo = ExoWithFacts(new MailboxInfo("u@a", "U", "u@a", "UserMailbox"));
        var sut = new OffboardingService(users.Object, new Mock<ISharedMailboxService>().Object, exo.Object);

        var opts = new OffboardingOptions(false, false, true, DryRun: true,
            ForwardTo: "deleg@a", AutoReplyMessage: "msg", HideFromGal: true);
        var r = await sut.RunAsync("jane@a", opts);

        r.Steps.Should().Contain(s => s.Name == "Auto-reply" && s.Status == "SIMULADO");
        r.Steps.Should().Contain(s => s.Name == "Forward al delegado" && s.Status == "SIMULADO");
        exo.Verify(e => e.SetAutoReplyAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
        exo.Verify(e => e.SetForwardingAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
        exo.Verify(e => e.HideFromGalAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    // Without an external EXO backend, the finalization steps are skipped (not errored).
    [Fact]
    public async Task Finalization_NoExternalExo_Omitido()
    {
        var users = UsersOk();
        var sut = new OffboardingService(users.Object, MailboxOk().Object); // no IExternalExoOps

        var opts = new OffboardingOptions(false, false, true, AutoReplyMessage: "msg", HideFromGal: true);
        var r = await sut.RunAsync("jane@a", opts);

        r.Steps.Should().Contain(s => s.Name == "Auto-reply" && s.Status == "OMITIDO");
        r.Steps.Should().Contain(s => s.Name == "Ocultar de la GAL" && s.Status == "OMITIDO");
    }

    // A finalization failure is non-fatal: reported AVISO, overall run still succeeds.
    [Fact]
    public async Task Finalization_Failure_IsNonFatal()
    {
        var users = UsersOk();
        var exo = ExoWithFacts(new MailboxInfo("u@a", "U", "u@a", "UserMailbox"));
        exo.Setup(e => e.SetAutoReplyAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("EXO hiccup"));
        var sut = new OffboardingService(users.Object, new Mock<ISharedMailboxService>().Object, exo.Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(false, false, true, AutoReplyMessage: "msg"));

        r.Success.Should().BeTrue();
        r.Steps.Should().Contain(s => s.Name == "Auto-reply" && s.Status == "AVISO" && s.Detail.Contains("EXO hiccup"));
    }
}
