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

    // Delegation (FullAccess + SendAs) goes through the external EXO ops — never the in-proc
    // ISharedMailboxService (which trips the GetResponseHeader bug).
    [Fact]
    public async Task Delegate_GrantsFullAccessAndSendAs_ViaExternalExo()
    {
        var users = UsersOk();
        var exo = ExoWithFacts(new MailboxInfo("u@a", "U", "u@a", "UserMailbox"));
        exo.Setup(e => e.GrantDelegateAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<bool>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync("FullAccess + SendAs -> deleg@a");
        var inProc = new Mock<ISharedMailboxService>(MockBehavior.Strict); // must NOT be touched
        var sut = new OffboardingService(users.Object, inProc.Object, exo.Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(false, false, true, DelegateMailboxTo: "deleg@a"));

        r.Success.Should().BeTrue();
        r.Steps.Should().Contain(s => s.Name.Contains("Delegar") && s.Status == "OK");
        exo.Verify(e => e.GrantDelegateAsync("jane@a", "deleg@a", true, It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    private static Mock<IUsersService> UsersWithGroups(params GroupSummary[] groups)
    {
        var m = UsersOk();
        m.Setup(u => u.GetGroupMembershipsAsync("uid", It.IsAny<CancellationToken>()))
            .ReturnsAsync(groups.ToList());
        return m;
    }

    // Offboarding removes the leaver from every group / DL (best-effort, via IGroupsService).
    [Fact]
    public async Task RemovesFromGroups_WhenRequested()
    {
        var users = UsersWithGroups(new("g1", "Grupo 1", null, "Unified"), new("g2", "DL 2", null, "Distribution"));
        var groups = new Mock<IGroupsService>();
        var sut = new OffboardingService(users.Object, MailboxOk().Object, null, groups.Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(false, false, false, RemoveFromGroups: true));

        r.Success.Should().BeTrue();
        r.Steps.Should().Contain(s => s.Name.Contains("grupos") && s.Status == "OK" && s.Detail.Contains("2"));
        groups.Verify(g => g.RemoveMemberAsync("g1", "uid", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
        groups.Verify(g => g.RemoveMemberAsync("g2", "uid", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    // Groups that can't be removed (dynamic / on-prem / classic DL) are reported, not fatal.
    [Fact]
    public async Task RemoveFromGroups_PartialFailure_ReportsAvisoButSucceeds()
    {
        var users = UsersWithGroups(new("g1", "Dinámico", null, "Unified"), new("g2", "OK", null, "Unified"));
        var groups = new Mock<IGroupsService>();
        groups.Setup(g => g.RemoveMemberAsync("g1", "uid", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("dynamic group"));
        var sut = new OffboardingService(users.Object, MailboxOk().Object, null, groups.Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(false, false, false, RemoveFromGroups: true));

        r.Success.Should().BeTrue();
        r.Steps.Should().Contain(s => s.Name.Contains("grupos") && s.Status == "AVISO" && s.Detail.Contains("Dinámico"));
        groups.Verify(g => g.RemoveMemberAsync("g2", "uid", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task RemoveFromGroups_NoGroupsService_Omitido()
    {
        var users = UsersOk();
        var sut = new OffboardingService(users.Object, MailboxOk().Object); // no IGroupsService

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(false, false, false, RemoveFromGroups: true));

        r.Steps.Should().Contain(s => s.Name.Contains("grupos") && s.Status == "OMITIDO");
    }

    // Classic DLs / mail-enabled security groups route through external EXO (Graph can't write
    // their membership). M365 groups stay on Graph. A DL without SMTP falls back to its id.
    [Fact]
    public async Task ClassicDls_RemovedViaExternalExo_NotGraph()
    {
        var users = UsersWithGroups(
            new("g1", "Equipo", null, "M365"),
            new("g2", "DL Ventas", "ventas@a", "DistributionList"),
            new("g3", "Seguridad Mail", null, "MailSecurity"));
        var groups = new Mock<IGroupsService>();
        var exo = ExoWithFacts(new MailboxInfo("u@a", "U", "u@a", "UserMailbox"));
        IReadOnlyList<string>? sentIdentities = null;
        exo.Setup(e => e.RemoveFromDistributionGroupsAsync("jane@a", It.IsAny<IReadOnlyList<string>>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .Callback<string, IReadOnlyList<string>, IProgress<LogEntry>?, CancellationToken>((_, ids, _, _) => sentIdentities = ids)
            .ReturnsAsync(new[]
            {
                new DistributionGroupRemovalResult("ventas@a", true, "quitado"),
                new DistributionGroupRemovalResult("g3", true, "quitado"),
            });
        var sut = new OffboardingService(users.Object, MailboxOk().Object, exo.Object, groups.Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(false, false, false, RemoveFromGroups: true));

        r.Success.Should().BeTrue();
        r.Steps.Should().Contain(s => s.Name.Contains("grupos") && s.Status == "OK" && s.Detail.Contains("3") && s.Detail.Contains("EXO"));
        sentIdentities.Should().BeEquivalentTo(new[] { "ventas@a", "g3" }); // SMTP preferred, id fallback
        groups.Verify(g => g.RemoveMemberAsync("g1", "uid", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
        groups.Verify(g => g.RemoveMemberAsync("g2", It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
        groups.Verify(g => g.RemoveMemberAsync("g3", It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    // A per-DL EXO failure is reported AVISO with the group's display name, not fatal.
    [Fact]
    public async Task ClassicDl_ExoFailure_ReportsAviso()
    {
        var users = UsersWithGroups(new GroupSummary("g2", "DL Ventas", "ventas@a", "DistributionList"));
        var exo = ExoWithFacts(new MailboxInfo("u@a", "U", "u@a", "UserMailbox"));
        exo.Setup(e => e.RemoveFromDistributionGroupsAsync(It.IsAny<string>(), It.IsAny<IReadOnlyList<string>>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new[] { new DistributionGroupRemovalResult("ventas@a", false, "gestionado por el propietario") });
        var sut = new OffboardingService(users.Object, MailboxOk().Object, exo.Object, new Mock<IGroupsService>().Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(false, false, false, RemoveFromGroups: true));

        r.Success.Should().BeTrue();
        r.Steps.Should().Contain(s => s.Name.Contains("grupos") && s.Status == "AVISO" && s.Detail.Contains("DL Ventas"));
    }

    // Without external EXO wired, DL kinds keep the old behavior: attempted via Graph.
    [Fact]
    public async Task ClassicDl_NoExternalExo_FallsBackToGraph()
    {
        var users = UsersWithGroups(new GroupSummary("g2", "DL Ventas", "ventas@a", "DistributionList"));
        var groups = new Mock<IGroupsService>();
        var sut = new OffboardingService(users.Object, MailboxOk().Object, null, groups.Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(false, false, false, RemoveFromGroups: true));

        r.Steps.Should().Contain(s => s.Name.Contains("grupos") && s.Status == "OK");
        groups.Verify(g => g.RemoveMemberAsync("g2", "uid", It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Theory]
    [InlineData("DistributionList", true)]
    [InlineData("MailSecurity", true)]
    [InlineData("mailsecurity", true)]
    [InlineData("M365", false)]
    [InlineData("Security", false)]
    [InlineData("Other", false)]
    [InlineData(null, false)]
    public void IsExoManagedGroup_RoutesByKind(string? kind, bool expected) =>
        OffboardingService.IsExoManagedGroup(kind).Should().Be(expected);

    // ---- directory (admin) roles ----

    private static void WithRoles(Mock<IUsersService> users, params DirectoryRoleSummary[] roles) =>
        users.Setup(u => u.GetDirectoryRolesAsync("uid", It.IsAny<CancellationToken>()))
            .ReturnsAsync(roles);

    [Fact]
    public async Task RemoveDirectoryRoles_RemovesEach_ReportsNames()
    {
        var users = UsersOk();
        WithRoles(users, new DirectoryRoleSummary("r1", "Exchange Administrator"), new DirectoryRoleSummary("r2", "User Administrator"));
        var sut = new OffboardingService(users.Object, MailboxOk().Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(false, false, false, RemoveDirectoryRoles: true));

        r.Success.Should().BeTrue();
        r.Steps.Should().Contain(s => s.Name.Contains("roles") && s.Status == "OK" && s.Detail.Contains("Exchange Administrator"));
        users.Verify(u => u.RemoveFromDirectoryRoleAsync("r1", "uid", null, It.IsAny<CancellationToken>()), Times.Once);
        users.Verify(u => u.RemoveFromDirectoryRoleAsync("r2", "uid", null, It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task RemoveDirectoryRoles_NoRoles_ReportsOkWithoutCalls()
    {
        var users = UsersOk();
        WithRoles(users); // empty
        var sut = new OffboardingService(users.Object, MailboxOk().Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(false, false, false, RemoveDirectoryRoles: true));

        r.Steps.Should().Contain(s => s.Name.Contains("roles") && s.Status == "OK" && s.Detail.Contains("sin roles"));
        users.Verify(u => u.RemoveFromDirectoryRoleAsync(It.IsAny<string>(), It.IsAny<string>(), null, It.IsAny<CancellationToken>()), Times.Never);
    }

    // A missing RoleManagement.ReadWrite.Directory scope surfaces as AVISO with the exact
    // remediation, never as a fatal error.
    [Fact]
    public async Task RemoveDirectoryRoles_Forbidden_ReportsAvisoWithScopeHint()
    {
        var users = UsersOk();
        WithRoles(users, new DirectoryRoleSummary("r1", "Global Administrator"));
        users.Setup(u => u.RemoveFromDirectoryRoleAsync("r1", "uid", null, It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("Insufficient privileges to complete the operation."));
        var sut = new OffboardingService(users.Object, MailboxOk().Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(false, false, false, RemoveDirectoryRoles: true));

        r.Success.Should().BeTrue();
        r.Steps.Should().Contain(s => s.Name.Contains("roles") && s.Status == "AVISO"
            && s.Detail.Contains("Global Administrator") && s.Detail.Contains("RoleManagement.ReadWrite.Directory"));
    }

    [Fact]
    public async Task RemoveDirectoryRoles_DryRun_SimulatesWithNames()
    {
        var users = UsersOk();
        WithRoles(users, new DirectoryRoleSummary("r1", "Helpdesk Administrator"));
        var sut = new OffboardingService(users.Object, MailboxOk().Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(false, false, false, DryRun: true, RemoveDirectoryRoles: true));

        r.Steps.Should().Contain(s => s.Name.Contains("roles") && s.Status == "SIMULADO" && s.Detail.Contains("Helpdesk Administrator"));
        users.Verify(u => u.RemoveFromDirectoryRoleAsync(It.IsAny<string>(), It.IsAny<string>(), null, It.IsAny<CancellationToken>()), Times.Never);
    }

    // ---- result notification (sendMail) ----

    [Fact]
    public async Task NotifyResult_SendsSummaryFromLeaver_AsLastStep()
    {
        var users = UsersOk();
        string? sentBody = null; string? sentSubject = null;
        users.Setup(u => u.SendMailAsync("jane@a", "hr@a", It.IsAny<string>(), It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .Callback<string, string, string, string, IProgress<LogEntry>?, CancellationToken>((_, _, s, b, _, _) => { sentSubject = s; sentBody = b; })
            .Returns(Task.CompletedTask);
        var sut = new OffboardingService(users.Object, MailboxOk().Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(true, false, false, NotifyResultTo: "hr@a"));

        r.Success.Should().BeTrue();
        r.Steps.Last().Name.Should().Be("Notificar resultado");
        r.Steps.Last().Status.Should().Be("OK");
        sentSubject.Should().Contain("completado");
        sentBody.Should().Contain("Deshabilitar cuenta"); // the summary covers the earlier steps
    }

    // A failed send (e.g. missing Mail.Send) is AVISO with the scope remediation, never fatal.
    [Fact]
    public async Task NotifyResult_SendFails_AvisoWithScopeHint()
    {
        var users = UsersOk();
        users.Setup(u => u.SendMailAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("Access is denied: Forbidden"));
        var sut = new OffboardingService(users.Object, MailboxOk().Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(true, false, false, NotifyResultTo: "hr@a"));

        r.Success.Should().BeTrue();
        r.Steps.Should().Contain(s => s.Name == "Notificar resultado" && s.Status == "AVISO" && s.Detail.Contains("Mail.Send"));
    }

    [Fact]
    public async Task NotifyResult_DryRun_DoesNotSend()
    {
        var users = UsersOk();
        var sut = new OffboardingService(users.Object, MailboxOk().Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(true, false, false, DryRun: true, NotifyResultTo: "hr@a"));

        r.Steps.Should().Contain(s => s.Name == "Notificar resultado" && s.Status == "SIMULADO");
        users.Verify(u => u.SendMailAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    // ---- MFA auth methods ----

    private static void WithAuthMethods(Mock<IUsersService> users, params AuthMethodSummary[] methods) =>
        users.Setup(u => u.GetAuthMethodsAsync("uid", It.IsAny<CancellationToken>()))
            .ReturnsAsync(methods);

    [Fact]
    public async Task RemoveAuthMethods_DeletesRemovableOnly_PasswordSkipped()
    {
        var users = UsersOk();
        var phone = new AuthMethodSummary("m1", "Phone", "+34...", true);
        var pwd = new AuthMethodSummary("m2", "Password", null, false);
        var fido = new AuthMethodSummary("m3", "Fido2", "YubiKey", true);
        WithAuthMethods(users, phone, pwd, fido);
        var sut = new OffboardingService(users.Object, MailboxOk().Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(false, false, false, RemoveAuthMethods: true));

        r.Success.Should().BeTrue();
        r.Steps.Should().Contain(s => s.Name.Contains("MFA") && s.Status == "OK" && s.Detail.Contains("Phone") && s.Detail.Contains("Fido2"));
        users.Verify(u => u.RemoveAuthMethodAsync("uid", phone, null, It.IsAny<CancellationToken>()), Times.Once);
        users.Verify(u => u.RemoveAuthMethodAsync("uid", fido, null, It.IsAny<CancellationToken>()), Times.Once);
        users.Verify(u => u.RemoveAuthMethodAsync("uid", pwd, null, It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task RemoveAuthMethods_OnlyPassword_ReportsOkNoCalls()
    {
        var users = UsersOk();
        WithAuthMethods(users, new AuthMethodSummary("m1", "Password", null, false));
        var sut = new OffboardingService(users.Object, MailboxOk().Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(false, false, false, RemoveAuthMethods: true));

        r.Steps.Should().Contain(s => s.Name.Contains("MFA") && s.Status == "OK" && s.Detail.Contains("sin métodos"));
        users.Verify(u => u.RemoveAuthMethodAsync(It.IsAny<string>(), It.IsAny<AuthMethodSummary>(), null, It.IsAny<CancellationToken>()), Times.Never);
    }

    // Missing UserAuthenticationMethod.ReadWrite.All on the READ surfaces the scope hint too.
    [Fact]
    public async Task RemoveAuthMethods_ReadForbidden_AvisoWithScopeHint()
    {
        var users = UsersOk();
        users.Setup(u => u.GetAuthMethodsAsync("uid", It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("Access is denied. Insufficient privileges."));
        var sut = new OffboardingService(users.Object, MailboxOk().Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(false, false, false, RemoveAuthMethods: true));

        r.Success.Should().BeTrue();
        r.Steps.Should().Contain(s => s.Name.Contains("MFA") && s.Status == "AVISO" && s.Detail.Contains("UserAuthenticationMethod.ReadWrite.All"));
    }

    [Fact]
    public async Task RemoveAuthMethods_DryRun_SimulatesWithKinds()
    {
        var users = UsersOk();
        WithAuthMethods(users, new AuthMethodSummary("m1", "MicrosoftAuthenticator", "iPhone", true));
        var sut = new OffboardingService(users.Object, MailboxOk().Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(false, false, false, DryRun: true, RemoveAuthMethods: true));

        r.Steps.Should().Contain(s => s.Name.Contains("MFA") && s.Status == "SIMULADO" && s.Detail.Contains("MicrosoftAuthenticator"));
        users.Verify(u => u.RemoveAuthMethodAsync(It.IsAny<string>(), It.IsAny<AuthMethodSummary>(), null, It.IsAny<CancellationToken>()), Times.Never);
    }

    // ---- litigation hold (inactive-mailbox retention path) ----

    // Hold + license removal WITHOUT converting to shared is the supported inactive-mailbox
    // path: the hold is applied while the mailbox is still licensed, then the license goes.
    [Fact]
    public async Task LitigationHold_Enabled_AppliedBeforeLicenseRemoval_WithDuration()
    {
        var users = UsersOk();
        var exo = ExoWithFacts(new MailboxInfo("u@a", "U", "u@a", "UserMailbox"));
        var order = new List<string>();
        exo.Setup(e => e.SetLitigationHoldAsync("jane@a", 2555, It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .Callback(() => order.Add("hold"))
            .ReturnsAsync("LitigationHoldEnabled=true (2555 dias)");
        users.Setup(u => u.RemoveAllLicensesAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .Callback(() => order.Add("licenses"))
            .Returns(Task.CompletedTask);
        var sut = new OffboardingService(users.Object, MailboxOk().Object, exo.Object) { VerifyPollDelay = TimeSpan.Zero };

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(
            DisableAccount: false, RemoveLicenses: true, ConvertMailboxToShared: false,
            EnableLitigationHold: true, LitigationHoldDays: 2555));

        r.Success.Should().BeTrue();
        r.Steps.Should().Contain(s => s.Name == "Litigation Hold" && s.Status == "OK");
        r.Steps.Should().Contain(s => s.Name.Contains("Quitar licencias") && s.Status == "OK");
        order.Should().Equal("hold", "licenses");
    }

    // Idempotent: an already-on-hold mailbox reports OMITIDO and still counts as retained,
    // so license removal proceeds (inactive-mailbox path).
    [Fact]
    public async Task LitigationHold_AlreadyOn_SkipsAndStillRemovesLicenses()
    {
        var users = UsersOk();
        var exo = ExoWithFacts(new MailboxInfo("u@a", "U", "u@a", "UserMailbox", LitigationHoldEnabled: true));
        var sut = new OffboardingService(users.Object, MailboxOk().Object, exo.Object) { VerifyPollDelay = TimeSpan.Zero };

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(
            DisableAccount: false, RemoveLicenses: true, ConvertMailboxToShared: false,
            EnableLitigationHold: true));

        r.Success.Should().BeTrue();
        r.Steps.Should().Contain(s => s.Name == "Litigation Hold" && s.Status == "OMITIDO" && s.Detail.Contains("ya tiene"));
        exo.Verify(e => e.SetLitigationHoldAsync(It.IsAny<string>(), It.IsAny<int?>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
        users.Verify(u => u.RemoveAllLicensesAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    // Safety gate: the operator asked for retention and the hold failed — stripping the
    // license would start the 30-day deletion clock on data they meant to keep.
    [Fact]
    public async Task LitigationHold_Fails_BlocksLicenseRemoval()
    {
        var users = UsersOk();
        var exo = ExoWithFacts(new MailboxInfo("u@a", "U", "u@a", "UserMailbox"));
        exo.Setup(e => e.SetLitigationHoldAsync(It.IsAny<string>(), It.IsAny<int?>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("the mailbox license does not permit holds"));
        var sut = new OffboardingService(users.Object, MailboxOk().Object, exo.Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(
            DisableAccount: false, RemoveLicenses: true, ConvertMailboxToShared: false,
            EnableLitigationHold: true));

        r.Success.Should().BeFalse();
        r.Steps.Should().Contain(s => s.Name == "Litigation Hold" && s.Status == "ERROR" && s.Detail.Contains("Plan 2"));
        r.Steps.Should().Contain(s => s.Name.Contains("Quitar licencias") && s.Status == "OMITIDO" && s.Detail.Contains("hold"));
        users.Verify(u => u.RemoveAllLicensesAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    // Hold enabled this run + convert to shared: a shared mailbox with a hold still needs a
    // license, so the existing shared+hold gate must also see the just-enabled hold.
    [Fact]
    public async Task LitigationHold_JustEnabled_PlusConvertShared_BlocksLicenseRemoval()
    {
        var users = UsersOk();
        var exo = ExoWithFacts(new MailboxInfo("u@a", "U", "u@a", "UserMailbox"));
        exo.Setup(e => e.SetLitigationHoldAsync(It.IsAny<string>(), It.IsAny<int?>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync("LitigationHoldEnabled=true (indefinido)");
        var sut = new OffboardingService(users.Object, MailboxOk().Object, exo.Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(
            DisableAccount: false, RemoveLicenses: true, ConvertMailboxToShared: true,
            EnableLitigationHold: true));

        r.Success.Should().BeFalse();
        r.Steps.Should().Contain(s => s.Name == "Litigation Hold" && s.Status == "OK");
        r.Steps.Should().Contain(s => s.Name.Contains("Quitar licencias") && s.Status == "OMITIDO" && s.Detail.Contains("hold"));
        users.Verify(u => u.RemoveAllLicensesAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    // Dry-run simulates the hold (no EXO mutation) and the license step stays simulated too.
    [Fact]
    public async Task LitigationHold_DryRun_Simulated()
    {
        var users = UsersOk();
        var exo = ExoWithFacts(new MailboxInfo("u@a", "U", "u@a", "UserMailbox"));
        var sut = new OffboardingService(users.Object, MailboxOk().Object, exo.Object);

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(
            DisableAccount: false, RemoveLicenses: true, ConvertMailboxToShared: false,
            DryRun: true, EnableLitigationHold: true, LitigationHoldDays: 365));

        r.Success.Should().BeTrue();
        r.Steps.Should().Contain(s => s.Name == "Litigation Hold" && s.Status == "SIMULADO" && s.Detail.Contains("365"));
        exo.Verify(e => e.SetLitigationHoldAsync(It.IsAny<string>(), It.IsAny<int?>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
        users.Verify(u => u.RemoveAllLicensesAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    // No mailbox in EXO → the hold step is skipped, and the requested-but-failed gate must
    // NOT fire (there is no data to retain); licenses are removed normally.
    [Fact]
    public async Task LitigationHold_NoMailbox_SkippedAndLicensesStillRemoved()
    {
        var users = UsersOk();
        var exo = new Mock<IExternalExoOps>();
        exo.Setup(x => x.GetMailboxFactsAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((MailboxInfo?)null);
        var sut = new OffboardingService(users.Object, MailboxOk().Object, exo.Object) { VerifyPollDelay = TimeSpan.Zero };

        var r = await sut.RunAsync("jane@a", new OffboardingOptions(
            DisableAccount: false, RemoveLicenses: true, ConvertMailboxToShared: false,
            EnableLitigationHold: true));

        r.Success.Should().BeTrue();
        r.Steps.Should().Contain(s => s.Name == "Litigation Hold" && s.Status == "OMITIDO" && s.Detail.Contains("no tiene buzón"));
        users.Verify(u => u.RemoveAllLicensesAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }
}
