using FluentAssertions;
using Grex365.App.ViewModels;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Moq;

namespace Grex365.App.Tests;

public class MailboxRulesViewModelTests
{
    private sealed class Harness
    {
        public Mock<IMailboxRulesService> Rules { get; } = new();
        public Mock<IRbacGuard> Rbac { get; } = new();
        public TestUiLogSink Log { get; } = new();
        public TestDialogService Dialogs { get; } = new();
        public MailboxRulesViewModel Vm { get; }

        public Harness(bool rbacAllowed = true)
        {
            Rbac.Setup(r => r.EvaluateAsync(It.IsAny<CancellationToken>()))
                .ReturnsAsync(new RbacDecision(rbacAllowed, rbacAllowed ? "OK" : "Not in group"));
            Vm = new MailboxRulesViewModel(Rules.Object, Log, Rbac.Object, Dialogs);
        }
    }

    [Fact]
    public async Task ApplyForwarding_EmptyIdentity_SetsStatus()
    {
        var h = new Harness();
        h.Vm.Identity = string.Empty;

        await h.Vm.ApplyForwardingCommand.ExecuteAsync(null);

        h.Vm.StatusMessage.Should().Be("Buzón vacío.");
        h.Dialogs.Confirmations.Should().BeEmpty();
    }

    [Fact]
    public async Task ApplyForwarding_RbacDenied_Blocks()
    {
        var h = new Harness(rbacAllowed: false);
        h.Vm.Identity = "u@a";
        h.Vm.ForwardingSmtp = "ext@b.com";

        await h.Vm.ApplyForwardingCommand.ExecuteAsync(null);

        h.Vm.StatusMessage.Should().Be("Not in group");
        h.Dialogs.Confirmations.Should().BeEmpty();
    }

    [Fact]
    public async Task ApplyForwarding_ConfirmNo_DoesNothing()
    {
        var h = new Harness();
        h.Vm.Identity = "u@a";
        h.Vm.ForwardingSmtp = "ext@b.com";
        h.Dialogs.ConfirmResult = false;

        await h.Vm.ApplyForwardingCommand.ExecuteAsync(null);

        h.Vm.StatusMessage.Should().Be("Cancelado por el usuario.");
        h.Rules.Verify(r => r.SetForwardingAsync(It.IsAny<string>(), It.IsAny<string>(),
            It.IsAny<bool>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task ApplyForwarding_ConfirmYes_CallsService()
    {
        var h = new Harness();
        h.Vm.Identity = "  u@a  ";
        h.Vm.ForwardingSmtp = "  ext@b.com  ";
        h.Vm.DeliverToMailboxAndForward = true;
        h.Dialogs.ConfirmResult = true;
        h.Rules.Setup(r => r.SetForwardingAsync("u@a", "ext@b.com", true,
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>())).Returns(Task.CompletedTask);

        await h.Vm.ApplyForwardingCommand.ExecuteAsync(null);

        h.Rules.Verify(r => r.SetForwardingAsync("u@a", "ext@b.com", true,
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
        h.Vm.StatusMessage.Should().Be("Reenvío aplicado.");
        h.Vm.CurrentForwardingDisplay.Should().Contain("ext@b.com");
    }

    [Fact]
    public async Task ClearForwarding_ConfirmYes_CallsService_AndResetsFields()
    {
        var h = new Harness();
        h.Vm.Identity = "u@a";
        h.Vm.ForwardingSmtp = "ext@b.com";
        h.Vm.DeliverToMailboxAndForward = true;
        h.Dialogs.ConfirmResult = true;
        h.Rules.Setup(r => r.ClearForwardingAsync("u@a",
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>())).Returns(Task.CompletedTask);

        await h.Vm.ClearForwardingCommand.ExecuteAsync(null);

        h.Rules.Verify(r => r.ClearForwardingAsync("u@a",
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
        h.Vm.ForwardingSmtp.Should().BeEmpty();
        h.Vm.DeliverToMailboxAndForward.Should().BeFalse();
        h.Vm.CurrentForwardingDisplay.Should().Be("(sin configurar)");
        h.Vm.StatusMessage.Should().Be("Reenvío eliminado.");
        h.Dialogs.Confirmations[0].Icon.Should().Be(DialogIcon.Warning);
    }

    [Fact]
    public async Task ClearForwarding_ConfirmNo_DoesNothing()
    {
        var h = new Harness();
        h.Vm.Identity = "u@a";
        h.Dialogs.ConfirmResult = false;

        await h.Vm.ClearForwardingCommand.ExecuteAsync(null);

        h.Rules.Verify(r => r.ClearForwardingAsync(It.IsAny<string>(),
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task RemoveCalendarPermission_NoSelection_SetsStatus()
    {
        var h = new Harness();
        h.Vm.Identity = "u@a";
        h.Vm.SelectedCalendarPermission = null;

        await h.Vm.RemoveCalendarPermissionCommand.ExecuteAsync(null);

        h.Vm.StatusMessage.Should().Be("Selecciona un permiso.");
        h.Dialogs.Confirmations.Should().BeEmpty();
    }

    [Fact]
    public async Task RemoveCalendarPermission_ConfirmYes_CallsService_AndRemovesFromCollection()
    {
        var h = new Harness();
        var perm = new CalendarPermissionEntry("jane@a", "Reviewer");
        h.Vm.Identity = "u@a";
        h.Vm.CalendarPermissions.Add(perm);
        h.Vm.SelectedCalendarPermission = perm;
        h.Dialogs.ConfirmResult = true;
        h.Rules.Setup(r => r.RemoveCalendarPermissionAsync("u@a", "jane@a",
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>())).Returns(Task.CompletedTask);

        await h.Vm.RemoveCalendarPermissionCommand.ExecuteAsync(null);

        h.Rules.Verify(r => r.RemoveCalendarPermissionAsync("u@a", "jane@a",
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
        h.Vm.CalendarPermissions.Should().NotContain(perm);
        h.Vm.StatusMessage.Should().Be("Permiso eliminado.");
    }

    [Fact]
    public async Task RemoveCalendarPermission_ConfirmNo_DoesNothing()
    {
        var h = new Harness();
        var perm = new CalendarPermissionEntry("jane@a", "Reviewer");
        h.Vm.Identity = "u@a";
        h.Vm.CalendarPermissions.Add(perm);
        h.Vm.SelectedCalendarPermission = perm;
        h.Dialogs.ConfirmResult = false;

        await h.Vm.RemoveCalendarPermissionCommand.ExecuteAsync(null);

        h.Rules.Verify(r => r.RemoveCalendarPermissionAsync(It.IsAny<string>(), It.IsAny<string>(),
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
        h.Vm.CalendarPermissions.Should().Contain(perm);
    }
}
