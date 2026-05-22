using FluentAssertions;
using Grex365.App.ViewModels;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Moq;

namespace Grex365.App.Tests;

public class UsersViewModelTests
{
    private static UserSummary SampleUser(bool enabled = true, int licenses = 2) =>
        new("uid-1", "Jane Doe", "jane@contoso.onmicrosoft.com", "jane@contoso.onmicrosoft.com", enabled, false, licenses, null);

    private sealed class Harness
    {
        public Mock<IUsersService> Users { get; } = new();
        public Mock<IRbacGuard> Rbac { get; } = new();
        public TestUiLogSink Log { get; } = new();
        public TestDialogService Dialogs { get; } = new();
        public UsersViewModel Vm { get; }

        public Harness(bool rbacAllowed = true)
        {
            Rbac.Setup(r => r.EvaluateAsync(It.IsAny<CancellationToken>()))
                .ReturnsAsync(new RbacDecision(rbacAllowed, rbacAllowed ? "OK" : "Not in group"));
            Vm = new UsersViewModel(Users.Object, Log, Rbac.Object, Dialogs);
        }
    }

    [Fact]
    public async Task Disable_NoUserSelected_SetsStatus()
    {
        var h = new Harness();
        h.Vm.SelectedUser = null;

        await h.Vm.DisableCommand.ExecuteAsync(null);

        h.Vm.StatusMessage.Should().Be("Selecciona un usuario.");
        h.Dialogs.Confirmations.Should().BeEmpty();
    }

    [Fact]
    public async Task Disable_RbacDenied_BlocksWithoutConfirm()
    {
        var h = new Harness(rbacAllowed: false);
        h.Vm.SelectedUser = SampleUser(enabled: true);

        await h.Vm.DisableCommand.ExecuteAsync(null);

        h.Vm.StatusMessage.Should().Be("Not in group");
        h.Dialogs.Confirmations.Should().BeEmpty();
        h.Users.Verify(u => u.SetAccountEnabledAsync(It.IsAny<string>(), It.IsAny<bool>(),
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task Disable_ConfirmNo_DoesNotInvokeService()
    {
        var h = new Harness();
        h.Vm.SelectedUser = SampleUser(enabled: true);
        h.Dialogs.ConfirmResult = false;

        await h.Vm.DisableCommand.ExecuteAsync(null);

        h.Vm.StatusMessage.Should().Be("Cancelado por el usuario.");
        h.Dialogs.Confirmations.Should().HaveCount(1);
        h.Dialogs.Confirmations[0].Icon.Should().Be(DialogIcon.Warning);
        h.Users.Verify(u => u.SetAccountEnabledAsync(It.IsAny<string>(), It.IsAny<bool>(),
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task Disable_ConfirmYes_CallsService_WithFalse()
    {
        var h = new Harness();
        var user = SampleUser(enabled: true);
        h.Vm.SelectedUser = user;
        h.Dialogs.ConfirmResult = true;
        h.Users.Setup(u => u.SetAccountEnabledAsync(user.Id, false, It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);
        h.Users.Setup(u => u.GetByIdAsync(user.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(user with { AccountEnabled = false });

        await h.Vm.DisableCommand.ExecuteAsync(null);

        h.Users.Verify(u => u.SetAccountEnabledAsync(user.Id, false, It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
        h.Vm.StatusMessage.Should().Be("Deshabilitado.");
    }

    [Fact]
    public async Task Enable_DoesNotRequireConfirm_NorRbac()
    {
        var h = new Harness(rbacAllowed: false); // even with rbac denied: Enable is constructive, no gate
        var user = SampleUser(enabled: false);
        h.Vm.SelectedUser = user;
        h.Users.Setup(u => u.SetAccountEnabledAsync(user.Id, true, It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);
        h.Users.Setup(u => u.GetByIdAsync(user.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(user with { AccountEnabled = true });

        await h.Vm.EnableCommand.ExecuteAsync(null);

        h.Dialogs.Confirmations.Should().BeEmpty();
        h.Users.Verify(u => u.SetAccountEnabledAsync(user.Id, true, It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
        h.Vm.StatusMessage.Should().Be("Habilitado.");
    }

    [Fact]
    public async Task AssignLicense_NoUser_SetsStatus()
    {
        var h = new Harness();
        h.Vm.SelectedUser = null;
        h.Vm.SelectedSku = new SkuInfo(Guid.NewGuid(), "E3", 10, 5);

        await h.Vm.AssignLicenseCommand.ExecuteAsync(null);

        h.Vm.StatusMessage.Should().Be("Selecciona un usuario.");
    }

    [Fact]
    public async Task AssignLicense_NoSku_SetsStatus()
    {
        var h = new Harness();
        h.Vm.SelectedUser = SampleUser();
        h.Vm.SelectedSku = null;

        await h.Vm.AssignLicenseCommand.ExecuteAsync(null);

        h.Vm.StatusMessage.Should().Contain("Cargar SKUs");
    }

    [Fact]
    public async Task AssignLicense_SkuWithSeats_NoConfirmNeeded()
    {
        var h = new Harness();
        var user = SampleUser();
        var sku = new SkuInfo(Guid.NewGuid(), "E3", 10, 5); // 5 available
        h.Vm.SelectedUser = user;
        h.Vm.SelectedSku = sku;
        h.Users.Setup(u => u.AssignLicenseAsync(user.Id, sku.SkuId, It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);
        h.Users.Setup(u => u.GetByIdAsync(user.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(user);

        await h.Vm.AssignLicenseCommand.ExecuteAsync(null);

        h.Dialogs.Confirmations.Should().BeEmpty();
        h.Users.Verify(u => u.AssignLicenseAsync(user.Id, sku.SkuId, It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task AssignLicense_SkuNoSeats_ConfirmYes_StillAssigns()
    {
        var h = new Harness();
        var user = SampleUser();
        var sku = new SkuInfo(Guid.NewGuid(), "E3", 5, 5); // 0 available
        h.Vm.SelectedUser = user;
        h.Vm.SelectedSku = sku;
        h.Dialogs.ConfirmResult = true;
        h.Users.Setup(u => u.AssignLicenseAsync(user.Id, sku.SkuId, It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);
        h.Users.Setup(u => u.GetByIdAsync(user.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(user);

        await h.Vm.AssignLicenseCommand.ExecuteAsync(null);

        h.Dialogs.Confirmations.Should().HaveCount(1);
        h.Users.Verify(u => u.AssignLicenseAsync(user.Id, sku.SkuId, It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task AssignLicense_SkuNoSeats_ConfirmNo_DoesNotAssign()
    {
        var h = new Harness();
        var user = SampleUser();
        var sku = new SkuInfo(Guid.NewGuid(), "E3", 5, 5);
        h.Vm.SelectedUser = user;
        h.Vm.SelectedSku = sku;
        h.Dialogs.ConfirmResult = false;

        await h.Vm.AssignLicenseCommand.ExecuteAsync(null);

        h.Vm.StatusMessage.Should().Be("Cancelado por el usuario.");
        h.Users.Verify(u => u.AssignLicenseAsync(It.IsAny<string>(), It.IsAny<Guid>(),
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task RemoveLicenses_NoUser_SetsStatus()
    {
        var h = new Harness();
        h.Vm.SelectedUser = null;

        await h.Vm.RemoveLicensesCommand.ExecuteAsync(null);

        h.Vm.StatusMessage.Should().Be("Selecciona un usuario.");
    }

    [Fact]
    public async Task RemoveLicenses_ZeroLicenses_SkipsConfirm()
    {
        var h = new Harness();
        var user = SampleUser(licenses: 0);
        h.Vm.SelectedUser = user;
        h.Users.Setup(u => u.RemoveAllLicensesAsync(user.Id, It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);
        h.Users.Setup(u => u.GetByIdAsync(user.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(user);

        await h.Vm.RemoveLicensesCommand.ExecuteAsync(null);

        h.Dialogs.Confirmations.Should().BeEmpty();
        h.Users.Verify(u => u.RemoveAllLicensesAsync(user.Id, It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task RemoveLicenses_RbacDenied_Blocks()
    {
        var h = new Harness(rbacAllowed: false);
        var user = SampleUser(licenses: 3);
        h.Vm.SelectedUser = user;

        await h.Vm.RemoveLicensesCommand.ExecuteAsync(null);

        h.Vm.StatusMessage.Should().Be("Not in group");
        h.Dialogs.Confirmations.Should().BeEmpty();
        h.Users.Verify(u => u.RemoveAllLicensesAsync(It.IsAny<string>(),
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task RemoveLicenses_ConfirmNo_DoesNothing()
    {
        var h = new Harness();
        var user = SampleUser(licenses: 3);
        h.Vm.SelectedUser = user;
        h.Dialogs.ConfirmResult = false;

        await h.Vm.RemoveLicensesCommand.ExecuteAsync(null);

        h.Vm.StatusMessage.Should().Be("Cancelado por el usuario.");
        h.Users.Verify(u => u.RemoveAllLicensesAsync(It.IsAny<string>(),
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task RemoveLicenses_ConfirmYes_CallsService()
    {
        var h = new Harness();
        var user = SampleUser(licenses: 3);
        h.Vm.SelectedUser = user;
        h.Dialogs.ConfirmResult = true;
        h.Users.Setup(u => u.RemoveAllLicensesAsync(user.Id, It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);
        h.Users.Setup(u => u.GetByIdAsync(user.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(user with { AssignedLicenseCount = 0 });

        await h.Vm.RemoveLicensesCommand.ExecuteAsync(null);

        h.Users.Verify(u => u.RemoveAllLicensesAsync(user.Id, It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
        h.Dialogs.Confirmations[0].Icon.Should().Be(DialogIcon.Warning);
    }
}
