using FluentAssertions;
using Grex365.App.ViewModels;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Moq;

namespace Grex365.App.Tests;

public class UserDetailsViewModelTests
{
    private static UserSummary SampleUser(bool enabled = true, int licenses = 2) =>
        new("uid-1", "Jane Doe", "jane@contoso.onmicrosoft.com", "jane@contoso.onmicrosoft.com", enabled, false, licenses, null);

    private static GroupSummary SampleGroup(string id, string name) =>
        new(id, name, $"{name}@contoso.onmicrosoft.com", "M365Group");

    private static SkuInfo SampleSku(string part, int enabled = 10, int consumed = 5) =>
        new(Guid.NewGuid(), part, enabled, consumed);

    private static async Task WaitForAsync(Func<bool> condition, int timeoutMs = 2000)
    {
        var deadline = DateTimeOffset.UtcNow.AddMilliseconds(timeoutMs);
        while (!condition() && DateTimeOffset.UtcNow < deadline)
        {
            await Task.Delay(10);
        }
    }

    private sealed class Harness
    {
        public Mock<IUsersService> Users { get; } = new(MockBehavior.Strict);
        public TestUserDetailsHost Host { get; } = new();
        public TestUiLogSink Log { get; } = new();
        public TestDialogService Dialogs { get; } = new();
        public TestClipboardService Clipboard { get; } = new();
        public UserDetailsViewModel Vm { get; }

        public Harness()
        {
            Vm = new UserDetailsViewModel(Users.Object, Host, Log, Dialogs, Clipboard);
        }

        public void StubLoadOk(UserSummary user, IReadOnlyList<GroupSummary>? groups = null,
            IReadOnlyList<SkuInfo>? allSkus = null, IReadOnlyList<Guid>? assigned = null)
        {
            Users.Setup(u => u.GetByIdAsync(user.Id, It.IsAny<CancellationToken>()))
                .ReturnsAsync(user);
            Users.Setup(u => u.GetGroupMembershipsAsync(user.Id, It.IsAny<CancellationToken>()))
                .ReturnsAsync(groups ?? Array.Empty<GroupSummary>());
            Users.Setup(u => u.ListSkusAsync(It.IsAny<CancellationToken>()))
                .ReturnsAsync(allSkus ?? Array.Empty<SkuInfo>());
            Users.Setup(u => u.GetAssignedLicensesAsync(user.Id, It.IsAny<CancellationToken>()))
                .ReturnsAsync(assigned ?? Array.Empty<Guid>());
        }
    }

    [Fact]
    public async Task OpenRequested_PopulatesUser_AndCollections()
    {
        var h = new Harness();
        var user = SampleUser();
        var groups = new[] { SampleGroup("g1", "Alpha"), SampleGroup("g2", "Beta") };
        var sku = SampleSku("ENTERPRISEPACK");
        h.StubLoadOk(user, groups, new[] { sku }, new[] { sku.SkuId });

        h.Host.RequestOpen(user.Id);

        await WaitForAsync(() => !h.Vm.IsBusy && h.Vm.HasUser);

        h.Vm.User.Should().Be(user);
        h.Vm.HasUser.Should().BeTrue();
        h.Vm.Memberships.Should().HaveCount(2);
        h.Vm.Memberships.Select(g => g.DisplayName).Should().ContainInOrder("Alpha", "Beta");
        h.Vm.AssignedLicenses.Should().ContainSingle()
            .Which.SkuPartNumber.Should().Be("ENTERPRISEPACK");
        h.Vm.AssignableSkus.Should().BeEmpty(); // all SKUs assigned, none available
    }

    [Fact]
    public async Task LoadAsync_UserNotFound_SetsErrorStatus()
    {
        var h = new Harness();
        h.Users.Setup(u => u.GetByIdAsync("ghost", It.IsAny<CancellationToken>()))
            .ReturnsAsync((UserSummary?)null);

        h.Host.RequestOpen("ghost");

        await WaitForAsync(() => !h.Vm.IsBusy);

        h.Vm.HasUser.Should().BeFalse();
        h.Vm.User.Should().BeNull();
        h.Vm.StatusMessage.Should().Be("Usuario no encontrado.");
    }

    [Fact]
    public async Task CloseRequested_ResetsState()
    {
        var h = new Harness();
        var user = SampleUser();
        h.StubLoadOk(user, new[] { SampleGroup("g1", "Alpha") });

        h.Host.RequestOpen(user.Id);
        await WaitForAsync(() => !h.Vm.IsBusy && h.Vm.HasUser);
        h.Vm.HasUser.Should().BeTrue();

        h.Host.RequestClose();

        h.Vm.HasUser.Should().BeFalse();
        h.Vm.User.Should().BeNull();
        h.Vm.UserId.Should().BeNull();
        h.Vm.Memberships.Should().BeEmpty();
        h.Vm.AssignedLicenses.Should().BeEmpty();
        h.Vm.StatusMessage.Should().BeEmpty();
    }

    [Fact]
    public async Task ToggleAccount_ConfirmNo_DoesNotCallService()
    {
        var h = new Harness();
        var user = SampleUser(enabled: true);
        h.StubLoadOk(user);
        h.Dialogs.ConfirmResult = false;

        h.Host.RequestOpen(user.Id);
        await WaitForAsync(() => h.Vm.HasUser);

        await h.Vm.ToggleAccountCommand.ExecuteAsync(null);

        h.Dialogs.Confirmations.Should().HaveCount(1);
        h.Users.Verify(u => u.SetAccountEnabledAsync(It.IsAny<string>(), It.IsAny<bool>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()),
            Times.Never);
    }

    [Fact]
    public async Task ToggleAccount_ConfirmYes_CallsSetAccountEnabledAsync_WithToggledValue()
    {
        var h = new Harness();
        var user = SampleUser(enabled: true);
        h.StubLoadOk(user);
        h.Dialogs.ConfirmResult = true;
        h.Users.Setup(u => u.SetAccountEnabledAsync(user.Id, false, It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        h.Host.RequestOpen(user.Id);
        await WaitForAsync(() => h.Vm.HasUser);

        await h.Vm.ToggleAccountCommand.ExecuteAsync(null);
        await WaitForAsync(() => !h.Vm.IsBusy);

        h.Users.Verify(u => u.SetAccountEnabledAsync(user.Id, false, It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public async Task RemoveLicense_NullRow_NoOp()
    {
        var h = new Harness();
        var user = SampleUser();
        h.StubLoadOk(user);
        h.Host.RequestOpen(user.Id);
        await WaitForAsync(() => h.Vm.HasUser);

        await h.Vm.RemoveLicenseCommand.ExecuteAsync(null);

        h.Dialogs.Confirmations.Should().BeEmpty();
        h.Users.Verify(u => u.RemoveLicenseAsync(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()),
            Times.Never);
    }

    [Fact]
    public async Task RemoveLicense_ConfirmYes_CallsService()
    {
        var h = new Harness();
        var user = SampleUser();
        var sku = SampleSku("E3");
        h.StubLoadOk(user, allSkus: new[] { sku }, assigned: new[] { sku.SkuId });
        h.Dialogs.ConfirmResult = true;
        h.Users.Setup(u => u.RemoveLicenseAsync(user.Id, sku.SkuId, It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        h.Host.RequestOpen(user.Id);
        await WaitForAsync(() => h.Vm.HasUser && h.Vm.AssignedLicenses.Count == 1);

        var row = h.Vm.AssignedLicenses[0];
        await h.Vm.RemoveLicenseCommand.ExecuteAsync(row);
        await WaitForAsync(() => !h.Vm.IsBusy);

        h.Users.Verify(u => u.RemoveLicenseAsync(user.Id, sku.SkuId, It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public async Task AssignSelectedSku_NoUser_NoOp()
    {
        var h = new Harness();
        h.Vm.SelectedSkuToAdd = SampleSku("E3");

        await h.Vm.AssignSelectedSkuCommand.ExecuteAsync(null);

        h.Users.Verify(u => u.AssignLicenseAsync(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()),
            Times.Never);
    }

    [Fact]
    public async Task AssignSelectedSku_NoSku_NoOp()
    {
        var h = new Harness();
        var user = SampleUser();
        h.StubLoadOk(user);
        h.Host.RequestOpen(user.Id);
        await WaitForAsync(() => h.Vm.HasUser);
        h.Vm.SelectedSkuToAdd = null;

        await h.Vm.AssignSelectedSkuCommand.ExecuteAsync(null);

        h.Users.Verify(u => u.AssignLicenseAsync(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()),
            Times.Never);
    }

    [Fact]
    public async Task AssignSelectedSku_Valid_CallsServiceAndClearsSelection()
    {
        var h = new Harness();
        var user = SampleUser();
        var assignedSku = SampleSku("E3");
        var availableSku = SampleSku("ENTERPRISEPACK");
        h.StubLoadOk(user, allSkus: new[] { assignedSku, availableSku }, assigned: new[] { assignedSku.SkuId });
        h.Users.Setup(u => u.AssignLicenseAsync(user.Id, availableSku.SkuId, It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        h.Host.RequestOpen(user.Id);
        await WaitForAsync(() => h.Vm.HasUser && h.Vm.AssignableSkus.Count == 1);

        h.Vm.SelectedSkuToAdd = h.Vm.AssignableSkus[0];
        await h.Vm.AssignSelectedSkuCommand.ExecuteAsync(null);
        await WaitForAsync(() => !h.Vm.IsBusy);

        h.Users.Verify(u => u.AssignLicenseAsync(user.Id, availableSku.SkuId, It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()),
            Times.Once);
        h.Vm.SelectedSkuToAdd.Should().BeNull();
    }

    [Fact]
    public async Task ResetPassword_ConfirmYes_CopiesToClipboard_AndShowsDialog()
    {
        var h = new Harness();
        var user = SampleUser();
        h.StubLoadOk(user);
        h.Dialogs.ConfirmResult = true;
        h.Users.Setup(u => u.ResetPasswordAsync(user.Id, true, It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync("Temp-Pw-123!");

        h.Host.RequestOpen(user.Id);
        await WaitForAsync(() => h.Vm.HasUser);

        await h.Vm.ResetPasswordCommand.ExecuteAsync(null);
        await WaitForAsync(() => !h.Vm.IsBusy);

        h.Clipboard.LastValue.Should().Be("Temp-Pw-123!");
        h.Clipboard.CallCount.Should().Be(1);
        h.Dialogs.Shows.Should().ContainSingle()
            .Which.Title.Should().Be("Password reseteada");
    }

    [Fact]
    public async Task ResetPassword_ConfirmNo_DoesNotCallService()
    {
        var h = new Harness();
        var user = SampleUser();
        h.StubLoadOk(user);
        h.Dialogs.ConfirmResult = false;

        h.Host.RequestOpen(user.Id);
        await WaitForAsync(() => h.Vm.HasUser);

        await h.Vm.ResetPasswordCommand.ExecuteAsync(null);

        h.Users.Verify(u => u.ResetPasswordAsync(It.IsAny<string>(), It.IsAny<bool>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()),
            Times.Never);
        h.Clipboard.CallCount.Should().Be(0);
    }

    [Fact]
    public async Task RevokeSessions_ConfirmYes_CallsService()
    {
        var h = new Harness();
        var user = SampleUser();
        h.StubLoadOk(user);
        h.Dialogs.ConfirmResult = true;
        h.Users.Setup(u => u.RevokeSignInSessionsAsync(user.Id, It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        h.Host.RequestOpen(user.Id);
        await WaitForAsync(() => h.Vm.HasUser);

        await h.Vm.RevokeSessionsCommand.ExecuteAsync(null);
        await WaitForAsync(() => !h.Vm.IsBusy);

        h.Users.Verify(u => u.RevokeSignInSessionsAsync(user.Id, It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()),
            Times.Once);
        h.Dialogs.Confirmations.Should().HaveCount(1);
        h.Dialogs.Confirmations[0].Icon.Should().Be(DialogIcon.Warning);
    }

    [Fact]
    public async Task RemoveAllLicenses_ConfirmNo_DoesNotCallService()
    {
        var h = new Harness();
        var user = SampleUser();
        h.StubLoadOk(user);
        h.Dialogs.ConfirmResult = false;

        h.Host.RequestOpen(user.Id);
        await WaitForAsync(() => h.Vm.HasUser);

        await h.Vm.RemoveAllLicensesCommand.ExecuteAsync(null);

        h.Users.Verify(u => u.RemoveAllLicensesAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()),
            Times.Never);
    }

    [Fact]
    public async Task RemoveAllLicenses_ConfirmYes_CallsService()
    {
        var h = new Harness();
        var user = SampleUser();
        h.StubLoadOk(user);
        h.Dialogs.ConfirmResult = true;
        h.Users.Setup(u => u.RemoveAllLicensesAsync(user.Id, It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        h.Host.RequestOpen(user.Id);
        await WaitForAsync(() => h.Vm.HasUser);

        await h.Vm.RemoveAllLicensesCommand.ExecuteAsync(null);
        await WaitForAsync(() => !h.Vm.IsBusy);

        h.Users.Verify(u => u.RemoveAllLicensesAsync(user.Id, It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()),
            Times.Once);
        h.Dialogs.Confirmations.Should().HaveCount(1);
        h.Dialogs.Confirmations[0].Icon.Should().Be(DialogIcon.Warning);
    }

    [Fact]
    public async Task Close_Command_CallsHostRequestClose()
    {
        var h = new Harness();
        var user = SampleUser();
        h.StubLoadOk(user);
        h.Host.RequestOpen(user.Id);
        await WaitForAsync(() => h.Vm.HasUser);

        h.Vm.CloseCommand.Execute(null);

        h.Vm.HasUser.Should().BeFalse();
    }
}
