using FluentAssertions;
using Grex365.App.ViewModels;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Moq;

namespace Grex365.App.Tests;

public class GroupsViewModelTests
{
    private static GroupSummary SampleGroup(string id = "g1", string name = "Alpha") =>
        new(id, name, $"{name}@contoso.onmicrosoft.com", "M365Group");

    private static GroupMember SampleMember(string id = "u1", string? name = "Jane Doe") =>
        new(id, name, "jane@contoso.onmicrosoft.com", "jane@contoso.onmicrosoft.com");

    private sealed class Harness
    {
        public Mock<IGroupsService> Groups { get; } = new();
        public Mock<IDistributionListsService> Dls { get; } = new();
        public Mock<IRbacGuard> Rbac { get; } = new();
        public TestUiLogSink Log { get; } = new();
        public TestDialogService Dialogs { get; } = new();
        public GroupsViewModel Vm { get; }

        public Harness(bool rbacAllowed = true)
        {
            Rbac.Setup(r => r.EvaluateAsync(It.IsAny<CancellationToken>()))
                .ReturnsAsync(new RbacDecision(rbacAllowed, rbacAllowed ? "OK" : "Not in group"));
            Vm = new GroupsViewModel(Groups.Object, Dls.Object, Log, Rbac.Object, Dialogs);
        }
    }

    [Fact]
    public async Task RemoveMember_NoGroupOrMember_SetsStatus()
    {
        var h = new Harness();

        await h.Vm.RemoveSelectedMemberCommand.ExecuteAsync(null);

        h.Vm.StatusMessage.Should().Be("Selecciona un miembro.");
        h.Dialogs.Confirmations.Should().BeEmpty();
    }

    [Fact]
    public async Task RemoveMember_RbacDenied_Blocks()
    {
        var h = new Harness(rbacAllowed: false);
        h.Vm.SelectedGroup = SampleGroup();
        h.Vm.SelectedMember = SampleMember();

        await h.Vm.RemoveSelectedMemberCommand.ExecuteAsync(null);

        h.Vm.StatusMessage.Should().Be("Not in group");
        h.Dialogs.Confirmations.Should().BeEmpty();
        h.Groups.Verify(g => g.RemoveMemberAsync(It.IsAny<string>(), It.IsAny<string>(),
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task RemoveMember_ConfirmNo_DoesNothing()
    {
        var h = new Harness();
        h.Vm.SelectedGroup = SampleGroup();
        h.Vm.SelectedMember = SampleMember();
        h.Dialogs.ConfirmResult = false;

        await h.Vm.RemoveSelectedMemberCommand.ExecuteAsync(null);

        h.Vm.StatusMessage.Should().Be("Cancelado por el usuario.");
        h.Dialogs.Confirmations.Should().HaveCount(1);
        h.Dialogs.Confirmations[0].Icon.Should().Be(DialogIcon.Warning);
        h.Groups.Verify(g => g.RemoveMemberAsync(It.IsAny<string>(), It.IsAny<string>(),
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task RemoveMember_ConfirmYes_CallsService_AndRemovesFromCollection()
    {
        var h = new Harness();
        var group = SampleGroup();
        var member = SampleMember();
        h.Vm.SelectedGroup = group;
        h.Vm.SelectedMember = member;
        h.Vm.Members.Add(member);
        h.Dialogs.ConfirmResult = true;
        h.Groups.Setup(g => g.RemoveMemberAsync(group.Id, member.Id,
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>())).Returns(Task.CompletedTask);

        await h.Vm.RemoveSelectedMemberCommand.ExecuteAsync(null);

        h.Groups.Verify(g => g.RemoveMemberAsync(group.Id, member.Id,
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Once);
        h.Vm.Members.Should().NotContain(member);
        h.Vm.StatusMessage.Should().Be("Eliminado: Jane Doe");
    }

    [Fact]
    public void AddPickedMember_AppendsToBox_AndClearsPicker()
    {
        var h = new Harness();
        h.Vm.NewMembersText = "a@x.com";
        h.Vm.MemberToAdd = "b@x.com";

        h.Vm.AddPickedMemberCommand.Execute(null);

        h.Vm.NewMembersText.Should().Be("a@x.com" + Environment.NewLine + "b@x.com");
        h.Vm.MemberToAdd.Should().BeEmpty();
    }

    [Fact]
    public void AddPickedMember_Duplicate_DoesNotDuplicate()
    {
        var h = new Harness();
        h.Vm.NewMembersText = "a@x.com";
        h.Vm.MemberToAdd = "A@X.com";

        h.Vm.AddPickedMemberCommand.Execute(null);

        h.Vm.NewMembersText.Should().Be("a@x.com");
        h.Vm.MemberToAdd.Should().BeEmpty();
    }

    [Fact]
    public void AddPickedMember_Empty_NoOp()
    {
        var h = new Harness();
        h.Vm.NewMembersText = "a@x.com";
        h.Vm.MemberToAdd = "   ";

        h.Vm.AddPickedMemberCommand.Execute(null);

        h.Vm.NewMembersText.Should().Be("a@x.com");
    }

    [Fact]
    public async Task RemoveMember_ServiceThrows_StatusAndLogError()
    {
        var h = new Harness();
        var group = SampleGroup();
        var member = SampleMember();
        h.Vm.SelectedGroup = group;
        h.Vm.SelectedMember = member;
        h.Vm.Members.Add(member);
        h.Dialogs.ConfirmResult = true;
        h.Groups.Setup(g => g.RemoveMemberAsync(group.Id, member.Id,
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("forbidden"));

        await h.Vm.RemoveSelectedMemberCommand.ExecuteAsync(null);

        h.Vm.StatusMessage.Should().StartWith("Error: ");
        h.Vm.Members.Should().Contain(member); // not removed because service threw
        h.Log.Entries.Should().Contain(e => e.Severity == LogSeverity.Error && e.Source == "Groups");
    }
}
