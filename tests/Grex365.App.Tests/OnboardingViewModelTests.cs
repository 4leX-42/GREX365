using FluentAssertions;
using Grex365.App.ViewModels;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Moq;

namespace Grex365.App.Tests;

public class OnboardingViewModelTests
{
    private sealed class Harness
    {
        public Mock<IOnboardingService> Onboarding { get; } = new();
        public Mock<IUsersService> Users { get; } = new();
        public TestUiLogSink Log { get; } = new();
        public TestDialogService Dialogs { get; } = new();
        public OnboardingViewModel Vm { get; }

        public Harness()
        {
            Vm = new OnboardingViewModel(Onboarding.Object, Users.Object, Log, Dialogs);
        }

        public void StubRunOk()
        {
            Onboarding.Setup(o => o.RunAsync(It.IsAny<OnboardingOptions>(),
                It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
                .ReturnsAsync(new OnboardingResult("u@a", "uid-1", true, new List<OnboardingStep>
                {
                    new("Crear", "OK", ""),
                    new("Asignar SKUs", "OK", ""),
                }));
        }
    }

    [Fact]
    public async Task Run_ConfirmNo_DoesNotInvokeService()
    {
        var h = new Harness();
        h.Vm.DisplayName = "Jane Doe";
        h.Vm.Upn = "jane@a";
        h.Vm.InitialPassword = "Temp-1!!";
        h.Vm.UsageLocation = "ES";
        h.Dialogs.ConfirmResult = false;

        await h.Vm.RunCommand.ExecuteAsync(null);

        h.Vm.StatusMessage.Should().Be("Cancelado por el usuario.");
        h.Dialogs.Confirmations.Should().HaveCount(1);
        h.Onboarding.Verify(o => o.RunAsync(It.IsAny<OnboardingOptions>(),
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task Run_ConfirmYes_InvokesService_PopulatesSteps()
    {
        var h = new Harness();
        h.Vm.DisplayName = "Jane Doe";
        h.Vm.Upn = "  jane@a  ";
        h.Vm.InitialPassword = "Temp-1!!";
        h.Vm.UsageLocation = "es";
        h.Vm.GroupsText = "g1, g2";
        h.Dialogs.ConfirmResult = true;
        h.StubRunOk();

        await h.Vm.RunCommand.ExecuteAsync(null);

        h.Onboarding.Verify(o => o.RunAsync(
            It.Is<OnboardingOptions>(opt =>
                opt.DisplayName == "Jane Doe"
                && opt.Upn == "jane@a"
                && opt.UsageLocation == "ES"
                && opt.GroupIdentifiers.Count == 2),
            It.IsAny<IProgress<LogEntry>>(),
            It.IsAny<CancellationToken>()), Times.Once);
        h.Vm.Steps.Should().HaveCount(2);
        h.Vm.StatusMessage.Should().StartWith("Onboarding OK");
    }

    [Fact]
    public async Task Run_ServiceReportsFailure_StatusIncludesErrorCount()
    {
        var h = new Harness();
        h.Vm.DisplayName = "Jane Doe";
        h.Vm.Upn = "jane@a";
        h.Vm.InitialPassword = "Temp-1!!";
        h.Vm.UsageLocation = "ES";
        h.Dialogs.ConfirmResult = true;
        h.Onboarding.Setup(o => o.RunAsync(It.IsAny<OnboardingOptions>(),
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new OnboardingResult("jane@a", null, false, new[]
            {
                new OnboardingStep("Crear", "OK", ""),
                new OnboardingStep("Asignar SKU", "ERROR", "boom"),
            }));

        await h.Vm.RunCommand.ExecuteAsync(null);

        h.Vm.StatusMessage.Should().StartWith("Onboarding con errores").And.Contain("1 fallos");
    }

    [Fact]
    public void AddPickedGroup_AppendsToBox_AndClearsPicker()
    {
        var h = new Harness();
        h.Vm.GroupsText = "Sales";
        h.Vm.GroupToAdd = "marketing@contoso.com";

        h.Vm.AddPickedGroupCommand.Execute(null);

        h.Vm.GroupsText.Should().Be("Sales" + Environment.NewLine + "marketing@contoso.com");
        h.Vm.GroupToAdd.Should().BeEmpty();
    }

    [Fact]
    public void AddPickedGroup_Empty_NoOp()
    {
        var h = new Harness();
        h.Vm.GroupsText = "Sales";
        h.Vm.GroupToAdd = "  ";

        h.Vm.AddPickedGroupCommand.Execute(null);

        h.Vm.GroupsText.Should().Be("Sales");
    }

    [Fact]
    public async Task AddSelectedSku_AppendsUnique()
    {
        var h = new Harness();
        var sku = new SkuInfo(Guid.NewGuid(), "E3", 10, 5);
        h.Vm.SelectedSku = sku;

        h.Vm.AddSelectedSkuCommand.Execute(null);
        h.Vm.AddSelectedSkuCommand.Execute(null); // duplicate ignored

        h.Vm.SelectedSkus.Should().ContainSingle().Which.SkuId.Should().Be(sku.SkuId);
    }

    [Fact]
    public void RemoveSku_RemovesFromSelected()
    {
        var h = new Harness();
        var sku = new SkuInfo(Guid.NewGuid(), "E3", 10, 5);
        h.Vm.SelectedSkus.Add(sku);

        h.Vm.RemoveSkuCommand.Execute(sku);

        h.Vm.SelectedSkus.Should().BeEmpty();
    }
}
