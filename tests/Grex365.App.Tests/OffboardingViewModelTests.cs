using FluentAssertions;
using Grex365.App.ViewModels;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Moq;

namespace Grex365.App.Tests;

public class OffboardingViewModelTests
{
    private sealed class Harness
    {
        public Mock<IOffboardingService> Service { get; } = new(MockBehavior.Strict);
        public Mock<IRbacGuard> Rbac { get; } = new();
        public TestUiLogSink Log { get; } = new();
        public TestDialogService Dialogs { get; } = new();
        public Mock<IClipboardService> Clipboard { get; } = new();
        public OffboardingViewModel Vm { get; }

        public Harness(bool rbacAllowed = true)
        {
            Rbac.Setup(r => r.EvaluateAsync(It.IsAny<CancellationToken>()))
                .ReturnsAsync(new RbacDecision(rbacAllowed, rbacAllowed ? "OK" : "Not in group"));
            Vm = new OffboardingViewModel(Service.Object, Log, Rbac.Object, Dialogs, clipboard: Clipboard.Object);
        }

        public void StubRunOk()
        {
            Service.Setup(s => s.RunAsync(It.IsAny<string>(), It.IsAny<OffboardingOptions>(),
                It.IsAny<IProgress<LogEntry>>(), It.IsAny<IProgress<OffboardingStep>>(), It.IsAny<CancellationToken>()))
                .ReturnsAsync(new OffboardingResult("u@a", true, new List<OffboardingStep>
                {
                    new("Buscar", "OK", "found"),
                    new("Deshabilitar", "OK", "done"),
                }));
        }
    }

    [Fact]
    public async Task EmptyUpn_SetsStatus_AndDoesNotPrompt()
    {
        var h = new Harness();
        h.Vm.Upn = string.Empty;

        await h.Vm.RunCommand.ExecuteAsync(null);

        h.Vm.StatusMessage.Should().Be("UPN vacío.");
        h.Dialogs.Confirmations.Should().BeEmpty();
        h.Service.VerifyNoOtherCalls();
    }

    [Fact]
    public async Task NoActionsSelected_SetsStatus_AndDoesNotRun()
    {
        var h = new Harness();
        h.Vm.Upn = "jane@a";
        h.Vm.DisableAccount = false;
        h.Vm.RemoveLicenses = false;
        h.Vm.ConvertMailboxToShared = false;

        await h.Vm.RunCommand.ExecuteAsync(null);

        h.Vm.StatusMessage.Should().Be("Ninguna acción seleccionada.");
        h.Dialogs.Confirmations.Should().BeEmpty();
    }

    [Fact]
    public async Task RbacDenied_BlocksRun_AndLogsWarn()
    {
        var h = new Harness(rbacAllowed: false);
        h.Vm.Upn = "jane@a";

        await h.Vm.RunCommand.ExecuteAsync(null);

        h.Vm.StatusMessage.Should().Be("Not in group");
        h.Log.Entries.Should().ContainSingle(e => e.Severity == LogSeverity.Warning
            && e.Source == "RBAC"
            && e.Message.Contains("Offboarding"));
        h.Dialogs.Confirmations.Should().BeEmpty();
    }

    [Fact]
    public async Task ConfirmNo_DoesNotInvokeService()
    {
        var h = new Harness();
        h.Vm.Upn = "jane@a";
        h.Dialogs.ConfirmResult = false;

        await h.Vm.RunCommand.ExecuteAsync(null);

        h.Vm.StatusMessage.Should().Be("Cancelado por el usuario.");
        h.Dialogs.Confirmations.Should().HaveCount(1);
        h.Dialogs.Confirmations[0].Icon.Should().Be(DialogIcon.Warning);
        h.Service.VerifyNoOtherCalls();
    }

    [Fact]
    public async Task ConfirmYes_RunsService_PopulatesSteps()
    {
        var h = new Harness();
        h.Vm.Upn = "  jane@a  ";
        h.Vm.DisableAccount = true;
        h.Vm.RemoveLicenses = false;
        h.Vm.ConvertMailboxToShared = false;
        h.Dialogs.ConfirmResult = true;
        h.StubRunOk();

        await h.Vm.RunCommand.ExecuteAsync(null);

        h.Service.Verify(s => s.RunAsync(
            "jane@a",
            It.Is<OffboardingOptions>(o => o.DisableAccount && !o.RemoveLicenses && !o.ConvertMailboxToShared),
            It.IsAny<IProgress<LogEntry>>(),
            It.IsAny<IProgress<OffboardingStep>>(),
            It.IsAny<CancellationToken>()), Times.Once);
        h.Vm.Result.Should().NotBeNull();
        h.Vm.Steps.Should().HaveCount(2);
        h.Vm.StatusMessage.Should().StartWith("Offboarding OK");
    }

    [Fact]
    public async Task ServiceReportsFailure_StatusIncludesErrorCount()
    {
        var h = new Harness();
        h.Vm.Upn = "jane@a";
        h.Dialogs.ConfirmResult = true;
        h.Service.Setup(s => s.RunAsync(It.IsAny<string>(), It.IsAny<OffboardingOptions>(),
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<IProgress<OffboardingStep>>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new OffboardingResult("jane@a", false, new[]
            {
                new OffboardingStep("Buscar", "OK", ""),
                new OffboardingStep("Deshabilitar", "ERROR", "boom"),
                new OffboardingStep("Licencias", "ERROR", "boom2"),
            }));

        await h.Vm.RunCommand.ExecuteAsync(null);

        h.Vm.StatusMessage.Should().StartWith("Offboarding con errores").And.Contain("2 fallos");
    }

    [Fact]
    public async Task DryRunChecked_PassesDryRunOption()
    {
        var h = new Harness();
        h.Vm.Upn = "jane@a";
        h.Vm.DisableAccount = true;
        h.Vm.RemoveLicenses = false;
        h.Vm.ConvertMailboxToShared = false;
        h.Vm.DryRun = true;
        h.Dialogs.ConfirmResult = true;
        h.StubRunOk();

        await h.Vm.RunCommand.ExecuteAsync(null);

        h.Service.Verify(s => s.RunAsync(
            "jane@a",
            It.Is<OffboardingOptions>(o => o.DryRun),
            It.IsAny<IProgress<LogEntry>>(),
            It.IsAny<IProgress<OffboardingStep>>(),
            It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public void ExportResults_NoResults_SetsEmptyStatus()
    {
        var h = new Harness();

        h.Vm.ExportResultsCommand.Execute(null);

        h.Vm.StatusMessage.Should().Be("No hay resultados que exportar. Ejecuta primero.");
        h.Clipboard.Verify(c => c.SetText(It.IsAny<string>()), Times.Never);
    }

    [Fact]
    public async Task ExportResults_AfterRun_CopiesCsvToClipboard()
    {
        var h = new Harness();
        h.Vm.Upn = "jane@a";
        h.Dialogs.ConfirmResult = true;
        h.StubRunOk();
        await h.Vm.RunCommand.ExecuteAsync(null);

        string? captured = null;
        h.Clipboard.Setup(c => c.SetText(It.IsAny<string>())).Callback<string>(s => captured = s);

        h.Vm.ExportResultsCommand.Execute(null);

        captured.Should().NotBeNull();
        captured!.Should().StartWith("upn,success,dryRun");
        captured.Should().Contain("u@a");        // result UPN from StubRunOk
        captured.Should().Contain("Deshabilitar"); // a step name from StubRunOk
        h.Vm.StatusMessage.Should().StartWith("CSV copiado");
    }

    [Fact]
    public async Task FinalizationOptions_FlowIntoService()
    {
        var h = new Harness();
        h.Vm.Upn = "jane@a";
        h.Vm.DelegateToAll = "deleg@a";
        h.Vm.ForwardToDelegate = true;
        h.Vm.HideFromGal = true;
        h.Vm.AutoReplyMessage = "Ya no trabaja aquí";
        h.Dialogs.ConfirmResult = true;
        h.StubRunOk();

        await h.Vm.RunCommand.ExecuteAsync(null);

        h.Service.Verify(s => s.RunAsync(
            "jane@a",
            It.Is<OffboardingOptions>(o =>
                o.ForwardTo == "deleg@a" && o.HideFromGal && o.AutoReplyMessage == "Ya no trabaja aquí"),
            It.IsAny<IProgress<LogEntry>>(),
            It.IsAny<IProgress<OffboardingStep>>(),
            It.IsAny<CancellationToken>()), Times.Once);
    }
}
