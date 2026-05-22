using FluentAssertions;
using Grex365.App.ViewModels;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Moq;

namespace Grex365.App.Tests;

public class PsConsoleViewModelTests
{
    private sealed class Harness
    {
        public Mock<IPowerShellRunner> Runner { get; } = new();
        public TestUiLogSink Log { get; } = new();
        public PsConsoleViewModel Vm { get; }

        public Harness()
        {
            Vm = new PsConsoleViewModel(Runner.Object, Log);
        }

        public void StubOk(string? value = "result", string[]? errors = null)
        {
            Runner.Setup(r => r.RunAsync(It.IsAny<string>(), It.IsAny<IDictionary<string, object?>>(),
                It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
                .ReturnsAsync(new PowerShellResult(
                    Success: errors is null || errors.Length == 0,
                    Output: value is null ? Array.Empty<object?>() : new object?[] { value },
                    Errors: errors ?? Array.Empty<string>()));
        }
    }

    [Fact]
    public async Task Run_EmptyInput_SetsStatus_DoesNotInvokeRunner()
    {
        var h = new Harness();
        h.Vm.InputText = "   ";

        await h.Vm.RunCommand.ExecuteAsync(null);

        h.Vm.StatusMessage.Should().Be("Comando vacío.");
        h.Runner.VerifyNoOtherCalls();
    }

    [Fact]
    public async Task Run_ValidScript_AppendsOutput_AndHistory()
    {
        var h = new Harness();
        h.Vm.InputText = "Get-Date";
        h.StubOk("2026-05-22");

        await h.Vm.RunCommand.ExecuteAsync(null);

        h.Vm.OutputText.Should().Contain("PS> Get-Date").And.Contain("2026-05-22");
        h.Vm.History.Should().ContainSingle().Which.Should().Be("Get-Date");
        h.Vm.StatusMessage.Should().StartWith("OK");
    }

    [Fact]
    public async Task Run_ServiceReportsErrors_StatusIncludesErrorCount_AndAppendsErrors()
    {
        var h = new Harness();
        h.Vm.InputText = "Bad-Cmd";
        h.StubOk(value: null, errors: new[] { "Cmd not found", "Other" });

        await h.Vm.RunCommand.ExecuteAsync(null);

        h.Vm.StatusMessage.Should().StartWith("ERROR").And.Contain("2 error(es)");
        h.Vm.OutputText.Should().Contain("[ERROR] Cmd not found").And.Contain("[ERROR] Other");
    }

    [Fact]
    public async Task Run_RunnerThrows_AppendsExceptionMarker_AndLogs()
    {
        var h = new Harness();
        h.Vm.InputText = "Throw-Me";
        h.Runner.Setup(r => r.RunAsync(It.IsAny<string>(), It.IsAny<IDictionary<string, object?>>(),
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("boom"));

        await h.Vm.RunCommand.ExecuteAsync(null);

        h.Vm.OutputText.Should().Contain("[EXCEPCION] boom");
        h.Vm.StatusMessage.Should().Be("Error: boom");
        h.Log.Entries.Should().Contain(e => e.Severity == LogSeverity.Error && e.Source == "PsConsole");
    }

    [Fact]
    public async Task Run_Cancelled_AppendsCancelMarker()
    {
        var h = new Harness();
        h.Vm.InputText = "Sleep";
        h.Runner.Setup(r => r.RunAsync(It.IsAny<string>(), It.IsAny<IDictionary<string, object?>>(),
            It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .ThrowsAsync(new OperationCanceledException());

        await h.Vm.RunCommand.ExecuteAsync(null);

        h.Vm.OutputText.Should().Contain("[CANCELADO]");
        h.Vm.StatusMessage.Should().Be("Cancelado.");
    }

    [Fact]
    public void Clear_EmptiesOutput_AndStatus()
    {
        var h = new Harness();
        h.Vm.OutputText = "stuff";

        h.Vm.ClearCommand.Execute(null);

        h.Vm.OutputText.Should().BeEmpty();
        h.Vm.StatusMessage.Should().Be("Salida limpiada.");
    }

    [Fact]
    public async Task History_PreservesOrder_AndDedupesConsecutiveDuplicates()
    {
        var h = new Harness();
        h.StubOk("ok");

        h.Vm.InputText = "A";
        await h.Vm.RunCommand.ExecuteAsync(null);
        h.Vm.InputText = "A"; // duplicate consecutive
        await h.Vm.RunCommand.ExecuteAsync(null);
        h.Vm.InputText = "B";
        await h.Vm.RunCommand.ExecuteAsync(null);

        h.Vm.History.Should().Equal("A", "B");
    }

    [Fact]
    public async Task HistoryPrev_NavigatesBackwards_FromLatest()
    {
        var h = new Harness();
        h.StubOk("ok");
        h.Vm.InputText = "A"; await h.Vm.RunCommand.ExecuteAsync(null);
        h.Vm.InputText = "B"; await h.Vm.RunCommand.ExecuteAsync(null);
        h.Vm.InputText = string.Empty;

        h.Vm.HistoryPrevCommand.Execute(null);
        h.Vm.InputText.Should().Be("B");

        h.Vm.HistoryPrevCommand.Execute(null);
        h.Vm.InputText.Should().Be("A");

        h.Vm.HistoryPrevCommand.Execute(null);
        h.Vm.InputText.Should().Be("A"); // stays at oldest
    }

    [Fact]
    public async Task HistoryNext_AdvancesForward_AndWrapsToEmpty()
    {
        var h = new Harness();
        h.StubOk("ok");
        h.Vm.InputText = "A"; await h.Vm.RunCommand.ExecuteAsync(null);
        h.Vm.InputText = "B"; await h.Vm.RunCommand.ExecuteAsync(null);
        h.Vm.InputText = string.Empty;

        h.Vm.HistoryPrevCommand.Execute(null);
        h.Vm.HistoryPrevCommand.Execute(null);
        h.Vm.InputText.Should().Be("A");

        h.Vm.HistoryNextCommand.Execute(null);
        h.Vm.InputText.Should().Be("B");

        h.Vm.HistoryNextCommand.Execute(null);
        h.Vm.InputText.Should().BeEmpty(); // wraps past end -> empty + index -1
    }
}
