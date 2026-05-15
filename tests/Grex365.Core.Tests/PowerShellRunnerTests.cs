using FluentAssertions;
using Grex365.Core.Models;
using Grex365.PowerShell;

namespace Grex365.Core.Tests;

[Collection("Runspace")]
public class PowerShellRunnerTests : IAsyncLifetime
{
    private RunspacePoolHost _host = null!;
    private PowerShellRunner _runner = null!;

    public Task InitializeAsync()
    {
        _host = new RunspacePoolHost(minRunspaces: 1, maxRunspaces: 2);
        _runner = new PowerShellRunner(_host);
        return Task.CompletedTask;
    }

    public async Task DisposeAsync()
    {
        await _runner.DisposeAsync();
        _host.Dispose();
    }

    [Fact]
    public async Task RunAsync_HappyPath_Returns_Output()
    {
        var result = await _runner.RunAsync("2 + 3");

        result.Success.Should().BeTrue();
        result.Output.Should().HaveCount(1);
        result.Output[0].Should().Be(5);
        result.Errors.Should().BeEmpty();
    }

    [Fact]
    public async Task RunAsync_With_Parameters_Binds_Correctly()
    {
        var result = await _runner.RunAsync(
            "param($x,$y) $x * $y",
            new Dictionary<string, object?> { ["x"] = 6, ["y"] = 7 });

        result.Success.Should().BeTrue();
        result.Output[0].Should().Be(42);
    }

    [Fact]
    public async Task RunAsync_Captures_Error_Stream()
    {
        var result = await _runner.RunAsync("Write-Error 'boom'");

        result.Success.Should().BeFalse();
        result.Errors.Should().NotBeEmpty();
        result.Errors[0].Should().Contain("boom");
    }

    [Fact]
    public async Task RunAsync_Forwards_Information_Stream_To_Progress()
    {
        var entries = new List<LogEntry>();
        var progress = new Progress<LogEntry>(e => entries.Add(e));

        await _runner.RunAsync("Write-Information 'hola' -InformationAction Continue", progress: progress);

        await Task.Delay(50);

        entries.Should().Contain(e => e.Severity == LogSeverity.Info && e.Message.Contains("hola"));
    }

    [Fact]
    public async Task RunAsync_Forwards_Warning_Stream_To_Progress()
    {
        var entries = new List<LogEntry>();
        var progress = new Progress<LogEntry>(e => entries.Add(e));

        await _runner.RunAsync("$WarningPreference='Continue'; Write-Warning 'cuidado'", progress: progress);

        await Task.Delay(50);

        entries.Should().Contain(e => e.Severity == LogSeverity.Warning && e.Message.Contains("cuidado"));
    }

    [Fact]
    public async Task RunAsync_Cancellation_Stops_Script()
    {
        using var cts = new CancellationTokenSource();
        cts.CancelAfter(TimeSpan.FromMilliseconds(200));

        Func<Task> act = () => _runner.RunAsync("Start-Sleep -Seconds 30", cancellationToken: cts.Token);

        await act.Should().ThrowAsync<OperationCanceledException>();
    }

    [Fact]
    public async Task RunAsync_Concurrent_Calls_Do_Not_Deadlock()
    {
        var tasks = Enumerable.Range(0, 4)
            .Select(_ => _runner.RunAsync("Start-Sleep -Milliseconds 100; 'ok'"))
            .ToArray();

        var results = await Task.WhenAll(tasks);

        results.Should().AllSatisfy(r =>
        {
            r.Success.Should().BeTrue();
            r.Output.Should().HaveCount(1);
        });
    }
}
