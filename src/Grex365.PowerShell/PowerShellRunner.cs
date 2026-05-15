using System.Collections.ObjectModel;
using System.Management.Automation;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.PowerShell;

public sealed class PowerShellRunner : IPowerShellRunner
{
    private readonly RunspacePoolHost _host;
    private bool _disposed;

    public PowerShellRunner(RunspacePoolHost host)
    {
        _host = host;
    }

    public async Task<PowerShellResult> RunAsync(
        string script,
        IDictionary<string, object?>? parameters = null,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        using var ps = System.Management.Automation.PowerShell.Create();
        ps.RunspacePool = _host.Pool;
        ps.AddScript(script);

        if (parameters is not null)
        {
            foreach (var kvp in parameters)
            {
                ps.AddParameter(kvp.Key, kvp.Value);
            }
        }

        SubscribeStreams(ps, progress);

        var output = new PSDataCollection<PSObject>();
        var errors = new List<string>();

        ps.Streams.Error.DataAdded += (s, e) =>
        {
            var rec = ((PSDataCollection<ErrorRecord>)s!)[e.Index];
            errors.Add(rec.ToString());
            progress?.Report(LogEntry.Error("PS", rec.ToString(), rec.Exception));
        };

        await using var registration = cancellationToken.Register(() =>
        {
            try { ps.Stop(); } catch { /* ignore */ }
        }).ConfigureAwait(false);

        var invokeResult = ps.BeginInvoke<PSObject, PSObject>(input: null, output);
        try
        {
            await Task.Factory.FromAsync(invokeResult, ps.EndInvoke).ConfigureAwait(false);
        }
        catch (PipelineStoppedException) when (cancellationToken.IsCancellationRequested)
        {
            progress?.Report(LogEntry.Warn("PS", "Ejecución cancelada por el usuario."));
            throw new OperationCanceledException(cancellationToken);
        }

        var outList = new List<object?>(output.Count);
        foreach (var item in output)
        {
            outList.Add(item?.BaseObject);
        }

        return new PowerShellResult(
            Success: errors.Count == 0 && !ps.HadErrors,
            Output: new ReadOnlyCollection<object?>(outList),
            Errors: new ReadOnlyCollection<string>(errors));
    }

    private static void SubscribeStreams(System.Management.Automation.PowerShell ps, IProgress<LogEntry>? progress)
    {
        if (progress is null)
        {
            return;
        }

        ps.Streams.Information.DataAdded += (s, e) =>
        {
            var rec = ((PSDataCollection<InformationRecord>)s!)[e.Index];
            progress.Report(LogEntry.Info("PS", rec.MessageData?.ToString() ?? string.Empty));
        };
        ps.Streams.Warning.DataAdded += (s, e) =>
        {
            var rec = ((PSDataCollection<WarningRecord>)s!)[e.Index];
            progress.Report(LogEntry.Warn("PS", rec.Message));
        };
        ps.Streams.Verbose.DataAdded += (s, e) =>
        {
            var rec = ((PSDataCollection<VerboseRecord>)s!)[e.Index];
            progress.Report(LogEntry.Debug("PS", rec.Message));
        };
        ps.Streams.Debug.DataAdded += (s, e) =>
        {
            var rec = ((PSDataCollection<DebugRecord>)s!)[e.Index];
            progress.Report(LogEntry.Debug("PS", rec.Message));
        };
    }

    public ValueTask DisposeAsync()
    {
        _disposed = true;
        return ValueTask.CompletedTask;
    }
}
