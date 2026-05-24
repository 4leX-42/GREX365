using System.Collections.ObjectModel;
using System.Windows;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Serilog;

namespace Grex365.App.Services;

public sealed class UiLogSink : IUiLogSink
{
    private const int MaxEntries = 5000;
    private readonly Progress<LogEntry> _progress;
    private readonly INotifier? _notifier;
    private readonly IAuditLog? _audit;
    private readonly ITelemetry? _telemetry;
    private readonly string _actor;

    // Sources that should never emit toast notifications. Status bar + log panel
    // are the canonical surfaces for connection/reconnection state per user
    // preference ("solo status bar silencioso"). Adding a source here keeps it
    // in Serilog + audit + telemetry but suppresses the snackbar toast.
    private static readonly HashSet<string> SilentSources = new(StringComparer.OrdinalIgnoreCase)
    {
        "AutoConnect",
        "Connect",
        "ConnectionMonitor",
        "TenantLock",
        "Settings",
        "EXO",
        "Graph",
    };

    public UiLogSink(INotifier? notifier = null, IAuditLog? audit = null, ITelemetry? telemetry = null)
    {
        _notifier = notifier;
        _audit = audit;
        _telemetry = telemetry;
        _actor = Environment.UserName;
        Entries = new ObservableCollection<LogEntry>();
        _progress = new Progress<LogEntry>(OnEntry);
    }

    private bool ShouldNotify(LogEntry entry)
    {
        if (_notifier is null) return false;
        // Always suppress for Info/Debug (covered by log panel only).
        if (entry.Severity is LogSeverity.Info or LogSeverity.Debug) return false;
        // Silent sources never emit toasts regardless of severity.
        return !SilentSources.Contains(entry.Source);
    }

    public ObservableCollection<LogEntry> Entries { get; }

    public IProgress<LogEntry> Progress => _progress;

    public void Clear()
    {
        Application.Current.Dispatcher.Invoke(Entries.Clear);
    }

    private void OnEntry(LogEntry entry)
    {
        // Progress<T> already marshals to the captured SynchronizationContext (UI thread when constructed there).
        Entries.Add(entry);
        if (Entries.Count > MaxEntries)
        {
            Entries.RemoveAt(0);
        }

        var shouldNotify = ShouldNotify(entry);
        switch (entry.Severity)
        {
            case LogSeverity.Error:
                Log.Error(entry.Exception, "[{Source}] {Message}", entry.Source, entry.Message);
                if (shouldNotify) _notifier!.Notify(entry.Source, entry.Message, entry.Severity);
                FireAuditAsync(entry, "ERROR");
                TryTrack(entry, "ERROR");
                break;
            case LogSeverity.Warning:
                Log.Warning("[{Source}] {Message}", entry.Source, entry.Message);
                if (shouldNotify) _notifier!.Notify(entry.Source, entry.Message, entry.Severity);
                FireAuditAsync(entry, "WARN");
                TryTrack(entry, "WARN");
                break;
            case LogSeverity.Ok:
                Log.Information("[{Source}] OK · {Message}", entry.Source, entry.Message);
                // Ok severity never emits toast (success is surfaced via status bar / log panel).
                FireAuditAsync(entry, "OK");
                TryTrack(entry, "OK");
                break;
            case LogSeverity.Info:
                Log.Information("[{Source}] {Message}", entry.Source, entry.Message);
                break;
            default:
                Log.Debug("[{Source}] {Message}", entry.Source, entry.Message);
                break;
        }
    }

    private void TryTrack(LogEntry entry, string outcome)
    {
        if (_telemetry is null || !_telemetry.IsEnabled)
        {
            return;
        }
        try
        {
            var props = new Dictionary<string, string>
            {
                ["source"] = entry.Source,
                ["outcome"] = outcome,
                ["message"] = entry.Message,
                ["actor"] = _actor,
            };
            if (entry.Exception is not null)
            {
                _telemetry.TrackException(entry.Exception, props);
            }
            else
            {
                _telemetry.TrackEvent(entry.Source, props);
            }
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Telemetry track failed");
        }
    }

    private void FireAuditAsync(LogEntry entry, string outcome)
    {
        if (_audit is null)
        {
            return;
        }
        var record = new AuditRecord(
            Timestamp: entry.Timestamp,
            Actor: _actor,
            Source: entry.Source,
            Outcome: outcome,
            Message: entry.Message,
            Detail: entry.Exception?.ToString());
        _ = Task.Run(async () =>
        {
            try
            {
                await _audit.WriteAsync(record).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Audit log write failed");
            }
        });
    }
}
