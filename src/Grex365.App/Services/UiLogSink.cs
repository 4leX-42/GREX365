using System.Collections.ObjectModel;
using System.Windows;
using Grex365.Core.Models;
using Serilog;

namespace Grex365.App.Services;

public sealed class UiLogSink : IUiLogSink
{
    private const int MaxEntries = 5000;
    private readonly Progress<LogEntry> _progress;
    private readonly INotifier? _notifier;

    public UiLogSink(INotifier? notifier = null)
    {
        _notifier = notifier;
        Entries = new ObservableCollection<LogEntry>();
        _progress = new Progress<LogEntry>(OnEntry);
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

        switch (entry.Severity)
        {
            case LogSeverity.Error:
                Log.Error(entry.Exception, "[{Source}] {Message}", entry.Source, entry.Message);
                _notifier?.Notify(entry.Source, entry.Message, entry.Severity);
                break;
            case LogSeverity.Warning:
                Log.Warning("[{Source}] {Message}", entry.Source, entry.Message);
                _notifier?.Notify(entry.Source, entry.Message, entry.Severity);
                break;
            case LogSeverity.Ok:
                Log.Information("[{Source}] OK · {Message}", entry.Source, entry.Message);
                _notifier?.Notify(entry.Source, entry.Message, entry.Severity);
                break;
            case LogSeverity.Info:
                Log.Information("[{Source}] {Message}", entry.Source, entry.Message);
                break;
            default:
                Log.Debug("[{Source}] {Message}", entry.Source, entry.Message);
                break;
        }
    }
}
