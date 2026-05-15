using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using Grex365.App.Services;
using Grex365.Core.Models;

namespace Grex365.App.ViewModels;

public sealed partial class MainViewModel : ObservableObject
{
    public MainViewModel(ConnectViewModel connectViewModel, IUiLogSink uiLog)
    {
        Connect = connectViewModel;
        LogEntries = uiLog.Entries;
    }

    public ConnectViewModel Connect { get; }

    public ObservableCollection<LogEntry> LogEntries { get; }
}
