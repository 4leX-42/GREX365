using System.Collections.ObjectModel;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Models;
using Microsoft.Extensions.DependencyInjection;

namespace Grex365.App.ViewModels;

public sealed partial class MainViewModel : ObservableObject
{
    private readonly IServiceProvider _services;

    public MainViewModel(ConnectViewModel connectViewModel, IUiLogSink uiLog, IServiceProvider services)
    {
        Connect = connectViewModel;
        LogEntries = uiLog.Entries;
        _services = services;
    }

    public ConnectViewModel Connect { get; }

    public ObservableCollection<LogEntry> LogEntries { get; }

    [RelayCommand]
    private void OpenSettings()
    {
        var window = _services.GetRequiredService<SettingsWindow>();
        window.Owner = Application.Current?.MainWindow;
        window.ShowDialog();
    }
}
