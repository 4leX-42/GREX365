using System.Collections.ObjectModel;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Models;
using Microsoft.Extensions.DependencyInjection;

namespace Grex365.App.ViewModels;

public sealed class NavigationItem
{
    public NavigationItem(string title, string glyph, Type viewModelType)
    {
        Title = title;
        Glyph = glyph;
        ViewModelType = viewModelType;
    }

    public string Title { get; }
    public string Glyph { get; }
    public Type ViewModelType { get; }
}

public sealed partial class MainViewModel : ObservableObject
{
    private readonly IServiceProvider _services;

    [ObservableProperty] private NavigationItem? _selectedNavigation;
    [ObservableProperty] private ObservableObject? _currentPage;

    public MainViewModel(IUiLogSink uiLog, IServiceProvider services)
    {
        LogEntries = uiLog.Entries;
        _services = services;

        NavigationItems = new ObservableCollection<NavigationItem>
        {
            new("Dashboard",  "", typeof(DashboardViewModel)),
            new("Conexion",   "", typeof(ConnectViewModel)),
            new("Grupos",     "", typeof(GroupsViewModel)),
            new("Buzones",    "", typeof(SharedMailboxViewModel)),
            new("Auditoria",  "", typeof(AuditViewModel)),
        };

        SelectedNavigation = NavigationItems[0];
    }

    public ObservableCollection<NavigationItem> NavigationItems { get; }

    public ObservableCollection<LogEntry> LogEntries { get; }

    partial void OnSelectedNavigationChanged(NavigationItem? value)
    {
        if (value is null)
        {
            CurrentPage = null;
            return;
        }
        CurrentPage = (ObservableObject)_services.GetRequiredService(value.ViewModelType);
    }

    [RelayCommand]
    private void OpenSettings()
    {
        var window = _services.GetRequiredService<SettingsWindow>();
        window.Owner = Application.Current?.MainWindow;
        window.ShowDialog();
    }
}
