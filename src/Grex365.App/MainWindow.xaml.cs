using System.Windows;
using Grex365.App.Services;
using Grex365.App.ViewModels;
using Grex365.Core.Abstractions;
using Microsoft.Extensions.DependencyInjection;
using Wpf.Ui.Controls;

namespace Grex365.App;

public partial class MainWindow : FluentWindow
{
    private readonly IPreferencesStore _prefs;

    public MainWindow(MainViewModel viewModel, WpfUiNotifier notifier, IPreferencesStore prefs)
    {
        _prefs = prefs;
        InitializeComponent();
        DataContext = viewModel;
        Loaded += (_, _) =>
        {
            notifier.AttachPresenter(RootSnackbarPresenter);
            RestoreWindowState();
        };
        Closing += OnMainWindowClosing;
    }

    private void OnMainWindowClosing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        if (Application.Current is App app && !app.IsExplicitExitRequested)
        {
            // X click on main window: hide to tray, keep app running.
            e.Cancel = true;
            SaveWindowState();
            Hide();
            return;
        }
        SaveWindowState();
    }

    private async void RestoreWindowState()
    {
        try
        {
            var prefs = await _prefs.LoadAsync().ConfigureAwait(true);
            if (prefs.WindowWidth is { } w && w >= MinWidth)
            {
                Width = w;
            }
            if (prefs.WindowHeight is { } h && h >= MinHeight)
            {
                Height = h;
            }
            if (prefs.WindowLeft is { } l && IsOnScreen(l, prefs.WindowTop ?? Top, Width, Height))
            {
                Left = l;
            }
            if (prefs.WindowTop is { } t && IsOnScreen(prefs.WindowLeft ?? Left, t, Width, Height))
            {
                Top = t;
            }
            if (prefs.WindowMaximized)
            {
                WindowState = WindowState.Maximized;
            }
        }
        catch
        {
            // Non-critical; ignore and use defaults.
        }
    }

    private async void SaveWindowState()
    {
        try
        {
            var prefs = await _prefs.LoadAsync().ConfigureAwait(false);
            var isMax = WindowState == WindowState.Maximized;
            prefs.WindowMaximized = isMax;
            if (!isMax)
            {
                prefs.WindowWidth = Width;
                prefs.WindowHeight = Height;
                prefs.WindowLeft = Left;
                prefs.WindowTop = Top;
            }
            await _prefs.SaveAsync(prefs).ConfigureAwait(false);
        }
        catch
        {
            // Non-critical; ignore.
        }
    }

    private static bool IsOnScreen(double left, double top, double width, double height)
    {
        var virt = new VirtualScreenBounds(
            Left: System.Windows.SystemParameters.VirtualScreenLeft,
            Top: System.Windows.SystemParameters.VirtualScreenTop,
            Width: System.Windows.SystemParameters.VirtualScreenWidth,
            Height: System.Windows.SystemParameters.VirtualScreenHeight);
        return WindowPlacementGuard.IsOnScreen(left, top, width, height, virt);
    }
}
