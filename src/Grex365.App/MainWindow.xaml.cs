using System.Text;
using System.Windows;
using System.Windows.Input;
using Grex365.App.Services;
using Grex365.App.ViewModels;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
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

    // Ctrl+C copies the selected global-log rows (or all if none selected) as plain text.
    private void GlobalLogList_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.C || (Keyboard.Modifiers & ModifierKeys.Control) == 0)
        {
            return;
        }

        var rows = GlobalLogList.SelectedItems.Count > 0
            ? GlobalLogList.SelectedItems.Cast<object>()
            : GlobalLogList.Items.Cast<object>();

        var sb = new StringBuilder();
        foreach (var row in rows)
        {
            if (row is LogEntry l)
            {
                sb.Append(l.Timestamp.ToString("yyyy-MM-dd HH:mm:ss"))
                  .Append("  [").Append(l.Severity).Append("]  ")
                  .Append(l.Source).Append("  ")
                  .AppendLine(l.Message);
            }
            else if (row is not null)
            {
                sb.AppendLine(row.ToString());
            }
        }

        if (sb.Length > 0)
        {
            try { Clipboard.SetText(sb.ToString()); } catch { /* clipboard busy */ }
        }
        e.Handled = true;
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
