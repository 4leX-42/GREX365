using System.Windows;
using Grex365.Core.Models;
using Wpf.Ui;
using Wpf.Ui.Controls;

namespace Grex365.App.Services;

public sealed class WpfUiNotifier : INotifier
{
    private readonly SnackbarService _service = new();
    private SnackbarPresenter? _presenter;

    public void AttachPresenter(SnackbarPresenter presenter)
    {
        _presenter = presenter;
        _service.SetSnackbarPresenter(presenter);
    }

    public void Notify(string title, string message, LogSeverity severity)
    {
        if (_presenter is null)
        {
            return;
        }

        var dispatcher = Application.Current?.Dispatcher;
        if (dispatcher is not null && !dispatcher.CheckAccess())
        {
            dispatcher.Invoke(() => Show(title, message, severity));
        }
        else
        {
            Show(title, message, severity);
        }
    }

    private void Show(string title, string message, LogSeverity severity)
    {
        var (appearance, icon, duration) = severity switch
        {
            LogSeverity.Error => (ControlAppearance.Danger, SymbolRegular.ErrorCircle24, TimeSpan.FromSeconds(8)),
            LogSeverity.Warning => (ControlAppearance.Caution, SymbolRegular.Warning24, TimeSpan.FromSeconds(6)),
            LogSeverity.Ok => (ControlAppearance.Success, SymbolRegular.Checkmark24, TimeSpan.FromSeconds(4)),
            _ => (ControlAppearance.Info, SymbolRegular.Info24, TimeSpan.FromSeconds(4))
        };

        var truncated = message.Length > 280 ? message[..277] + "..." : message;
        _service.Show(title, truncated, appearance, new SymbolIcon { Symbol = icon }, duration);
    }
}
