using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Grex365.Core.Abstractions;
using Wpf.Ui.Controls;

namespace Grex365.App.Services;

/// <summary>
/// IDialogService sobre Wpf.Ui.Controls.MessageBox (FluentWindow) en lugar del
/// MessageBox nativo de Win32: respeta el tema Mica/dark, botón primario con
/// Appearance Danger en confirmaciones destructivas (Warning/Error) y textos L10n.
/// </summary>
public sealed class WpfDialogService : IDialogService
{
    public Task<bool> ConfirmAsync(string message, string title, DialogIcon icon = DialogIcon.Question, CancellationToken cancellationToken = default)
        => OnUiThreadAsync(async () =>
        {
            var box = BuildBox(message, title, icon);
            box.PrimaryButtonText = L10n.Get("Dialog.Yes");
            box.CloseButtonText = L10n.Get("Dialog.No");
            box.PrimaryButtonAppearance = icon is DialogIcon.Warning or DialogIcon.Error
                ? ControlAppearance.Danger
                : ControlAppearance.Primary;
            var result = await box.ShowDialogAsync(true, cancellationToken);
            return result == Wpf.Ui.Controls.MessageBoxResult.Primary;
        });

    public Task ShowAsync(string message, string title, DialogIcon icon = DialogIcon.Info, CancellationToken cancellationToken = default)
        => OnUiThreadAsync<object?>(async () =>
        {
            var box = BuildBox(message, title, icon);
            box.CloseButtonText = L10n.Get("Dialog.Ok");
            await box.ShowDialogAsync(true, cancellationToken);
            return null;
        });

    private static Task<T> OnUiThreadAsync<T>(Func<Task<T>> action)
    {
        var dispatcher = Application.Current?.Dispatcher;
        if (dispatcher is null || dispatcher.CheckAccess())
        {
            return action();
        }
        return dispatcher.InvokeAsync(action).Task.Unwrap();
    }

    private static Wpf.Ui.Controls.MessageBox BuildBox(string message, string title, DialogIcon icon)
    {
        var (symbol, brushKey) = icon switch
        {
            DialogIcon.Warning => (SymbolRegular.Warning24, "BrushSemanticWarn"),
            DialogIcon.Error => (SymbolRegular.ErrorCircle24, "BrushSemanticError"),
            DialogIcon.Question => (SymbolRegular.QuestionCircle24, "BrushSemanticInfo"),
            _ => (SymbolRegular.Info24, "BrushSemanticInfo"),
        };

        var panel = new Grid { Margin = new Thickness(0, 6, 0, 6) };
        panel.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        panel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        var iconControl = new SymbolIcon
        {
            Symbol = symbol,
            FontSize = 30,
            VerticalAlignment = VerticalAlignment.Top,
            Margin = new Thickness(0, 0, 14, 0),
        };
        if (Application.Current?.TryFindResource(brushKey) is Brush brush)
        {
            iconControl.Foreground = brush;
        }
        Grid.SetColumn(iconControl, 0);
        panel.Children.Add(iconControl);

        var text = new System.Windows.Controls.TextBlock
        {
            Text = message,
            TextWrapping = TextWrapping.Wrap,
            MaxWidth = 420,
            VerticalAlignment = VerticalAlignment.Center,
        };
        Grid.SetColumn(text, 1);
        panel.Children.Add(text);

        return new Wpf.Ui.Controls.MessageBox
        {
            Title = title,
            Content = panel,
            MinWidth = 380,
            MaxWidth = 560,
        };
    }
}
