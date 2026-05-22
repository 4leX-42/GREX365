using System.Windows;
using Grex365.Core.Abstractions;

namespace Grex365.App.Services;

public sealed class WpfDialogService : IDialogService
{
    public Task<bool> ConfirmAsync(string message, string title, DialogIcon icon = DialogIcon.Question, CancellationToken cancellationToken = default)
    {
        var image = MapIcon(icon);
        var result = MessageBox.Show(message, title, MessageBoxButton.YesNo, image);
        return Task.FromResult(result == MessageBoxResult.Yes);
    }

    public Task ShowAsync(string message, string title, DialogIcon icon = DialogIcon.Info, CancellationToken cancellationToken = default)
    {
        var image = MapIcon(icon);
        MessageBox.Show(message, title, MessageBoxButton.OK, image);
        return Task.CompletedTask;
    }

    private static MessageBoxImage MapIcon(DialogIcon icon) => icon switch
    {
        DialogIcon.Info => MessageBoxImage.Information,
        DialogIcon.Warning => MessageBoxImage.Warning,
        DialogIcon.Error => MessageBoxImage.Error,
        DialogIcon.Question => MessageBoxImage.Question,
        _ => MessageBoxImage.None,
    };
}
