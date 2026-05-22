namespace Grex365.Core.Abstractions;

public enum DialogIcon
{
    None,
    Info,
    Warning,
    Error,
    Question
}

public interface IDialogService
{
    Task<bool> ConfirmAsync(string message, string title, DialogIcon icon = DialogIcon.Question, CancellationToken cancellationToken = default);

    Task ShowAsync(string message, string title, DialogIcon icon = DialogIcon.Info, CancellationToken cancellationToken = default);
}
