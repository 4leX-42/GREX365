using System.ComponentModel;

namespace Grex365.Core.Abstractions;

public interface IUserDetailsHost : INotifyPropertyChanged
{
    bool IsOpen { get; }
    string? CurrentUserId { get; }

    event EventHandler<string>? OpenRequested;
    event EventHandler? CloseRequested;

    void RequestOpen(string userId);
    void RequestClose();
}
