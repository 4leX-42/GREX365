using System.ComponentModel;
using Grex365.Core.Abstractions;

namespace Grex365.Core.Users;

public sealed class UserDetailsHost : IUserDetailsHost
{
    private bool _isOpen;
    private string? _currentUserId;

    public bool IsOpen
    {
        get => _isOpen;
        private set
        {
            if (_isOpen == value) return;
            _isOpen = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsOpen)));
        }
    }

    public string? CurrentUserId
    {
        get => _currentUserId;
        private set
        {
            if (_currentUserId == value) return;
            _currentUserId = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(CurrentUserId)));
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    public event EventHandler<string>? OpenRequested;
    public event EventHandler? CloseRequested;

    public void RequestOpen(string userId)
    {
        if (string.IsNullOrWhiteSpace(userId)) return;
        CurrentUserId = userId;
        IsOpen = true;
        OpenRequested?.Invoke(this, userId);
    }

    public void RequestClose()
    {
        IsOpen = false;
        CurrentUserId = null;
        CloseRequested?.Invoke(this, EventArgs.Empty);
    }
}
