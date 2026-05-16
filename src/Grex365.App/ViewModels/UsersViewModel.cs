using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.App.ViewModels;

public sealed partial class UsersViewModel : ObservableObject
{
    private readonly IUsersService _users;
    private readonly IUiLogSink _log;
    private CancellationTokenSource? _cts;

    [ObservableProperty] private string _searchQuery = string.Empty;
    [ObservableProperty] private UserSummary? _selectedUser;
    [ObservableProperty] private string _statusMessage = string.Empty;
    [ObservableProperty] private bool _isBusy;

    public ObservableCollection<UserSummary> Users { get; } = new();
    public ObservableCollection<GroupSummary> Memberships { get; } = new();

    public UsersViewModel(IUsersService users, IUiLogSink log)
    {
        _users = users;
        _log = log;
    }

    partial void OnSelectedUserChanged(UserSummary? value)
    {
        Memberships.Clear();
        if (value is null)
        {
            return;
        }
        _ = LoadMembershipsAsync(value.Id);
    }

    [RelayCommand]
    private async Task SearchAsync()
    {
        EnsureToken();
        IsBusy = true;
        CancelCommand.NotifyCanExecuteChanged();
        StatusMessage = "Buscando usuarios...";
        try
        {
            Users.Clear();
            var found = await _users.SearchAsync(SearchQuery, _cts!.Token).ConfigureAwait(true);
            foreach (var u in found)
            {
                Users.Add(u);
            }
            StatusMessage = $"{found.Count} usuarios.";
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("Users", ex.Message, ex));
        }
        finally
        {
            DisposeToken();
        }
    }

    private async Task LoadMembershipsAsync(string userId)
    {
        EnsureToken();
        IsBusy = true;
        CancelCommand.NotifyCanExecuteChanged();
        try
        {
            var groups = await _users.GetGroupMembershipsAsync(userId, _cts!.Token).ConfigureAwait(true);
            foreach (var g in groups)
            {
                Memberships.Add(g);
            }
            StatusMessage = $"{groups.Count} membresías.";
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("Users", ex.Message, ex));
        }
        finally
        {
            DisposeToken();
        }
    }

    [RelayCommand]
    private async Task DisableAsync() => await SetEnabledAsync(false).ConfigureAwait(true);

    [RelayCommand]
    private async Task EnableAsync() => await SetEnabledAsync(true).ConfigureAwait(true);

    private async Task SetEnabledAsync(bool enabled)
    {
        if (SelectedUser is null)
        {
            StatusMessage = "Selecciona un usuario.";
            return;
        }
        EnsureToken();
        IsBusy = true;
        StatusMessage = enabled ? "Habilitando..." : "Deshabilitando...";
        try
        {
            await _users.SetAccountEnabledAsync(SelectedUser.Id, enabled, _log.Progress, _cts!.Token).ConfigureAwait(true);
            StatusMessage = enabled ? "Habilitado." : "Deshabilitado.";
            // refresh user
            var refreshed = await _users.GetByIdAsync(SelectedUser.Id, _cts.Token).ConfigureAwait(true);
            if (refreshed is not null)
            {
                var idx = Users.IndexOf(SelectedUser);
                if (idx >= 0)
                {
                    Users[idx] = refreshed;
                    SelectedUser = refreshed;
                }
            }
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("Users", ex.Message, ex));
        }
        finally
        {
            DisposeToken();
        }
    }

    [RelayCommand]
    private async Task RemoveLicensesAsync()
    {
        if (SelectedUser is null)
        {
            StatusMessage = "Selecciona un usuario.";
            return;
        }
        EnsureToken();
        IsBusy = true;
        StatusMessage = "Quitando licencias...";
        try
        {
            await _users.RemoveAllLicensesAsync(SelectedUser.Id, _log.Progress, _cts!.Token).ConfigureAwait(true);
            var refreshed = await _users.GetByIdAsync(SelectedUser.Id, _cts.Token).ConfigureAwait(true);
            if (refreshed is not null)
            {
                var idx = Users.IndexOf(SelectedUser);
                if (idx >= 0)
                {
                    Users[idx] = refreshed;
                    SelectedUser = refreshed;
                }
            }
            StatusMessage = "Licencias retiradas.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("Users", ex.Message, ex));
        }
        finally
        {
            DisposeToken();
        }
    }

    [RelayCommand(CanExecute = nameof(CanCancel))]
    private void Cancel() => _cts?.Cancel();

    private bool CanCancel() => IsBusy;

    private void EnsureToken()
    {
        _cts?.Dispose();
        _cts = new CancellationTokenSource();
    }

    private void DisposeToken()
    {
        IsBusy = false;
        _cts?.Dispose();
        _cts = null;
        CancelCommand.NotifyCanExecuteChanged();
    }
}
