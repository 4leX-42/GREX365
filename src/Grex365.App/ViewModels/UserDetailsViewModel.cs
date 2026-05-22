using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.App.ViewModels;

public sealed partial class UserDetailsViewModel : ObservableObject
{
    private readonly IUsersService _users;
    private readonly IUserDetailsHost _host;
    private readonly IUiLogSink _log;
    private CancellationTokenSource? _cts;

    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private string _statusMessage = string.Empty;
    [ObservableProperty] private UserSummary? _user;
    [ObservableProperty] private string? _userId;
    [ObservableProperty] private bool _hasUser;

    public ObservableCollection<GroupSummary> Memberships { get; } = new();

    public UserDetailsViewModel(IUsersService users, IUserDetailsHost host, IUiLogSink log)
    {
        _users = users;
        _host = host;
        _log = log;
        _host.OpenRequested += (_, id) => _ = LoadAsync(id);
        _host.CloseRequested += (_, _) => Reset();
    }

    private void Reset()
    {
        _cts?.Cancel();
        User = null;
        UserId = null;
        HasUser = false;
        Memberships.Clear();
        StatusMessage = string.Empty;
    }

    private async Task LoadAsync(string userId)
    {
        _cts?.Cancel();
        _cts = new CancellationTokenSource();
        var token = _cts.Token;

        UserId = userId;
        HasUser = false;
        User = null;
        Memberships.Clear();
        IsBusy = true;
        StatusMessage = "Cargando perfil...";
        try
        {
            var user = await _users.GetByIdAsync(userId, token).ConfigureAwait(true);
            if (user is null)
            {
                StatusMessage = "Usuario no encontrado.";
                return;
            }
            User = user;
            HasUser = true;
            StatusMessage = "Cargando grupos...";
            try
            {
                var groups = await _users.GetGroupMembershipsAsync(userId, token).ConfigureAwait(true);
                foreach (var g in groups.OrderBy(g => g.DisplayName, StringComparer.OrdinalIgnoreCase))
                {
                    Memberships.Add(g);
                }
            }
            catch (Exception ex)
            {
                _log.Progress.Report(LogEntry.Warn("UserDetails", $"Grupos: {ex.Message}"));
            }
            StatusMessage = $"{User.DisplayName} · {Memberships.Count} grupos · {User.AssignedLicenseCount} licencias";
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("UserDetails", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private void Close() => _host.RequestClose();

    [RelayCommand]
    private async Task ToggleAccountAsync()
    {
        if (User is null || string.IsNullOrEmpty(User.Id)) return;
        var newState = !User.AccountEnabled;
        var confirm = System.Windows.MessageBox.Show(
            (newState ? "Habilitar" : "Deshabilitar") + $" la cuenta {User.UserPrincipalName}?",
            "Confirmar",
            System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Question);
        if (confirm != System.Windows.MessageBoxResult.Yes) return;

        IsBusy = true;
        StatusMessage = newState ? "Habilitando..." : "Deshabilitando...";
        try
        {
            await _users.SetAccountEnabledAsync(User.Id, newState, _log.Progress).ConfigureAwait(true);
            await LoadAsync(User.Id).ConfigureAwait(true);
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("UserDetails", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private async Task RemoveAllLicensesAsync()
    {
        if (User is null || string.IsNullOrEmpty(User.Id)) return;
        var confirm = System.Windows.MessageBox.Show(
            $"Quitar TODAS las licencias asignadas a {User.UserPrincipalName}?",
            "Confirmar",
            System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Warning);
        if (confirm != System.Windows.MessageBoxResult.Yes) return;

        IsBusy = true;
        StatusMessage = "Quitando licencias...";
        try
        {
            await _users.RemoveAllLicensesAsync(User.Id, _log.Progress).ConfigureAwait(true);
            await LoadAsync(User.Id).ConfigureAwait(true);
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("UserDetails", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
        }
    }
}
