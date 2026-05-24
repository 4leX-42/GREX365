using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.App.ViewModels;

public sealed record AssignedLicenseRow(
    Guid SkuId,
    string SkuPartNumber,
    string FriendlyName,
    string CategoryLabel);

public sealed partial class UserDetailsViewModel : ObservableObject
{
    private readonly IUsersService _users;
    private readonly IUserDetailsHost _host;
    private readonly IUiLogSink _log;
    private readonly IDialogService _dialogs;
    private readonly IClipboardService _clipboard;
    private CancellationTokenSource? _cts;
    private List<SkuInfo> _allSkus = new();

    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private string _statusMessage = string.Empty;
    [ObservableProperty] private UserSummary? _user;
    [ObservableProperty] private string? _userId;
    [ObservableProperty] private bool _hasUser;
    [ObservableProperty] private SkuInfo? _selectedSkuToAdd;

    public ObservableCollection<GroupSummary> Memberships { get; } = new();
    public ObservableCollection<AssignedLicenseRow> AssignedLicenses { get; } = new();
    public ObservableCollection<SkuInfo> AssignableSkus { get; } = new();

    public UserDetailsViewModel(
        IUsersService users,
        IUserDetailsHost host,
        IUiLogSink log,
        IDialogService dialogs,
        IClipboardService clipboard)
    {
        _users = users;
        _host = host;
        _log = log;
        _dialogs = dialogs;
        _clipboard = clipboard;
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
        AssignedLicenses.Clear();
        AssignableSkus.Clear();
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
        AssignedLicenses.Clear();
        AssignableSkus.Clear();
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

            StatusMessage = "Cargando licencias...";
            try
            {
                if (_allSkus.Count == 0)
                {
                    _allSkus = (await _users.ListSkusAsync(token).ConfigureAwait(true)).ToList();
                }
                var assigned = await _users.GetAssignedLicensesAsync(userId, token).ConfigureAwait(true);
                var assignedSet = assigned.ToHashSet();

                var assignedRows = assigned
                    .Select(skuId =>
                    {
                        var sku = _allSkus.FirstOrDefault(s => s.SkuId == skuId);
                        var partNumber = sku?.SkuPartNumber ?? skuId.ToString();
                        var info = SkuCatalog.Resolve(partNumber);
                        return (Row: new AssignedLicenseRow(
                                SkuId: skuId,
                                SkuPartNumber: partNumber,
                                FriendlyName: info.FriendlyName,
                                CategoryLabel: SkuCatalog.CategoryLabel(info.Category)),
                            Priority: info.Priority);
                    })
                    .OrderBy(t => t.Priority)
                    .ThenBy(t => t.Row.FriendlyName, StringComparer.OrdinalIgnoreCase)
                    .Select(t => t.Row);
                foreach (var row in assignedRows)
                {
                    AssignedLicenses.Add(row);
                }

                foreach (var sku in _allSkus
                    .Where(s => !assignedSet.Contains(s.SkuId) && s.Available > 0)
                    .OrderBy(s => SkuCatalog.Resolve(s.SkuPartNumber).Priority))
                {
                    AssignableSkus.Add(sku);
                }
            }
            catch (Exception ex)
            {
                _log.Progress.Report(LogEntry.Warn("UserDetails", $"Licencias: {ex.Message}"));
            }

            StatusMessage = $"{User.DisplayName} · {Memberships.Count} grupos · {AssignedLicenses.Count} licencias";
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
        var ok = await _dialogs.ConfirmAsync(
            (newState ? "Habilitar" : "Deshabilitar") + $" la cuenta {User.UserPrincipalName}?",
            "Confirmar").ConfigureAwait(true);
        if (!ok) return;

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
    private async Task RemoveLicenseAsync(AssignedLicenseRow? row)
    {
        if (row is null || User is null || string.IsNullOrEmpty(User.Id)) return;
        var ok = await _dialogs.ConfirmAsync(
            $"Quitar la licencia '{row.FriendlyName}' ({row.SkuPartNumber}) de {User.UserPrincipalName}?",
            "Confirmar").ConfigureAwait(true);
        if (!ok) return;

        IsBusy = true;
        StatusMessage = $"Quitando {row.FriendlyName}...";
        try
        {
            await _users.RemoveLicenseAsync(User.Id, row.SkuId, _log.Progress).ConfigureAwait(true);
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
    private async Task AssignSelectedSkuAsync()
    {
        if (SelectedSkuToAdd is null || User is null || string.IsNullOrEmpty(User.Id)) return;
        var sku = SelectedSkuToAdd;
        var info = SkuCatalog.Resolve(sku.SkuPartNumber);
        IsBusy = true;
        StatusMessage = $"Asignando {info.FriendlyName}...";
        try
        {
            await _users.AssignLicenseAsync(User.Id, sku.SkuId, _log.Progress).ConfigureAwait(true);
            SelectedSkuToAdd = null;
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
    private async Task ResetPasswordAsync()
    {
        if (User is null || string.IsNullOrEmpty(User.Id)) return;
        var ok = await _dialogs.ConfirmAsync(
            $"Restablecer la contraseña de {User.UserPrincipalName}?\n\nSe generará una nueva contraseña aleatoria y se forzará cambio en el siguiente inicio de sesión.",
            "Confirmar reset password").ConfigureAwait(true);
        if (!ok) return;

        IsBusy = true;
        StatusMessage = "Reseteando password...";
        try
        {
            var newPwd = await _users.ResetPasswordAsync(User.Id, true, _log.Progress).ConfigureAwait(true);
            _clipboard.SetText(newPwd);
            await _dialogs.ShowAsync(
                $"Password temporal:\n\n{newPwd}\n\n(Copiada al portapapeles.)",
                "Password reseteada").ConfigureAwait(true);
            StatusMessage = "Password reseteada. (Forzará cambio próximo inicio.)";
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
    private async Task RevokeSessionsAsync()
    {
        if (User is null || string.IsNullOrEmpty(User.Id)) return;
        var ok = await _dialogs.ConfirmAsync(
            $"Revocar TODAS las sign-in sessions de {User.UserPrincipalName}?\n\nForzará re-autenticación en cualquier sesión activa (web, Teams, Outlook, etc).",
            "Confirmar revoke sessions",
            DialogIcon.Warning).ConfigureAwait(true);
        if (!ok) return;

        IsBusy = true;
        StatusMessage = "Revocando sessions...";
        try
        {
            await _users.RevokeSignInSessionsAsync(User.Id, _log.Progress).ConfigureAwait(true);
            StatusMessage = "Sign-in sessions revocadas.";
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
        var ok = await _dialogs.ConfirmAsync(
            $"Quitar TODAS las licencias asignadas a {User.UserPrincipalName}?",
            "Confirmar",
            DialogIcon.Warning).ConfigureAwait(true);
        if (!ok) return;

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
