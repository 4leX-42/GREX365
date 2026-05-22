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

                foreach (var skuId in assigned)
                {
                    var sku = _allSkus.FirstOrDefault(s => s.SkuId == skuId);
                    var partNumber = sku?.SkuPartNumber ?? skuId.ToString();
                    var info = SkuCatalog.Resolve(partNumber);
                    AssignedLicenses.Add(new AssignedLicenseRow(
                        SkuId: skuId,
                        SkuPartNumber: partNumber,
                        FriendlyName: info.FriendlyName,
                        CategoryLabel: SkuCatalog.CategoryLabel(info.Category)));
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
    private async Task RemoveLicenseAsync(AssignedLicenseRow? row)
    {
        if (row is null || User is null || string.IsNullOrEmpty(User.Id)) return;
        var confirm = System.Windows.MessageBox.Show(
            $"Quitar la licencia '{row.FriendlyName}' ({row.SkuPartNumber}) de {User.UserPrincipalName}?",
            "Confirmar",
            System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Question);
        if (confirm != System.Windows.MessageBoxResult.Yes) return;

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
