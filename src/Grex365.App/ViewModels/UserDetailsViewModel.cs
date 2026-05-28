using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Data;
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
    [ObservableProperty] private string _assignableSkuFilter = string.Empty;

    public ObservableCollection<GroupSummary> Memberships { get; } = new();
    public ObservableCollection<AssignedLicenseRow> AssignedLicenses { get; } = new();
    public ObservableCollection<SkuInfo> AssignableSkus { get; } = new();
    public ICollectionView AssignableSkusView { get; }

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
        AssignableSkusView = CollectionViewSource.GetDefaultView(AssignableSkus);
        AssignableSkusView.Filter = AssignableSkuMatches;
        _host.OpenRequested += (_, id) => _ = LoadAsync(id);
        _host.CloseRequested += (_, _) => Reset();
    }

    private bool AssignableSkuMatches(object obj)
    {
        if (obj is not SkuInfo sku) return false;
        var filter = (AssignableSkuFilter ?? string.Empty).Trim();
        if (filter.Length == 0) return true;
        var info = SkuCatalog.Resolve(sku.SkuPartNumber);
        return sku.SkuPartNumber.Contains(filter, StringComparison.OrdinalIgnoreCase)
            || info.FriendlyName.Contains(filter, StringComparison.OrdinalIgnoreCase)
            || SkuCatalog.CategoryLabel(info.Category).Contains(filter, StringComparison.OrdinalIgnoreCase);
    }

    partial void OnAssignableSkuFilterChanged(string value) => AssignableSkusView.Refresh();

    [RelayCommand]
    private void ClearAssignableSkuFilter() => AssignableSkuFilter = string.Empty;

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
        StatusMessage = L10n.Get("UserDetails.Status.LoadingProfile");
        try
        {
            var user = await _users.GetByIdAsync(userId, token).ConfigureAwait(true);
            if (user is null)
            {
                StatusMessage = L10n.Get("UserDetails.Status.NotFound");
                return;
            }
            User = user;
            HasUser = true;

            StatusMessage = L10n.Get("UserDetails.Status.LoadingGroups");
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

            StatusMessage = L10n.Get("UserDetails.Status.LoadingLicenses");
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

            StatusMessage = L10n.Format("UserDetails.Status.Summary", User.DisplayName, Memberships.Count, AssignedLicenses.Count);
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
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
            L10n.Format("UserDetails.Confirm.ToggleBody",
                newState ? L10n.Get("UserDetails.Verb.Enable") : L10n.Get("UserDetails.Verb.Disable"),
                User.UserPrincipalName),
            L10n.Get("Common.Confirm.Title")).ConfigureAwait(true);
        if (!ok) return;

        IsBusy = true;
        StatusMessage = newState ? L10n.Get("UserDetails.Status.Enabling") : L10n.Get("UserDetails.Status.Disabling");
        try
        {
            await _users.SetAccountEnabledAsync(User.Id, newState, _log.Progress).ConfigureAwait(true);
            await LoadAsync(User.Id).ConfigureAwait(true);
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
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
            L10n.Format("UserDetails.Confirm.RemoveLicenseBody", row.FriendlyName, row.SkuPartNumber, User.UserPrincipalName),
            L10n.Get("Common.Confirm.Title")).ConfigureAwait(true);
        if (!ok) return;

        IsBusy = true;
        StatusMessage = L10n.Format("UserDetails.Status.RemovingLicense", row.FriendlyName);
        try
        {
            await _users.RemoveLicenseAsync(User.Id, row.SkuId, _log.Progress).ConfigureAwait(true);
            await LoadAsync(User.Id).ConfigureAwait(true);
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
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
        StatusMessage = L10n.Format("UserDetails.Status.AssigningLicense", info.FriendlyName);
        try
        {
            await _users.AssignLicenseAsync(User.Id, sku.SkuId, _log.Progress).ConfigureAwait(true);
            SelectedSkuToAdd = null;
            await LoadAsync(User.Id).ConfigureAwait(true);
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
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
            L10n.Format("UserDetails.Confirm.ResetPasswordBody", User.UserPrincipalName),
            L10n.Get("UserDetails.Confirm.ResetPasswordTitle")).ConfigureAwait(true);
        if (!ok) return;

        IsBusy = true;
        StatusMessage = L10n.Get("UserDetails.Status.ResettingPassword");
        try
        {
            var newPwd = await _users.ResetPasswordAsync(User.Id, true, _log.Progress).ConfigureAwait(true);
            _clipboard.SetText(newPwd);
            await _dialogs.ShowAsync(
                L10n.Format("UserDetails.Dialog.PasswordResetBody", newPwd),
                L10n.Get("UserDetails.Dialog.PasswordResetTitle")).ConfigureAwait(true);
            StatusMessage = L10n.Get("UserDetails.Status.PasswordReset");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
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
            L10n.Format("UserDetails.Confirm.RevokeSessionsBody", User.UserPrincipalName),
            L10n.Get("UserDetails.Confirm.RevokeSessionsTitle"),
            DialogIcon.Warning).ConfigureAwait(true);
        if (!ok) return;

        IsBusy = true;
        StatusMessage = L10n.Get("UserDetails.Status.RevokingSessions");
        try
        {
            await _users.RevokeSignInSessionsAsync(User.Id, _log.Progress).ConfigureAwait(true);
            StatusMessage = L10n.Get("UserDetails.Status.SessionsRevoked");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
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
            L10n.Format("UserDetails.Confirm.RemoveAllLicensesBody", User.UserPrincipalName),
            L10n.Get("Common.Confirm.Title"),
            DialogIcon.Warning).ConfigureAwait(true);
        if (!ok) return;

        IsBusy = true;
        StatusMessage = L10n.Get("UserDetails.Status.RemovingLicenses");
        try
        {
            await _users.RemoveAllLicensesAsync(User.Id, _log.Progress).ConfigureAwait(true);
            await LoadAsync(User.Id).ConfigureAwait(true);
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("UserDetails", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
        }
    }
}
