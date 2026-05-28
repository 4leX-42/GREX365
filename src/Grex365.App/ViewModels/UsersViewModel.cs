using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Text;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Csv;
using Grex365.Core.Models;
using Grex365.Core.Users;
using Microsoft.Win32;

namespace Grex365.App.ViewModels;

public sealed partial class UsersViewModel : ObservableObject
{
    private readonly IUsersService _users;
    private readonly IUiLogSink _log;
    private readonly IRbacGuard _rbac;
    private readonly IDialogService _dialogs;
    private readonly IConnectionStateMonitor? _monitor;
    private CancellationTokenSource? _cts;
    private CancellationTokenSource? _debounceCts;

    [ObservableProperty] private string _searchQuery = string.Empty;
    [ObservableProperty] private UserSummary? _selectedUser;
    [ObservableProperty] private SkuInfo? _selectedSku;
    [ObservableProperty] private string _statusMessage = string.Empty;
    [ObservableProperty] private bool _isBusy;

    public ObservableCollection<UserSummary> Users { get; } = new();
    public ObservableCollection<GroupSummary> Memberships { get; } = new();
    public ObservableCollection<BulkUserResult> BulkResults { get; } = new();
    public ObservableCollection<SkuInfo> AvailableSkus { get; } = new();

    public UsersViewModel(IUsersService users, IUiLogSink log, IRbacGuard rbac, IDialogService dialogs, IConnectionStateMonitor? monitor = null)
    {
        _users = users;
        _log = log;
        _rbac = rbac;
        _dialogs = dialogs;
        _monitor = monitor;
        if (_monitor is not null)
        {
            _monitor.PropertyChanged += OnMonitorChanged;
            // Fire once on construction in case Graph is already connected
            // (typical: VM singleton instantiated lazily after auto-connect completes).
            // Use dispatcher path so this also handles VM resolution from a
            // background thread (e.g. when monitor sends ctor through DI mid-poll).
            if (_monitor.Current.GraphConnected)
            {
                DispatchAutoLoad();
            }
        }
    }

    private void OnMonitorChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName != nameof(IConnectionStateMonitor.Current)) return;
        if (_monitor is null || !_monitor.Current.GraphConnected) return;
        // ConnectionStateMonitor fires PropertyChanged from its 1s background poll
        // loop — we must marshal to the UI dispatcher before touching
        // ObservableCollection (AvailableSkus has a CollectionView bound to it via
        // ComboBox.ItemsSource, which requires Dispatcher-thread mutations) and
        // before triggering an AsyncRelayCommand whose CanExecute callbacks read
        // DependencyObject state on completion.
        DispatchAutoLoad();
    }

    private void DispatchAutoLoad()
    {
        var dispatcher = System.Windows.Application.Current?.Dispatcher;
        if (dispatcher is null || dispatcher.CheckAccess())
        {
            _ = TryAutoLoadSkusAsync();
        }
        else
        {
            // Func<Task> overload — keeps the async chain bound to the dispatcher so the
            // AsyncRelayCommand completes (and mutates the bound AvailableSkus
            // CollectionView) on the UI thread. The Action overload fire-and-forgets the
            // inner Task and its continuations lose UI affinity → cross-thread
            // CollectionView crash when auto-connect fires from the monitor poll thread.
            dispatcher.InvokeAsync(TryAutoLoadSkusAsync);
        }
    }

    private async Task TryAutoLoadSkusAsync()
    {
        if (AvailableSkus.Count > 0 || IsBusy) return;
        try
        {
            await LoadSkusCommand.ExecuteAsync(null).ConfigureAwait(true);
        }
        catch
        {
            // LoadSkusAsync already surfaces errors via StatusMessage + log sink.
        }
    }

    private async Task<bool> RequireAuthorizedAsync(string contextName)
    {
        var decision = await _rbac.EvaluateAsync().ConfigureAwait(true);
        if (decision.Allowed)
        {
            return true;
        }
        StatusMessage = decision.Reason;
        _log.Progress.Report(LogEntry.Warn("RBAC", $"{contextName} bloqueado: {decision.Reason}"));
        return false;
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
        StatusMessage = L10n.Get("Users.Status.Searching");
        try
        {
            Users.Clear();
            var found = await _users.SearchAsync(SearchQuery, _cts!.Token).ConfigureAwait(true);
            foreach (var u in found)
            {
                Users.Add(u);
            }
            StatusMessage = L10n.Format("Users.Status.UsersFound", found.Count);
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
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
            StatusMessage = L10n.Format("Users.Status.MembershipsCount", groups.Count);
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
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
            StatusMessage = L10n.Get("Users.Status.SelectUser");
            return;
        }
        if (!enabled)
        {
            if (!await RequireAuthorizedAsync("Disable user").ConfigureAwait(true)) return;

            var ok = await _dialogs.ConfirmAsync(
                L10n.Format("Users.Confirm.DisableBody", SelectedUser.DisplayName, SelectedUser.UserPrincipalName),
                L10n.Get("Users.Confirm.DisableTitle"),
                DialogIcon.Warning).ConfigureAwait(true);
            if (!ok)
            {
                StatusMessage = L10n.Get("Common.Status.CancelledByUser");
                return;
            }
        }
        EnsureToken();
        IsBusy = true;
        StatusMessage = enabled ? L10n.Get("Users.Status.Enabling") : L10n.Get("Users.Status.Disabling");
        try
        {
            await _users.SetAccountEnabledAsync(SelectedUser.Id, enabled, _log.Progress, _cts!.Token).ConfigureAwait(true);
            StatusMessage = enabled ? L10n.Get("Users.Status.Enabled") : L10n.Get("Users.Status.Disabled");
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
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Users", ex.Message, ex));
        }
        finally
        {
            DisposeToken();
        }
    }

    [RelayCommand]
    private async Task LoadSkusAsync()
    {
        EnsureToken();
        IsBusy = true;
        StatusMessage = L10n.Get("Users.Status.LoadingSkus");
        try
        {
            var skus = await _users.ListSkusAsync(_cts!.Token).ConfigureAwait(true);
            AvailableSkus.Clear();
            foreach (var s in skus)
            {
                AvailableSkus.Add(s);
            }
            StatusMessage = L10n.Format("Users.Status.SkusAvailable", skus.Count);
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Users", ex.Message, ex));
        }
        finally
        {
            DisposeToken();
        }
    }

    [RelayCommand]
    private async Task AssignLicenseAsync()
    {
        if (SelectedUser is null)
        {
            StatusMessage = L10n.Get("Users.Status.SelectUser");
            return;
        }
        if (SelectedSku is null)
        {
            StatusMessage = L10n.Get("Users.Status.SelectSku");
            return;
        }
        if (SelectedSku.Available <= 0)
        {
            var ok = await _dialogs.ConfirmAsync(
                L10n.Format("Users.Confirm.AssignNoSeatsBody", SelectedSku.SkuPartNumber, SelectedSku.Available, SelectedSku.Enabled),
                L10n.Get("Users.Confirm.AssignTitle"),
                DialogIcon.Warning).ConfigureAwait(true);
            if (!ok)
            {
                StatusMessage = L10n.Get("Common.Status.CancelledByUser");
                return;
            }
        }
        EnsureToken();
        IsBusy = true;
        StatusMessage = L10n.Format("Users.Status.Assigning", SelectedSku.SkuPartNumber);
        try
        {
            await _users.AssignLicenseAsync(SelectedUser.Id, SelectedSku.SkuId, _log.Progress, _cts!.Token).ConfigureAwait(true);
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
            StatusMessage = L10n.Format("Users.Status.LicenseAssigned", SelectedSku.SkuPartNumber);
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
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
            StatusMessage = L10n.Get("Users.Status.SelectUser");
            return;
        }
        if (SelectedUser.AssignedLicenseCount > 0)
        {
            if (!await RequireAuthorizedAsync("Remove licenses").ConfigureAwait(true)) return;

            var ok = await _dialogs.ConfirmAsync(
                L10n.Format("Users.Confirm.RemoveLicensesBody", SelectedUser.AssignedLicenseCount, SelectedUser.DisplayName),
                L10n.Get("Users.Confirm.RemoveLicensesTitle"),
                DialogIcon.Warning).ConfigureAwait(true);
            if (!ok)
            {
                StatusMessage = L10n.Get("Common.Status.CancelledByUser");
                return;
            }
        }
        EnsureToken();
        IsBusy = true;
        StatusMessage = L10n.Get("Users.Status.RemovingLicenses");
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
            StatusMessage = L10n.Get("Users.Status.LicensesRemoved");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Users", ex.Message, ex));
        }
        finally
        {
            DisposeToken();
        }
    }

    [RelayCommand]
    private async Task ImportBulkCsvAsync()
    {
        var dlg = new OpenFileDialog
        {
            Title = L10n.Get("Users.Dialog.BulkCsv"),
            Filter = "CSV (*.csv)|*.csv|Todos|*.*",
            CheckFileExists = true
        };
        if (dlg.ShowDialog() != true)
        {
            return;
        }

        if (!await RequireAuthorizedAsync("Bulk users").ConfigureAwait(true)) return;

        EnsureToken();
        IsBusy = true;
        CancelCommand.NotifyCanExecuteChanged();
        StatusMessage = L10n.Format("Users.Status.Processing", Path.GetFileName(dlg.FileName));
        BulkResults.Clear();
        try
        {
            var rows = FlexibleCsvReader.Read(dlg.FileName);
            var parsed = rows.Select(r =>
            {
                r.TryGetValue("UPN", out var upn);
                r.TryGetValue("Action", out var actionRaw);
                return (Upn: upn, ActionRaw: actionRaw, Action: BulkUserActionParser.Parse(actionRaw));
            }).ToList();

            IReadOnlyList<SkuInfo> skus = Array.Empty<SkuInfo>();
            if (parsed.Any(p => p.Action.Kind == BulkUserActionKind.AssignLicense))
            {
                StatusMessage = L10n.Get("Users.Status.LoadingSkusAvailable");
                skus = await _users.ListSkusAsync(_cts!.Token).ConfigureAwait(true);
            }

            var ok = 0; var skipped = 0; var err = 0;
            foreach (var entry in parsed)
            {
                _cts!.Token.ThrowIfCancellationRequested();
                var upn = entry.Upn;
                var actionDisplay = entry.ActionRaw ?? string.Empty;
                if (string.IsNullOrWhiteSpace(upn) || entry.Action.Kind == BulkUserActionKind.Unknown)
                {
                    BulkResults.Add(new BulkUserResult(upn ?? string.Empty, actionDisplay, "INVALIDO",
                        string.IsNullOrWhiteSpace(upn) ? "UPN vacío" : "Action no soportada"));
                    skipped++;
                    continue;
                }

                try
                {
                    var user = await _users.GetByIdAsync(upn.Trim(), _cts.Token).ConfigureAwait(true);
                    if (user is null)
                    {
                        BulkResults.Add(new BulkUserResult(upn, actionDisplay, "NO_RESUELTO", "Usuario no encontrado"));
                        err++;
                        continue;
                    }

                    switch (entry.Action.Kind)
                    {
                        case BulkUserActionKind.Enable:
                            await _users.SetAccountEnabledAsync(user.Id, true, _log.Progress, _cts.Token).ConfigureAwait(true);
                            BulkResults.Add(new BulkUserResult(upn, actionDisplay, "OK", "Habilitado"));
                            ok++;
                            break;
                        case BulkUserActionKind.Disable:
                            await _users.SetAccountEnabledAsync(user.Id, false, _log.Progress, _cts.Token).ConfigureAwait(true);
                            BulkResults.Add(new BulkUserResult(upn, actionDisplay, "OK", "Deshabilitado"));
                            ok++;
                            break;
                        case BulkUserActionKind.RemoveLicenses:
                            await _users.RemoveAllLicensesAsync(user.Id, _log.Progress, _cts.Token).ConfigureAwait(true);
                            BulkResults.Add(new BulkUserResult(upn, actionDisplay, "OK", $"{user.AssignedLicenseCount} licencias retiradas"));
                            ok++;
                            break;
                        case BulkUserActionKind.AssignLicense:
                            var sku = BulkUserActionParser.FindByPartNumber(skus, entry.Action.SkuPartNumber);
                            if (sku is null)
                            {
                                BulkResults.Add(new BulkUserResult(upn, actionDisplay, "INVALIDO", $"SKU no encontrada: {entry.Action.SkuPartNumber}"));
                                skipped++;
                                break;
                            }
                            await _users.AssignLicenseAsync(user.Id, sku.SkuId, _log.Progress, _cts.Token).ConfigureAwait(true);
                            BulkResults.Add(new BulkUserResult(upn, actionDisplay, "OK", $"Asignada {sku.SkuPartNumber}"));
                            ok++;
                            break;
                        default:
                            BulkResults.Add(new BulkUserResult(upn, actionDisplay, "INVALIDO", "Action no soportada"));
                            skipped++;
                            break;
                    }
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    BulkResults.Add(new BulkUserResult(upn, actionDisplay, "ERROR", ex.Message));
                    err++;
                }
            }

            StatusMessage = L10n.Format("Users.Status.BulkSummary", ok, skipped, err);
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Users", ex.Message, ex));
        }
        finally
        {
            DisposeToken();
        }
    }

    [RelayCommand]
    private void ExportBulkResults()
    {
        if (BulkResults.Count == 0)
        {
            StatusMessage = L10n.Get("Common.Status.NoResultsToExport");
            return;
        }
        var dlg = new SaveFileDialog
        {
            Title = L10n.Get("Common.Dialog.SaveResults"),
            Filter = "CSV (*.csv)|*.csv",
            FileName = $"users_bulk_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };
        if (dlg.ShowDialog() != true)
        {
            return;
        }
        try
        {
            var sb = new StringBuilder();
            sb.AppendLine("UPN,Action,Status,Detail");
            foreach (var r in BulkResults)
            {
                sb.Append(Escape(r.Upn)).Append(',');
                sb.Append(Escape(r.Action)).Append(',');
                sb.Append(Escape(r.Status)).Append(',');
                sb.Append(Escape(r.Detail)).AppendLine();
            }
            File.WriteAllText(dlg.FileName, sb.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
            StatusMessage = L10n.Format("Common.Status.Exported", Path.GetFileName(dlg.FileName));
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
        }
    }

    private static string Escape(string? value) => Grex365.Core.Csv.CsvEscaper.Escape(value);

    [RelayCommand(CanExecute = nameof(CanCancel))]
    private void Cancel() => _cts?.Cancel();

    // Real-time typeahead: debounced 250ms; min 2 chars; cancels in-flight Graph call.
    partial void OnSearchQueryChanged(string value)
    {
        _debounceCts?.Cancel();
        _debounceCts = new CancellationTokenSource();
        var token = _debounceCts.Token;
        var snapshot = value ?? string.Empty;

        _ = Task.Run(async () =>
        {
            try
            {
                await Task.Delay(250, token).ConfigureAwait(false);
            }
            catch (OperationCanceledException) { return; }

            await System.Windows.Application.Current.Dispatcher.InvokeAsync(async () =>
            {
                if (token.IsCancellationRequested) return;
                if (!string.Equals(SearchQuery, snapshot, StringComparison.Ordinal)) return;
                if (string.IsNullOrWhiteSpace(snapshot))
                {
                    Users.Clear();
                    StatusMessage = string.Empty;
                    return;
                }
                if (snapshot.Trim().Length < 2) return;
                await SearchAsync().ConfigureAwait(true);
            });
        });
    }

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
