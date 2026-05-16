using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Csv;
using Grex365.Core.Models;
using Microsoft.Win32;

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
    public ObservableCollection<BulkUserResult> BulkResults { get; } = new();

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

    [RelayCommand]
    private async Task ImportBulkCsvAsync()
    {
        var dlg = new OpenFileDialog
        {
            Title = "CSV de usuarios (UPN;Action) — Action=enable|disable|remove-licenses",
            Filter = "CSV (*.csv)|*.csv|Todos|*.*",
            CheckFileExists = true
        };
        if (dlg.ShowDialog() != true)
        {
            return;
        }

        EnsureToken();
        IsBusy = true;
        CancelCommand.NotifyCanExecuteChanged();
        StatusMessage = $"Procesando {Path.GetFileName(dlg.FileName)}...";
        BulkResults.Clear();
        try
        {
            var rows = FlexibleCsvReader.Read(dlg.FileName);
            var ok = 0; var skipped = 0; var err = 0;
            foreach (var row in rows)
            {
                _cts!.Token.ThrowIfCancellationRequested();
                row.TryGetValue("UPN", out var upn);
                row.TryGetValue("Action", out var actionRaw);
                var action = (actionRaw ?? string.Empty).Trim().ToLowerInvariant();
                if (string.IsNullOrWhiteSpace(upn) || string.IsNullOrWhiteSpace(action))
                {
                    BulkResults.Add(new BulkUserResult(upn ?? string.Empty, action, "INVALIDO", "UPN o Action vacíos"));
                    skipped++;
                    continue;
                }

                try
                {
                    var user = await _users.GetByIdAsync(upn.Trim(), _cts.Token).ConfigureAwait(true);
                    if (user is null)
                    {
                        BulkResults.Add(new BulkUserResult(upn, action, "NO_RESUELTO", "Usuario no encontrado"));
                        err++;
                        continue;
                    }

                    switch (action)
                    {
                        case "enable":
                            await _users.SetAccountEnabledAsync(user.Id, true, _log.Progress, _cts.Token).ConfigureAwait(true);
                            BulkResults.Add(new BulkUserResult(upn, action, "OK", "Habilitado"));
                            ok++;
                            break;
                        case "disable":
                            await _users.SetAccountEnabledAsync(user.Id, false, _log.Progress, _cts.Token).ConfigureAwait(true);
                            BulkResults.Add(new BulkUserResult(upn, action, "OK", "Deshabilitado"));
                            ok++;
                            break;
                        case "remove-licenses":
                            await _users.RemoveAllLicensesAsync(user.Id, _log.Progress, _cts.Token).ConfigureAwait(true);
                            BulkResults.Add(new BulkUserResult(upn, action, "OK", $"{user.AssignedLicenseCount} licencias retiradas"));
                            ok++;
                            break;
                        default:
                            BulkResults.Add(new BulkUserResult(upn, action, "INVALIDO", "Action no soportada"));
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
                    BulkResults.Add(new BulkUserResult(upn, action, "ERROR", ex.Message));
                    err++;
                }
            }

            StatusMessage = $"Bulk: OK={ok}  Skip={skipped}  Err={err}";
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
    private void ExportBulkResults()
    {
        if (BulkResults.Count == 0)
        {
            StatusMessage = "Sin resultados para exportar.";
            return;
        }
        var dlg = new SaveFileDialog
        {
            Title = "Guardar resultados",
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
            StatusMessage = $"Exportado: {Path.GetFileName(dlg.FileName)}";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
        }
    }

    private static string Escape(string? value)
    {
        var v = value ?? string.Empty;
        if (v.Contains(',') || v.Contains('"') || v.Contains('\n'))
        {
            return '"' + v.Replace("\"", "\"\"") + '"';
        }
        return v;
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
