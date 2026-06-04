using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Csv;
using Grex365.Core.Groups;
using Grex365.Core.Models;
using Microsoft.Win32;

namespace Grex365.App.ViewModels;

public sealed partial class GroupsViewModel : ObservableObject
{
    private readonly IGroupsService _groups;
    private readonly IDistributionListsService _dls;
    private readonly IUiLogSink _log;
    private readonly IRbacGuard _rbac;
    private readonly IDialogService _dialogs;
    private CancellationTokenSource? _cts;
    private CancellationTokenSource? _debounceCts;

    [ObservableProperty] private string _searchQuery = string.Empty;
    [ObservableProperty] private GroupSummary? _selectedGroup;
    [ObservableProperty] private GroupMember? _selectedMember;
    [ObservableProperty] private string _newMembersText = string.Empty;
    [ObservableProperty] private string _memberToAdd = string.Empty;
    [ObservableProperty] private string _bulkDomain = string.Empty;
    [ObservableProperty] private string _bulkTypeChoice = "Auto"; // "Auto" | "M365" | "DL"
    [ObservableProperty] private string _statusMessage = string.Empty;
    [ObservableProperty] private bool _isBusy;

    public ObservableCollection<GroupSummary> Groups { get; } = new();
    public ObservableCollection<GroupMember> Members { get; } = new();
    public ObservableCollection<AddMemberResult> LastAddResults { get; } = new();
    public ObservableCollection<BulkGroupResult> BulkCreateResults { get; } = new();

    public GroupsViewModel(IGroupsService groups, IDistributionListsService dls, IUiLogSink log, IRbacGuard rbac, IDialogService dialogs)
    {
        _groups = groups;
        _dls = dls;
        _log = log;
        _rbac = rbac;
        _dialogs = dialogs;
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

    partial void OnSelectedGroupChanged(GroupSummary? value)
    {
        Members.Clear();
        LastAddResults.Clear();
        if (value is null)
        {
            return;
        }
        _ = LoadMembersAsync(value.Id);
    }

    [RelayCommand]
    private async Task SearchAsync()
    {
        EnsureToken();
        IsBusy = true;
        CancelCommand.NotifyCanExecuteChanged();
        StatusMessage = L10n.Get("Groups.Status.Searching");
        try
        {
            Groups.Clear();
            var found = await _groups.SearchAsync(SearchQuery, _cts!.Token).ConfigureAwait(true);
            foreach (var g in found)
            {
                Groups.Add(g);
            }
            StatusMessage = L10n.Format("Groups.Status.GroupsFound", found.Count);
            _log.Progress.Report(LogEntry.Info("Groups", $"Search '{SearchQuery}' -> {found.Count} resultados"));
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Groups", ex.Message, ex));
        }
        finally
        {
            DisposeToken();
        }
    }

    [RelayCommand(CanExecute = nameof(CanCancel))]
    private void Cancel() => _cts?.Cancel();

    private bool CanCancel() => IsBusy;

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
                    Groups.Clear();
                    StatusMessage = string.Empty;
                    return;
                }
                if (snapshot.Trim().Length < 2) return;
                await SearchAsync().ConfigureAwait(true);
            });
        });
    }

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

    private async Task LoadMembersAsync(string groupId)
    {
        EnsureToken();
        IsBusy = true;
        CancelCommand.NotifyCanExecuteChanged();
        try
        {
            var members = await _groups.GetMembersAsync(groupId, _cts!.Token).ConfigureAwait(true);
            Members.Clear();
            foreach (var m in members)
            {
                Members.Add(m);
            }
            StatusMessage = L10n.Format("Groups.Status.MembersCount", members.Count);
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Groups.Status.MembersError", ex.Message);
            _log.Progress.Report(LogEntry.Error("Groups", ex.Message, ex));
        }
        finally
        {
            DisposeToken();
        }
    }

    [RelayCommand]
    private async Task ImportCsvAsync()
    {
        if (SelectedGroup is null)
        {
            StatusMessage = L10n.Get("Groups.Status.SelectGroupFirst");
            return;
        }

        var dlg = new OpenFileDialog
        {
            Title = L10n.Get("Groups.Dialog.SelectMembersCsv"),
            Filter = "CSV (*.csv)|*.csv|Todos los archivos|*.*",
            CheckFileExists = true
        };
        if (dlg.ShowDialog() != true)
        {
            return;
        }

        EnsureToken();
        IsBusy = true;
        CancelCommand.NotifyCanExecuteChanged();
        StatusMessage = L10n.Format("Groups.Status.Reading", Path.GetFileName(dlg.FileName));
        try
        {
            var rows = FlexibleCsvReader.Read(dlg.FileName);
            var identifiers = new List<string>(rows.Count);
            foreach (var row in rows)
            {
                row.TryGetValue("Id", out var id);
                row.TryGetValue("Email", out var email);
                var pick = !string.IsNullOrWhiteSpace(id) ? id : email;
                if (!string.IsNullOrWhiteSpace(pick))
                {
                    identifiers.Add(pick.Trim());
                }
            }

            if (identifiers.Count == 0)
            {
                StatusMessage = L10n.Get("Groups.Status.CsvNoUsefulColumns");
                return;
            }

            StatusMessage = L10n.Format("Groups.Status.AddingFromCsv", identifiers.Count);
            var results = await _groups.AddMembersAsync(SelectedGroup.Id, identifiers, _log.Progress, _cts!.Token).ConfigureAwait(true);
            LastAddResults.Clear();
            foreach (var r in results)
            {
                LastAddResults.Add(r);
            }
            var ok = results.Count(r => r.Status == "AGREGADO");
            var existed = results.Count(r => r.Status == "YA_EXISTE");
            var errors = results.Count(r => r.Status is "ERROR" or "NO_RESUELTO");
            StatusMessage = L10n.Format("Groups.Status.CsvSummary", ok, existed, errors);
            DisposeToken();
            await LoadMembersAsync(SelectedGroup.Id).ConfigureAwait(true);
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
            DisposeToken();
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Groups", ex.Message, ex));
            DisposeToken();
        }
    }

    [RelayCommand]
    private async Task RemoveSelectedMemberAsync()
    {
        if (SelectedGroup is null || SelectedMember is null)
        {
            StatusMessage = L10n.Get("Groups.Status.SelectMember");
            return;
        }

        if (!await RequireAuthorizedAsync("Remove member").ConfigureAwait(true)) return;

        var ok = await _dialogs.ConfirmAsync(
            L10n.Format("Groups.Confirm.RemoveMemberBody", SelectedMember.DisplayName ?? SelectedMember.Id, SelectedGroup.DisplayName),
            L10n.Get("Groups.Confirm.RemoveMemberTitle"),
            DialogIcon.Warning).ConfigureAwait(true);
        if (!ok)
        {
            StatusMessage = L10n.Get("Common.Status.CancelledByUser");
            return;
        }

        EnsureToken();
        IsBusy = true;
        CancelCommand.NotifyCanExecuteChanged();
        var target = SelectedMember;
        StatusMessage = L10n.Format("Groups.Status.Removing", target.DisplayName ?? target.Id);
        try
        {
            await _groups.RemoveMemberAsync(SelectedGroup.Id, target.Id, _log.Progress, _cts!.Token).ConfigureAwait(true);
            Members.Remove(target);
            StatusMessage = L10n.Format("Groups.Status.Removed", target.DisplayName ?? target.Id);
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Groups", ex.Message, ex));
        }
        finally
        {
            DisposeToken();
        }
    }

    [RelayCommand]
    private void ExportMembers()
    {
        if (SelectedGroup is null || Members.Count == 0)
        {
            StatusMessage = L10n.Get("Groups.Status.NoMembersToExport");
            return;
        }

        var dlg = new SaveFileDialog
        {
            Title = L10n.Get("Groups.Dialog.SaveMembers"),
            Filter = "CSV (*.csv)|*.csv",
            FileName = $"members_{SelectedGroup.DisplayName}_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };
        if (dlg.ShowDialog() != true)
        {
            return;
        }

        try
        {
            var sb = new StringBuilder();
            sb.AppendLine("Id,DisplayName,Mail,UserPrincipalName");
            foreach (var m in Members)
            {
                sb.Append(Escape(m.Id)).Append(',');
                sb.Append(Escape(m.DisplayName)).Append(',');
                sb.Append(Escape(m.Mail)).Append(',');
                sb.Append(Escape(m.UserPrincipalName)).AppendLine();
            }
            File.WriteAllText(dlg.FileName, sb.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
            StatusMessage = L10n.Format("Common.Status.Exported", Path.GetFileName(dlg.FileName));
            _log.Progress.Report(LogEntry.Ok("Groups", "Miembros exportados: " + dlg.FileName));
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Groups", ex.Message, ex));
        }
    }

    [RelayCommand]
    private void ExportResults()
    {
        if (LastAddResults.Count == 0)
        {
            StatusMessage = L10n.Get("Common.Status.NoResultsToExport");
            return;
        }

        var dlg = new SaveFileDialog
        {
            Title = L10n.Get("Common.Dialog.SaveResults"),
            Filter = "CSV (*.csv)|*.csv",
            FileName = $"add_members_result_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };
        if (dlg.ShowDialog() != true)
        {
            return;
        }

        try
        {
            var sb = new StringBuilder();
            sb.AppendLine("Input,Status,Detail");
            foreach (var r in LastAddResults)
            {
                sb.Append(Escape(r.Input)).Append(',');
                sb.Append(Escape(r.Status)).Append(',');
                sb.Append(Escape(r.Detail)).AppendLine();
            }
            File.WriteAllText(dlg.FileName, sb.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
            StatusMessage = L10n.Format("Common.Status.Exported", Path.GetFileName(dlg.FileName));
            _log.Progress.Report(LogEntry.Ok("Groups", "Resultados exportados: " + dlg.FileName));
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Groups", ex.Message, ex));
        }
    }

    private static string Escape(string? value) => Grex365.Core.Csv.CsvEscaper.Escape(value);

    [RelayCommand]
    private async Task BulkCreateFromCsvAsync()
    {
        var domain = (BulkDomain ?? string.Empty).Trim().TrimStart('@');
        if (string.IsNullOrEmpty(domain))
        {
            StatusMessage = L10n.Get("Groups.Status.EnterDomain");
            return;
        }

        var dlg = new OpenFileDialog
        {
            Title = L10n.Get("Groups.Dialog.BulkCsv"),
            Filter = "CSV (*.csv)|*.csv|Todos|*.*",
            CheckFileExists = true
        };
        if (dlg.ShowDialog() != true)
        {
            return;
        }

        IReadOnlyList<BulkGroupRow> rows;
        try
        {
            var raw = FlexibleCsvReader.Read(dlg.FileName);
            rows = BulkGroupRowPreprocessor.Normalize(raw);
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Groups.Status.CsvError", ex.Message);
            return;
        }

        if (rows.Count == 0)
        {
            StatusMessage = L10n.Get("Groups.Status.CsvNoValidRows");
            return;
        }

        if (!await RequireAuthorizedAsync("Bulk create groups").ConfigureAwait(true)) return;

        var plan = BulkGroupPlanner.Plan(rows, BulkTypeChoice);
        var confirmMsg = BulkGroupPlanner.BuildConfirmMessage(plan, rows.Count, domain);

        var ok = await _dialogs.ConfirmAsync(confirmMsg, L10n.Get("Groups.Confirm.BulkCreateTitle")).ConfigureAwait(true);
        if (!ok)
        {
            StatusMessage = L10n.Get("Common.Status.CancelledByUser");
            return;
        }

        EnsureToken();
        IsBusy = true;
        CancelCommand.NotifyCanExecuteChanged();
        StatusMessage = L10n.Format("Groups.Status.Creating", plan.Breakdown);
        BulkCreateResults.Clear();
        try
        {
            var allResults = new List<BulkGroupResult>();
            if (plan.M365Rows.Count > 0)
            {
                var r1 = await _groups.CreateM365GroupsFromRowsAsync(plan.M365Rows, domain, _log.Progress, _cts!.Token).ConfigureAwait(true);
                allResults.AddRange(r1);
            }
            if (plan.DlRows.Count > 0)
            {
                var r2 = await _dls.CreateFromRowsAsync(plan.DlRows, domain, _log.Progress, _cts!.Token).ConfigureAwait(true);
                allResults.AddRange(r2);
            }
            foreach (var r in allResults)
            {
                BulkCreateResults.Add(r);
            }
            StatusMessage = BulkGroupPlanner.Summarize(allResults);
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("BulkGroups", ex.Message, ex));
        }
        finally
        {
            DisposeToken();
        }
    }

    [RelayCommand]
    private void ExportBulkCreateResults()
    {
        if (BulkCreateResults.Count == 0)
        {
            StatusMessage = L10n.Get("Common.Status.NoResultsToExport");
            return;
        }
        var dlg = new SaveFileDialog
        {
            Title = L10n.Get("Groups.Dialog.SaveCreateLog"),
            Filter = "CSV (*.csv)|*.csv",
            FileName = $"new_groups_log_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };
        if (dlg.ShowDialog() != true)
        {
            return;
        }
        try
        {
            var sb = new StringBuilder();
            sb.AppendLine("GroupName,GroupEmail,Action,UserEmail,Detail");
            foreach (var r in BulkCreateResults)
            {
                sb.Append(Escape(r.GroupName)).Append(',');
                sb.Append(Escape(r.GroupEmail)).Append(',');
                sb.Append(Escape(r.Action)).Append(',');
                sb.Append(Escape(r.UserEmail)).Append(',');
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

    // "Pick & append": take the UPN chosen in the UserPickerBox and add it (deduped) to the
    // multiline add-members box, then clear the picker for the next pick.
    [RelayCommand]
    private void AddPickedMember()
    {
        var pick = (MemberToAdd ?? string.Empty).Trim();
        if (pick.Length == 0)
        {
            return;
        }
        NewMembersText = MemberTextAppender.Append(NewMembersText, pick);
        MemberToAdd = string.Empty;
    }

    [RelayCommand]
    private async Task AddMembersAsync()
    {
        if (SelectedGroup is null)
        {
            StatusMessage = L10n.Get("Groups.Status.SelectGroupFirst");
            return;
        }

        var lines = (NewMembersText ?? string.Empty)
            .Split(new[] { '\n', '\r', ',', ';' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(l => !string.IsNullOrWhiteSpace(l))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (lines.Count == 0)
        {
            StatusMessage = L10n.Get("Groups.Status.NoEntriesToAdd");
            return;
        }

        EnsureToken();
        IsBusy = true;
        CancelCommand.NotifyCanExecuteChanged();
        StatusMessage = L10n.Format("Groups.Status.Adding", lines.Count);
        try
        {
            var results = await _groups.AddMembersAsync(SelectedGroup.Id, lines, _log.Progress, _cts!.Token).ConfigureAwait(true);
            LastAddResults.Clear();
            foreach (var r in results)
            {
                LastAddResults.Add(r);
            }
            var ok = results.Count(r => r.Status == "AGREGADO");
            var existed = results.Count(r => r.Status == "YA_EXISTE");
            var errors = results.Count(r => r.Status == "ERROR" || r.Status == "NO_RESUELTO");
            StatusMessage = L10n.Format("Groups.Status.AddSummary", ok, existed, errors);
            DisposeToken();
            await LoadMembersAsync(SelectedGroup.Id).ConfigureAwait(true);
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Groups", ex.Message, ex));
        }
        finally
        {
            if (IsBusy)
            {
                DisposeToken();
            }
        }
    }
}
