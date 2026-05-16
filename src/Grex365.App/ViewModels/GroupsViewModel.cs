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

public sealed partial class GroupsViewModel : ObservableObject
{
    private readonly IGroupsService _groups;
    private readonly IUiLogSink _log;
    private CancellationTokenSource? _cts;

    [ObservableProperty] private string _searchQuery = string.Empty;
    [ObservableProperty] private GroupSummary? _selectedGroup;
    [ObservableProperty] private GroupMember? _selectedMember;
    [ObservableProperty] private string _newMembersText = string.Empty;
    [ObservableProperty] private string _statusMessage = string.Empty;
    [ObservableProperty] private bool _isBusy;

    public ObservableCollection<GroupSummary> Groups { get; } = new();
    public ObservableCollection<GroupMember> Members { get; } = new();
    public ObservableCollection<AddMemberResult> LastAddResults { get; } = new();

    public GroupsViewModel(IGroupsService groups, IUiLogSink log)
    {
        _groups = groups;
        _log = log;
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
        StatusMessage = "Buscando...";
        try
        {
            Groups.Clear();
            var found = await _groups.SearchAsync(SearchQuery, _cts!.Token).ConfigureAwait(true);
            foreach (var g in found)
            {
                Groups.Add(g);
            }
            StatusMessage = $"{found.Count} grupos.";
            _log.Progress.Report(LogEntry.Info("Groups", $"Search '{SearchQuery}' -> {found.Count} resultados"));
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
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
            StatusMessage = $"{members.Count} miembros.";
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error miembros: " + ex.Message;
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
            StatusMessage = "Selecciona un grupo primero.";
            return;
        }

        var dlg = new OpenFileDialog
        {
            Title = "Seleccionar CSV de miembros",
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
        StatusMessage = $"Leyendo {Path.GetFileName(dlg.FileName)}...";
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
                StatusMessage = "CSV sin columnas Email/Id útiles.";
                return;
            }

            StatusMessage = $"Añadiendo {identifiers.Count} desde CSV...";
            var results = await _groups.AddMembersAsync(SelectedGroup.Id, identifiers, _log.Progress, _cts!.Token).ConfigureAwait(true);
            LastAddResults.Clear();
            foreach (var r in results)
            {
                LastAddResults.Add(r);
            }
            var ok = results.Count(r => r.Status == "AGREGADO");
            var existed = results.Count(r => r.Status == "YA_EXISTE");
            var errors = results.Count(r => r.Status is "ERROR" or "NO_RESUELTO");
            StatusMessage = $"CSV: OK={ok}  YaExistia={existed}  Error={errors}";
            DisposeToken();
            await LoadMembersAsync(SelectedGroup.Id).ConfigureAwait(true);
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
            DisposeToken();
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("Groups", ex.Message, ex));
            DisposeToken();
        }
    }

    [RelayCommand]
    private async Task RemoveSelectedMemberAsync()
    {
        if (SelectedGroup is null || SelectedMember is null)
        {
            StatusMessage = "Selecciona un miembro.";
            return;
        }

        EnsureToken();
        IsBusy = true;
        CancelCommand.NotifyCanExecuteChanged();
        var target = SelectedMember;
        StatusMessage = $"Eliminando {target.DisplayName ?? target.Id}...";
        try
        {
            await _groups.RemoveMemberAsync(SelectedGroup.Id, target.Id, _log.Progress, _cts!.Token).ConfigureAwait(true);
            Members.Remove(target);
            StatusMessage = $"Eliminado: {target.DisplayName ?? target.Id}";
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
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
            StatusMessage = "Sin miembros para exportar.";
            return;
        }

        var dlg = new SaveFileDialog
        {
            Title = "Guardar miembros",
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
            StatusMessage = $"Exportado: {Path.GetFileName(dlg.FileName)}";
            _log.Progress.Report(LogEntry.Ok("Groups", "Miembros exportados: " + dlg.FileName));
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("Groups", ex.Message, ex));
        }
    }

    [RelayCommand]
    private void ExportResults()
    {
        if (LastAddResults.Count == 0)
        {
            StatusMessage = "Sin resultados para exportar.";
            return;
        }

        var dlg = new SaveFileDialog
        {
            Title = "Guardar resultados",
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
            StatusMessage = $"Exportado: {Path.GetFileName(dlg.FileName)}";
            _log.Progress.Report(LogEntry.Ok("Groups", "Resultados exportados: " + dlg.FileName));
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("Groups", ex.Message, ex));
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

    [RelayCommand]
    private async Task AddMembersAsync()
    {
        if (SelectedGroup is null)
        {
            StatusMessage = "Selecciona un grupo primero.";
            return;
        }

        var lines = (NewMembersText ?? string.Empty)
            .Split(new[] { '\n', '\r', ',', ';' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(l => !string.IsNullOrWhiteSpace(l))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (lines.Count == 0)
        {
            StatusMessage = "Sin entradas para añadir.";
            return;
        }

        EnsureToken();
        IsBusy = true;
        CancelCommand.NotifyCanExecuteChanged();
        StatusMessage = $"Añadiendo {lines.Count}...";
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
            StatusMessage = $"OK={ok}  YaExistia={existed}  Error={errors}";
            DisposeToken();
            await LoadMembersAsync(SelectedGroup.Id).ConfigureAwait(true);
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
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
