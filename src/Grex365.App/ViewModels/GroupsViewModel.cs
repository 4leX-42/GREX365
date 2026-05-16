using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.App.ViewModels;

public sealed partial class GroupsViewModel : ObservableObject
{
    private readonly IGroupsService _groups;
    private readonly IUiLogSink _log;

    [ObservableProperty] private string _searchQuery = string.Empty;
    [ObservableProperty] private GroupSummary? _selectedGroup;
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
        IsBusy = true;
        StatusMessage = "Buscando...";
        try
        {
            Groups.Clear();
            var found = await _groups.SearchAsync(SearchQuery).ConfigureAwait(true);
            foreach (var g in found)
            {
                Groups.Add(g);
            }
            StatusMessage = $"{found.Count} grupos.";
            _log.Progress.Report(LogEntry.Info("Groups", $"Search '{SearchQuery}' → {found.Count} resultados"));
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("Groups", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
        }
    }

    private async Task LoadMembersAsync(string groupId)
    {
        IsBusy = true;
        try
        {
            var members = await _groups.GetMembersAsync(groupId).ConfigureAwait(true);
            Members.Clear();
            foreach (var m in members)
            {
                Members.Add(m);
            }
            StatusMessage = $"{members.Count} miembros.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error miembros: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("Groups", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
        }
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

        IsBusy = true;
        StatusMessage = $"Añadiendo {lines.Count}...";
        try
        {
            var results = await _groups.AddMembersAsync(SelectedGroup.Id, lines, _log.Progress).ConfigureAwait(true);
            LastAddResults.Clear();
            foreach (var r in results)
            {
                LastAddResults.Add(r);
            }
            var ok = results.Count(r => r.Status == "AGREGADO");
            var existed = results.Count(r => r.Status == "YA_EXISTE");
            var errors = results.Count(r => r.Status == "ERROR" || r.Status == "NO_RESUELTO");
            StatusMessage = $"OK={ok}  YaExistía={existed}  Error={errors}";
            await LoadMembersAsync(SelectedGroup.Id).ConfigureAwait(true);
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("Groups", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
        }
    }
}
