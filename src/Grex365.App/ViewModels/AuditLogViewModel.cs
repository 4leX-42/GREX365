using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.App.ViewModels;

public sealed partial class AuditLogViewModel : ObservableObject
{
    private readonly IAuditLog _audit;
    private readonly IUiLogSink _log;

    [ObservableProperty] private DateTime _month = new(DateTime.Today.Year, DateTime.Today.Month, 1);
    [ObservableProperty] private string _sourceFilter = string.Empty;
    [ObservableProperty] private string _outcomeFilter = string.Empty;
    [ObservableProperty] private string _statusMessage = "Pulsa 'Cargar' para leer el mes seleccionado.";
    [ObservableProperty] private bool _isBusy;

    private readonly List<AuditRecord> _all = new();
    public ObservableCollection<AuditRecord> Filtered { get; } = new();

    public AuditLogViewModel(IAuditLog audit, IUiLogSink log)
    {
        _audit = audit;
        _log = log;
    }

    [RelayCommand]
    private async Task LoadAsync()
    {
        IsBusy = true;
        StatusMessage = $"Cargando {Month:yyyy-MM}...";
        try
        {
            var records = await _audit.ReadMonthAsync(Month.Year, Month.Month).ConfigureAwait(true);
            _all.Clear();
            _all.AddRange(records);
            ApplyFilters();
            StatusMessage = $"{records.Count} registros en {Month:yyyy-MM} · {Filtered.Count} filtrados";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("AuditLog", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
        }
    }

    partial void OnSourceFilterChanged(string value) => ApplyFilters();
    partial void OnOutcomeFilterChanged(string value) => ApplyFilters();

    private void ApplyFilters()
    {
        Filtered.Clear();
        IEnumerable<AuditRecord> q = _all;
        if (!string.IsNullOrWhiteSpace(SourceFilter))
        {
            q = q.Where(r => r.Source.Contains(SourceFilter, StringComparison.OrdinalIgnoreCase));
        }
        if (!string.IsNullOrWhiteSpace(OutcomeFilter))
        {
            q = q.Where(r => r.Outcome.Contains(OutcomeFilter, StringComparison.OrdinalIgnoreCase));
        }
        foreach (var r in q.OrderByDescending(r => r.Timestamp))
        {
            Filtered.Add(r);
        }
    }

    [RelayCommand]
    private void OpenFolder()
    {
        var path = _audit.GetMonthFilePath(Month.Year, Month.Month);
        var dir = Path.GetDirectoryName(path);
        if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir))
        {
            StatusMessage = "Carpeta de auditoría aún no existe.";
            return;
        }
        try
        {
            Process.Start(new ProcessStartInfo("explorer.exe", $"\"{dir}\"") { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            StatusMessage = "Error al abrir: " + ex.Message;
        }
    }
}
