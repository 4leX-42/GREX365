using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Audit;
using Grex365.Core.Models;

namespace Grex365.App.ViewModels;

public sealed partial class AuditLogViewModel : ObservableObject
{
    private readonly IAuditLog _audit;
    private readonly IUiLogSink _log;

    [ObservableProperty] private DateTime _month = new(DateTime.Today.Year, DateTime.Today.Month, 1);
    [ObservableProperty] private string _sourceFilter = string.Empty;
    [ObservableProperty] private string _outcomeFilter = string.Empty;
    [ObservableProperty] private string _statusMessage = L10n.Get("AuditLog.Status.Initial");
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private AuditMetrics? _summary;

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
        StatusMessage = L10n.Format("AuditLog.Status.Loading", Month.ToString("yyyy-MM"));
        try
        {
            var records = await _audit.ReadMonthAsync(Month.Year, Month.Month).ConfigureAwait(true);
            _all.Clear();
            _all.AddRange(records);
            ApplyFilters();
            Summary = MetricsAggregator.Compute(records);
            StatusMessage = L10n.Format("AuditLog.Status.Summary", records.Count, Month.ToString("yyyy-MM"), Filtered.Count);
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
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
            StatusMessage = L10n.Get("AuditLog.Status.FolderNotExist");
            return;
        }
        try
        {
            Process.Start(new ProcessStartInfo("explorer.exe", $"\"{dir}\"") { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("AuditLog.Status.OpenError", ex.Message);
        }
    }
}
