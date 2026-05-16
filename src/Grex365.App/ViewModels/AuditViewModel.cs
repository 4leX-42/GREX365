using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Microsoft.Win32;

namespace Grex365.App.ViewModels;

public sealed partial class AuditViewModel : ObservableObject
{
    private readonly IAuditService _audit;
    private readonly IUiLogSink _log;

    [ObservableProperty] private AuditSummary? _summary;
    [ObservableProperty] private string _statusMessage = "Pulsa 'Ejecutar' para auditar.";
    [ObservableProperty] private bool _isBusy;

    public ObservableCollection<AuditFinding> Findings { get; } = new();

    public AuditViewModel(IAuditService audit, IUiLogSink log)
    {
        _audit = audit;
        _log = log;
    }

    [RelayCommand]
    private async Task RunAsync()
    {
        if (IsBusy)
        {
            return;
        }
        IsBusy = true;
        StatusMessage = "Ejecutando auditoría de identidades...";
        Findings.Clear();
        try
        {
            var (summary, findings) = await _audit.RunIdentityAuditAsync(_log.Progress).ConfigureAwait(true);
            Summary = summary;
            foreach (var f in findings)
            {
                Findings.Add(f);
            }
            StatusMessage = $"{summary.UsersTotal} usuarios analizados · {findings.Count} hallazgos";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("Audit", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private void ExportFindings()
    {
        if (Findings.Count == 0)
        {
            StatusMessage = "Sin hallazgos para exportar.";
            return;
        }

        var dlg = new SaveFileDialog
        {
            Title = "Guardar hallazgos",
            Filter = "CSV (*.csv)|*.csv",
            FileName = $"identity_audit_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };
        if (dlg.ShowDialog() != true)
        {
            return;
        }

        try
        {
            var sb = new StringBuilder();
            sb.AppendLine("Categoria,Identity,Detalle,Severidad");
            foreach (var f in Findings)
            {
                sb.Append(Escape(f.Category)).Append(',');
                sb.Append(Escape(f.Identity)).Append(',');
                sb.Append(Escape(f.Detail)).Append(',');
                sb.Append(Escape(f.Severity)).AppendLine();
            }
            File.WriteAllText(dlg.FileName, sb.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
            StatusMessage = $"Exportado: {Path.GetFileName(dlg.FileName)}";
            _log.Progress.Report(LogEntry.Ok("Audit", "Hallazgos exportados: " + dlg.FileName));
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("Audit", ex.Message, ex));
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
}
