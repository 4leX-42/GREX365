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
    private CancellationTokenSource? _cts;

    [ObservableProperty] private AuditSummary? _summary;
    [ObservableProperty] private string _statusMessage = "Pulsa 'Ejecutar' para auditar.";
    [ObservableProperty] private bool _isBusy;

    public ObservableCollection<AuditFinding> Findings { get; } = new();

    public AuditViewModel(IAuditService audit, IUiLogSink log)
    {
        _audit = audit;
        _log = log;
    }

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunAsync()
    {
        if (IsBusy)
        {
            return;
        }
        _cts = new CancellationTokenSource();
        IsBusy = true;
        RunCommand.NotifyCanExecuteChanged();
        CancelCommand.NotifyCanExecuteChanged();
        StatusMessage = "Ejecutando auditoría de identidades...";
        Findings.Clear();
        try
        {
            var (summary, findings) = await _audit.RunIdentityAuditAsync(_log.Progress, _cts.Token).ConfigureAwait(true);
            Summary = summary;
            foreach (var f in findings)
            {
                Findings.Add(f);
            }

            var groupFindings = await _audit.RunGroupsAuditAsync(_log.Progress, _cts.Token).ConfigureAwait(true);
            foreach (var f in groupFindings)
            {
                Findings.Add(f);
            }

            StatusMessage = $"{summary.UsersTotal} usuarios · {findings.Count + groupFindings.Count} hallazgos totales";
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("Audit", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
            _cts?.Dispose();
            _cts = null;
            RunCommand.NotifyCanExecuteChanged();
            CancelCommand.NotifyCanExecuteChanged();
        }
    }

    [RelayCommand(CanExecute = nameof(CanCancel))]
    private void Cancel() => _cts?.Cancel();

    private bool CanRun() => !IsBusy;
    private bool CanCancel() => IsBusy;

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
