using System.Collections.ObjectModel;
using System.Text;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.App.ViewModels;

public sealed partial class PsConsoleViewModel : ObservableObject
{
    private readonly IPowerShellRunner _runner;
    private readonly IUiLogSink _log;
    private CancellationTokenSource? _cts;
    private const int MaxHistory = 50;

    [ObservableProperty] private string _inputText = string.Empty;
    [ObservableProperty] private string _outputText = string.Empty;
    [ObservableProperty] private string _statusMessage = "Escribe un comando y pulsa Ejecutar (o Ctrl+Enter).";
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private int _historyIndex = -1;

    public ObservableCollection<string> History { get; } = new();

    public PsConsoleViewModel(IPowerShellRunner runner, IUiLogSink log)
    {
        _runner = runner;
        _log = log;
    }

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunAsync()
    {
        var script = (InputText ?? string.Empty).Trim();
        if (string.IsNullOrEmpty(script))
        {
            StatusMessage = "Comando vacío.";
            return;
        }

        AppendToHistory(script);
        HistoryIndex = -1;

        var sb = new StringBuilder(OutputText);
        sb.AppendLine($"PS> {script}");

        _cts = new CancellationTokenSource();
        IsBusy = true;
        RunCommand.NotifyCanExecuteChanged();
        CancelCommand.NotifyCanExecuteChanged();
        StatusMessage = "Ejecutando...";

        try
        {
            var result = await _runner.RunAsync(script, parameters: null, progress: _log.Progress, _cts.Token).ConfigureAwait(true);
            foreach (var item in result.Output)
            {
                sb.AppendLine(FormatOutput(item));
            }
            foreach (var err in result.Errors)
            {
                sb.AppendLine($"[ERROR] {err}");
            }
            sb.AppendLine();
            OutputText = sb.ToString();
            StatusMessage = result.Success
                ? $"OK · {result.Output.Count} resultado(s)"
                : $"ERROR · {result.Errors.Count} error(es)";
        }
        catch (OperationCanceledException)
        {
            sb.AppendLine("[CANCELADO]");
            sb.AppendLine();
            OutputText = sb.ToString();
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            sb.AppendLine($"[EXCEPCION] {ex.Message}");
            sb.AppendLine();
            OutputText = sb.ToString();
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("PsConsole", ex.Message, ex));
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

    private static string FormatOutput(object? item)
    {
        if (item is null) return "(null)";
        if (item is System.Management.Automation.PSObject pso)
        {
            // PSCustomObject: render property bag concisely.
            var props = pso.Properties;
            var parts = new List<string>();
            foreach (var p in props)
            {
                parts.Add($"{p.Name}={p.Value}");
            }
            return parts.Count > 0 ? string.Join("; ", parts) : pso.ToString();
        }
        return item.ToString() ?? string.Empty;
    }

    private bool CanRun() => !IsBusy;
    private bool CanCancel() => IsBusy;

    [RelayCommand(CanExecute = nameof(CanCancel))]
    private void Cancel()
    {
        try { _cts?.Cancel(); } catch { /* ignore */ }
    }

    [RelayCommand]
    private void Clear()
    {
        OutputText = string.Empty;
        StatusMessage = "Salida limpiada.";
    }

    [RelayCommand]
    private void HistoryPrev()
    {
        if (History.Count == 0) return;
        if (HistoryIndex < 0) HistoryIndex = History.Count - 1;
        else if (HistoryIndex > 0) HistoryIndex--;
        InputText = History[HistoryIndex];
    }

    [RelayCommand]
    private void HistoryNext()
    {
        if (History.Count == 0 || HistoryIndex < 0) return;
        if (HistoryIndex < History.Count - 1)
        {
            HistoryIndex++;
            InputText = History[HistoryIndex];
        }
        else
        {
            HistoryIndex = -1;
            InputText = string.Empty;
        }
    }

    private void AppendToHistory(string script)
    {
        // Don't add consecutive duplicates.
        if (History.Count > 0 && string.Equals(History[^1], script, StringComparison.Ordinal))
        {
            return;
        }
        History.Add(script);
        while (History.Count > MaxHistory)
        {
            History.RemoveAt(0);
        }
    }
}
