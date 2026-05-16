using System.IO;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Microsoft.Win32;

namespace Grex365.App.ViewModels;

public sealed partial class CertWizardViewModel : ObservableObject
{
    private readonly ICertificateGenerator _generator;
    private readonly IUiLogSink _log;

    [ObservableProperty] private string _commonName = "Grex365-Local";
    [ObservableProperty] private int _validDays = 365;
    [ObservableProperty] private string _exportDirectory =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Grex365", "certs");

    [ObservableProperty] private GeneratedCertificate? _generated;
    [ObservableProperty] private string _statusMessage = "Configura CN/validez y pulsa Generar.";
    [ObservableProperty] private bool _isBusy;

    public CertWizardViewModel(ICertificateGenerator generator, IUiLogSink log)
    {
        _generator = generator;
        _log = log;
    }

    [RelayCommand]
    private void BrowseExportFolder()
    {
        var dlg = new OpenFolderDialog
        {
            Title = "Carpeta de exportación del .cer"
        };
        if (string.IsNullOrWhiteSpace(ExportDirectory) is false && Directory.Exists(ExportDirectory))
        {
            dlg.InitialDirectory = ExportDirectory;
        }
        if (dlg.ShowDialog() == true)
        {
            ExportDirectory = dlg.FolderName;
        }
    }

    [RelayCommand]
    private async Task GenerateAsync()
    {
        IsBusy = true;
        StatusMessage = "Generando certificado...";
        try
        {
            var cn = CommonName?.Trim() ?? string.Empty;
            var dir = ExportDirectory?.Trim() ?? string.Empty;
            var days = ValidDays;

            var result = await Task.Run(() => _generator.GenerateAndStore(cn, days, dir, _log.Progress)).ConfigureAwait(true);
            Generated = result;
            StatusMessage = $"Generado. Thumbprint={result.Thumbprint}";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("Cert", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private void OpenExportFolder()
    {
        if (Generated is null)
        {
            return;
        }
        try
        {
            var folder = Path.GetDirectoryName(Generated.CerPath);
            if (folder is not null && Directory.Exists(folder))
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = folder,
                    UseShellExecute = true
                });
            }
        }
        catch (Exception ex)
        {
            _log.Progress.Report(LogEntry.Error("Cert", ex.Message, ex));
        }
    }
}
