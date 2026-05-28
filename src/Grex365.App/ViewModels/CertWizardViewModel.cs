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
    private readonly IAppRegistrationService _appReg;
    private readonly IGraphConnection _graph;
    private readonly ICertConfigStore _certStore;

    [ObservableProperty] private string _commonName = "Grex365-Local";
    [ObservableProperty] private int _validDays = 365;
    [ObservableProperty] private string _exportDirectory =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Grex365", "certs");

    [ObservableProperty] private GeneratedCertificate? _generated;
    [ObservableProperty] private string _statusMessage = "Configura CN/validez y pulsa Generar.";
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private string _pfxPath = string.Empty;
    [ObservableProperty] private string _pfxStatus = string.Empty;

    [ObservableProperty] private string _appRegDisplayName = "Grex365";
    [ObservableProperty] private string _appRegStatus = string.Empty;
    [ObservableProperty] private AppRegistrationResult? _appRegResult;

    public CertWizardViewModel(
        ICertificateGenerator generator,
        IUiLogSink log,
        IAppRegistrationService appReg,
        IGraphConnection graph,
        ICertConfigStore certStore)
    {
        _generator = generator;
        _log = log;
        _appReg = appReg;
        _graph = graph;
        _certStore = certStore;
    }

    [RelayCommand]
    private void BrowseExportFolder()
    {
        var dlg = new OpenFolderDialog
        {
            Title = L10n.Get("CertWizard.Dialog.ExportFolder")
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
        StatusMessage = L10n.Get("CertWizard.Status.Generating");
        try
        {
            var cn = CommonName?.Trim() ?? string.Empty;
            var dir = ExportDirectory?.Trim() ?? string.Empty;
            var days = ValidDays;

            var result = await Task.Run(() => _generator.GenerateAndStore(cn, days, dir, _log.Progress)).ConfigureAwait(true);
            Generated = result;
            StatusMessage = L10n.Format("CertWizard.Status.Generated", result.Thumbprint);
            if (string.IsNullOrWhiteSpace(PfxPath))
            {
                PfxPath = Path.Combine(dir, $"grex365-{result.Thumbprint}.pfx");
            }
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Cert", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
        }
    }

    public async Task ExportPfxAsync(string password)
    {
        if (Generated is null)
        {
            PfxStatus = L10n.Get("CertWizard.Status.GenerateFirst");
            return;
        }
        if (string.IsNullOrEmpty(password))
        {
            PfxStatus = L10n.Get("CertWizard.Status.PasswordRequired");
            return;
        }
        if (string.IsNullOrWhiteSpace(PfxPath))
        {
            PfxStatus = L10n.Get("CertWizard.Status.PfxPathRequired");
            return;
        }

        try
        {
            var result = await Task.Run(() => _generator.ExportPfx(Generated.Thumbprint, PfxPath, password, _log.Progress))
                .ConfigureAwait(true);
            PfxStatus = L10n.Format("CertWizard.Status.PfxOk", result.BytesWritten, result.PfxPath);
        }
        catch (Exception ex)
        {
            PfxStatus = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Cert", ex.Message, ex));
        }
    }

    [RelayCommand]
    private async Task CreateAppRegistrationAsync()
    {
        if (Generated is null)
        {
            AppRegStatus = L10n.Get("CertWizard.Status.GenerateFirst");
            return;
        }
        if (!_graph.IsConnected)
        {
            AppRegStatus = L10n.Get("CertWizard.Status.GraphNotConnected");
            return;
        }
        if (string.IsNullOrWhiteSpace(AppRegDisplayName))
        {
            AppRegStatus = L10n.Get("CertWizard.Status.DisplayNameRequired");
            return;
        }

        IsBusy = true;
        AppRegStatus = L10n.Get("CertWizard.Status.CreatingAppReg");
        try
        {
            var cerBytes = await File.ReadAllBytesAsync(Generated.CerPath).ConfigureAwait(true);
            var result = await _appReg.CreateAndConfigureAsync(
                AppRegDisplayName.Trim(),
                cerBytes,
                Generated.Thumbprint,
                _log.Progress).ConfigureAwait(true);
            AppRegResult = result;

            // Persist CertConfig so Connect by certificate works after admin consent.
            await _certStore.SaveAsync(new CertConfig(
                AppId: result.AppId,
                TenantId: result.TenantId,
                Organization: AppRegDisplayName.Trim(),
                CertThumbprint: Generated.Thumbprint)).ConfigureAwait(true);

            AppRegStatus = L10n.Format("CertWizard.Status.AppRegOk", result.AppId);
            _log.Progress.Report(LogEntry.Ok("AppReg", AppRegStatus));
        }
        catch (Exception ex)
        {
            AppRegStatus = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("AppReg", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private void OpenAdminConsent()
    {
        if (AppRegResult is null)
        {
            return;
        }
        try
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = AppRegResult.AdminConsentUrl,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            _log.Progress.Report(LogEntry.Error("AppReg", ex.Message, ex));
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
