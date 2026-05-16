using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.App.ViewModels;

public sealed partial class SettingsViewModel : ObservableObject
{
    private readonly IPreferencesStore _prefsStore;
    private readonly ICertConfigStore _certStore;
    private readonly ICertValidator _certValidator;
    private readonly IUiLogSink _log;

    [ObservableProperty] private string _connectionMethod = "cert";
    [ObservableProperty] private string? _expectedTenantId;
    [ObservableProperty] private string? _expectedTenantDomain;
    [ObservableProperty] private bool _enforceTenantLock;

    [ObservableProperty] private string _certAppId = string.Empty;
    [ObservableProperty] private string _certTenantId = string.Empty;
    [ObservableProperty] private string _certOrganization = string.Empty;
    [ObservableProperty] private string _certThumbprint = string.Empty;

    [ObservableProperty] private string _certStatusMessage = "—";
    [ObservableProperty] private bool _certIsValid;

    [ObservableProperty] private string _saveStatus = string.Empty;

    public SettingsViewModel(
        IPreferencesStore prefsStore,
        ICertConfigStore certStore,
        ICertValidator certValidator,
        IUiLogSink log)
    {
        _prefsStore = prefsStore;
        _certStore = certStore;
        _certValidator = certValidator;
        _log = log;
    }

    [RelayCommand]
    private async Task LoadAsync()
    {
        var prefs = await _prefsStore.LoadAsync().ConfigureAwait(true);
        ConnectionMethod = prefs.ConnectionMethod ?? "cert";
        ExpectedTenantId = prefs.ExpectedTenantId;
        ExpectedTenantDomain = prefs.ExpectedTenantDomain;
        EnforceTenantLock = prefs.EnforceTenantLock;

        var cert = await _certStore.LoadAsync().ConfigureAwait(true);
        if (cert is not null)
        {
            CertAppId = cert.AppId;
            CertTenantId = cert.TenantId;
            CertOrganization = cert.Organization;
            CertThumbprint = cert.CertThumbprint;
        }

        ValidateCert();
    }

    [RelayCommand]
    private async Task SaveAsync()
    {
        try
        {
            var prefs = await _prefsStore.LoadAsync().ConfigureAwait(true);
            prefs.ConnectionMethod = ConnectionMethod;
            prefs.ExpectedTenantId = ExpectedTenantId;
            prefs.ExpectedTenantDomain = ExpectedTenantDomain;
            prefs.EnforceTenantLock = EnforceTenantLock;
            await _prefsStore.SaveAsync(prefs).ConfigureAwait(true);

            if (!string.IsNullOrWhiteSpace(CertAppId)
                && !string.IsNullOrWhiteSpace(CertTenantId)
                && !string.IsNullOrWhiteSpace(CertOrganization)
                && !string.IsNullOrWhiteSpace(CertThumbprint))
            {
                await _certStore.SaveAsync(new CertConfig(
                    CertAppId, CertTenantId, CertOrganization, CertThumbprint))
                    .ConfigureAwait(true);
            }

            SaveStatus = $"Guardado · {DateTime.Now:HH:mm:ss}";
            _log.Progress.Report(LogEntry.Ok("Settings", "Preferencias y certificado guardados."));
            ValidateCert();
        }
        catch (Exception ex)
        {
            SaveStatus = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("Settings", ex.Message, ex));
        }
    }

    [RelayCommand]
    private void ValidateCert()
    {
        var cfg = new CertConfig(
            CertAppId ?? string.Empty,
            CertTenantId ?? string.Empty,
            CertOrganization ?? string.Empty,
            CertThumbprint ?? string.Empty);
        var r = _certValidator.Validate(cfg);
        CertStatusMessage = r.Message;
        CertIsValid = r.IsValid;
    }
}
