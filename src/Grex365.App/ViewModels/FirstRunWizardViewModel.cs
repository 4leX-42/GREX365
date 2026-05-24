using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.App.ViewModels;

public enum FirstRunStep
{
    Welcome = 0,
    Connection = 1,
    TenantLock = 2,
    Theme = 3,
    Summary = 4
}

public sealed partial class FirstRunWizardViewModel : ObservableObject
{
    private readonly IPreferencesStore _prefs;

    [ObservableProperty] private FirstRunStep _currentStep = FirstRunStep.Welcome;
    [ObservableProperty] private string _connectionMethod = "devicecode";
    [ObservableProperty] private bool _enforceTenantLock;
    [ObservableProperty] private string _expectedTenantId = string.Empty;
    [ObservableProperty] private string _expectedTenantDomain = string.Empty;
    [ObservableProperty] private string _theme = "Dark";
    [ObservableProperty] private bool _isSaving;
    [ObservableProperty] private string _saveStatus = string.Empty;
    [ObservableProperty] private bool _completed;
    [ObservableProperty] private bool _skipped;

    public FirstRunWizardViewModel(IPreferencesStore prefs)
    {
        _prefs = prefs;
    }

    public bool IsWelcome => CurrentStep == FirstRunStep.Welcome;
    public bool IsConnection => CurrentStep == FirstRunStep.Connection;
    public bool IsTenantLock => CurrentStep == FirstRunStep.TenantLock;
    public bool IsTheme => CurrentStep == FirstRunStep.Theme;
    public bool IsSummary => CurrentStep == FirstRunStep.Summary;
    public bool CanGoBack => CurrentStep > FirstRunStep.Welcome;
    public bool CanGoNext => CurrentStep < FirstRunStep.Summary;
    public string ConnectionMethodLabel => string.Equals(ConnectionMethod, "cert", StringComparison.OrdinalIgnoreCase)
        ? L10n.Get("Wizard.Connection.Label.Cert")
        : L10n.Get("Wizard.Connection.Label.DeviceCode");
    public string TenantLockLabel
    {
        get
        {
            if (!EnforceTenantLock) return L10n.Get("Wizard.TenantLock.Label.Disabled");
            var empty = L10n.Get("Wizard.TenantLock.Label.Empty");
            var id = string.IsNullOrWhiteSpace(ExpectedTenantId) ? empty : ExpectedTenantId;
            var domain = string.IsNullOrWhiteSpace(ExpectedTenantDomain) ? empty : ExpectedTenantDomain;
            return L10n.Format("Wizard.TenantLock.Label.Enabled", id, domain);
        }
    }

    partial void OnCurrentStepChanged(FirstRunStep value)
    {
        OnPropertyChanged(nameof(IsWelcome));
        OnPropertyChanged(nameof(IsConnection));
        OnPropertyChanged(nameof(IsTenantLock));
        OnPropertyChanged(nameof(IsTheme));
        OnPropertyChanged(nameof(IsSummary));
        OnPropertyChanged(nameof(CanGoBack));
        OnPropertyChanged(nameof(CanGoNext));
        NextCommand.NotifyCanExecuteChanged();
        BackCommand.NotifyCanExecuteChanged();
    }

    partial void OnConnectionMethodChanged(string value) => OnPropertyChanged(nameof(ConnectionMethodLabel));
    partial void OnEnforceTenantLockChanged(bool value) => OnPropertyChanged(nameof(TenantLockLabel));
    partial void OnExpectedTenantIdChanged(string value) => OnPropertyChanged(nameof(TenantLockLabel));
    partial void OnExpectedTenantDomainChanged(string value) => OnPropertyChanged(nameof(TenantLockLabel));

    [RelayCommand(CanExecute = nameof(CanGoNext))]
    private void Next()
    {
        if (CurrentStep < FirstRunStep.Summary)
        {
            CurrentStep++;
        }
    }

    [RelayCommand(CanExecute = nameof(CanGoBack))]
    private void Back()
    {
        if (CurrentStep > FirstRunStep.Welcome)
        {
            CurrentStep--;
        }
    }

    [RelayCommand]
    private async Task SkipAsync()
    {
        IsSaving = true;
        SaveStatus = L10n.Get("Wizard.Status.Skipping");
        try
        {
            var prefs = await _prefs.LoadAsync().ConfigureAwait(true);
            prefs.FirstRunCompleted = true;
            await _prefs.SaveAsync(prefs).ConfigureAwait(true);
            Skipped = true;
            Completed = true;
        }
        catch (Exception ex)
        {
            SaveStatus = L10n.Format("Settings.SaveStatus.ErrorPrefix", ex.Message);
        }
        finally
        {
            IsSaving = false;
        }
    }

    [RelayCommand]
    private async Task FinishAsync()
    {
        IsSaving = true;
        SaveStatus = L10n.Get("Wizard.Status.Saving");
        try
        {
            var prefs = await _prefs.LoadAsync().ConfigureAwait(true);
            prefs.FirstRunCompleted = true;
            prefs.ConnectionMethod = ConnectionMethod;
            prefs.EnforceTenantLock = EnforceTenantLock;
            prefs.ExpectedTenantId = string.IsNullOrWhiteSpace(ExpectedTenantId) ? null : ExpectedTenantId.Trim();
            prefs.ExpectedTenantDomain = string.IsNullOrWhiteSpace(ExpectedTenantDomain) ? null : ExpectedTenantDomain.Trim();
            prefs.Theme = Theme;
            await _prefs.SaveAsync(prefs).ConfigureAwait(true);
            SaveStatus = L10n.Get("Wizard.Status.Saved");
            Completed = true;
        }
        catch (Exception ex)
        {
            SaveStatus = L10n.Format("Settings.SaveStatus.ErrorPrefix", ex.Message);
        }
        finally
        {
            IsSaving = false;
        }
    }
}
