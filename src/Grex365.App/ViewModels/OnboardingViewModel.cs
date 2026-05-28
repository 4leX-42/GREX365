using System.Collections.ObjectModel;
using System.ComponentModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.App.ViewModels;

public sealed partial class OnboardingViewModel : ObservableObject
{
    private readonly IOnboardingService _onboarding;
    private readonly IUsersService _users;
    private readonly IUiLogSink _log;
    private readonly IDialogService _dialogs;
    private readonly IConnectionStateMonitor? _monitor;
    private CancellationTokenSource? _cts;

    [ObservableProperty] private string _displayName = string.Empty;
    [ObservableProperty] private string _upn = string.Empty;
    [ObservableProperty] private string _mailNickname = string.Empty;
    [ObservableProperty] private string _initialPassword = string.Empty;
    [ObservableProperty] private string _usageLocation = "ES";
    [ObservableProperty] private bool _forceChangePassword = true;
    [ObservableProperty] private string _groupsText = string.Empty;
    [ObservableProperty] private SkuInfo? _selectedSku;
    [ObservableProperty] private string _statusMessage = L10n.Get("Onboarding.Status.Initial");
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private OnboardingResult? _result;

    public ObservableCollection<OnboardingStep> Steps { get; } = new();
    public ObservableCollection<SkuInfo> AvailableSkus { get; } = new();
    public ObservableCollection<SkuInfo> SelectedSkus { get; } = new();

    public OnboardingViewModel(IOnboardingService onboarding, IUsersService users, IUiLogSink log, IDialogService dialogs, IConnectionStateMonitor? monitor = null)
    {
        _onboarding = onboarding;
        _users = users;
        _log = log;
        _dialogs = dialogs;
        _monitor = monitor;
        if (_monitor is not null)
        {
            _monitor.PropertyChanged += OnMonitorChanged;
            if (_monitor.Current.GraphConnected)
            {
                DispatchAutoLoad();
            }
        }
    }

    private void OnMonitorChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName != nameof(IConnectionStateMonitor.Current)) return;
        if (_monitor is null || !_monitor.Current.GraphConnected) return;
        // ConnectionStateMonitor fires from its 1s background poll loop — marshal
        // to UI dispatcher before mutating ObservableCollection (AvailableSkus is
        // bound to a ComboBox CollectionView) or triggering AsyncRelayCommand.
        DispatchAutoLoad();
    }

    private void DispatchAutoLoad()
    {
        var dispatcher = System.Windows.Application.Current?.Dispatcher;
        if (dispatcher is null || dispatcher.CheckAccess())
        {
            _ = TryAutoLoadSkusAsync();
        }
        else
        {
            // Func<Task> overload — keeps the async chain on the dispatcher so collection
            // mutations + AsyncRelayCommand completion stay on the UI thread (the Action
            // overload fire-and-forgets the inner Task → cross-thread CollectionView crash).
            dispatcher.InvokeAsync(TryAutoLoadSkusAsync);
        }
    }

    private async Task TryAutoLoadSkusAsync()
    {
        if (AvailableSkus.Count > 0 || IsBusy) return;
        try
        {
            await LoadSkusCommand.ExecuteAsync(null).ConfigureAwait(true);
        }
        catch
        {
            // LoadSkusAsync surfaces errors via StatusMessage + log sink.
        }
    }

    [RelayCommand]
    private async Task LoadSkusAsync()
    {
        EnsureToken();
        IsBusy = true;
        StatusMessage = L10n.Get("Users.Status.LoadingSkus");
        try
        {
            var skus = await _users.ListSkusAsync(_cts!.Token).ConfigureAwait(true);
            AvailableSkus.Clear();
            foreach (var s in skus)
            {
                AvailableSkus.Add(s);
            }
            StatusMessage = L10n.Format("Users.Status.SkusAvailable", skus.Count);
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Onboarding", ex.Message, ex));
        }
        finally
        {
            DisposeToken();
        }
    }

    [RelayCommand]
    private void AddSelectedSku()
    {
        if (SelectedSku is null) return;
        if (!SelectedSkus.Any(s => s.SkuId == SelectedSku.SkuId))
        {
            SelectedSkus.Add(SelectedSku);
        }
    }

    [RelayCommand]
    private void RemoveSku(SkuInfo? sku)
    {
        if (sku is null) return;
        SelectedSkus.Remove(sku);
    }

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunAsync()
    {
        var groupKeys = (GroupsText ?? string.Empty)
            .Split(new[] { '\n', '\r', ',', ';' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        var options = new OnboardingOptions(
            DisplayName: DisplayName?.Trim() ?? string.Empty,
            Upn: Upn?.Trim() ?? string.Empty,
            InitialPassword: InitialPassword ?? string.Empty,
            UsageLocation: UsageLocation?.Trim().ToUpperInvariant() ?? string.Empty,
            MailNickname: string.IsNullOrWhiteSpace(MailNickname) ? null : MailNickname.Trim(),
            SkuIds: SelectedSkus.Select(s => s.SkuId).ToList(),
            GroupIdentifiers: groupKeys,
            ForceChangePasswordNextSignIn: ForceChangePassword);

        var summary = new List<string>();
        summary.Add(L10n.Format("Onboarding.Confirm.PartCreate", options.Upn));
        if (options.SkuIds.Count > 0) summary.Add(L10n.Format("Onboarding.Confirm.PartLicenses", options.SkuIds.Count));
        if (options.GroupIdentifiers.Count > 0) summary.Add(L10n.Format("Onboarding.Confirm.PartGroups", options.GroupIdentifiers.Count));
        var ok = await _dialogs.ConfirmAsync(
            L10n.Format("Onboarding.Confirm.Body", string.Join(", ", summary)),
            L10n.Get("Onboarding.Confirm.Title")).ConfigureAwait(true);
        if (!ok)
        {
            StatusMessage = L10n.Get("Common.Status.CancelledByUser");
            return;
        }

        EnsureToken();
        IsBusy = true;
        RunCommand.NotifyCanExecuteChanged();
        CancelCommand.NotifyCanExecuteChanged();
        StatusMessage = L10n.Format("Onboarding.Status.Running", options.Upn);
        Steps.Clear();
        Result = null;

        try
        {
            var result = await _onboarding.RunAsync(options, _log.Progress, _cts!.Token).ConfigureAwait(true);
            Result = result;
            foreach (var step in result.Steps)
            {
                Steps.Add(step);
            }
            StatusMessage = result.Success
                ? L10n.Format("Onboarding.Status.SuccessSummary", result.Steps.Count)
                : L10n.Format("Onboarding.Status.ErrorSummary", result.Steps.Count(s => s.Status == "ERROR"));
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Onboarding", ex.Message, ex));
        }
        finally
        {
            DisposeToken();
            RunCommand.NotifyCanExecuteChanged();
            CancelCommand.NotifyCanExecuteChanged();
        }
    }

    [RelayCommand(CanExecute = nameof(CanCancel))]
    private void Cancel() => _cts?.Cancel();

    private bool CanRun() => !IsBusy;
    private bool CanCancel() => IsBusy;

    private void EnsureToken()
    {
        _cts?.Dispose();
        _cts = new CancellationTokenSource();
    }

    private void DisposeToken()
    {
        IsBusy = false;
        _cts?.Dispose();
        _cts = null;
    }
}
