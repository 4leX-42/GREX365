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
    [ObservableProperty] private string _statusMessage = "Completa los campos y pulsa 'Ejecutar onboarding'.";
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
            dispatcher.InvokeAsync(() => _ = TryAutoLoadSkusAsync());
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
        StatusMessage = "Cargando SKUs...";
        try
        {
            var skus = await _users.ListSkusAsync(_cts!.Token).ConfigureAwait(true);
            AvailableSkus.Clear();
            foreach (var s in skus)
            {
                AvailableSkus.Add(s);
            }
            StatusMessage = $"{skus.Count} SKUs disponibles.";
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
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
        summary.Add($"Crear {options.Upn}");
        if (options.SkuIds.Count > 0) summary.Add($"{options.SkuIds.Count} licencia(s)");
        if (options.GroupIdentifiers.Count > 0) summary.Add($"{options.GroupIdentifiers.Count} grupo(s)");
        var ok = await _dialogs.ConfirmAsync(
            $"Onboarding:\n\n  {string.Join(", ", summary)}\n\n¿Continuar?",
            "Confirmar onboarding").ConfigureAwait(true);
        if (!ok)
        {
            StatusMessage = "Cancelado por el usuario.";
            return;
        }

        EnsureToken();
        IsBusy = true;
        RunCommand.NotifyCanExecuteChanged();
        CancelCommand.NotifyCanExecuteChanged();
        StatusMessage = $"Ejecutando onboarding de {options.Upn}...";
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
                ? $"Onboarding OK · {result.Steps.Count} pasos"
                : $"Onboarding con errores · {result.Steps.Count(s => s.Status == "ERROR")} fallos";
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
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
