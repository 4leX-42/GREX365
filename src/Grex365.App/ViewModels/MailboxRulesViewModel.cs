using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.App.ViewModels;

public sealed partial class MailboxRulesViewModel : ObservableObject
{
    private readonly IMailboxRulesService _rules;
    private readonly IUiLogSink _log;
    private CancellationTokenSource? _cts;

    [ObservableProperty] private string _identity = string.Empty;
    [ObservableProperty] private string _statusMessage = "Indica un buzón y pulsa 'Cargar'.";
    [ObservableProperty] private bool _isBusy;

    [ObservableProperty] private AutoReplyState _autoReplyState = AutoReplyState.Disabled;
    [ObservableProperty] private string _internalMessage = string.Empty;
    [ObservableProperty] private string _externalMessage = string.Empty;
    [ObservableProperty] private DateTime _startTime = DateTime.Today.AddDays(1);
    [ObservableProperty] private DateTime _endTime = DateTime.Today.AddDays(8);

    [ObservableProperty] private string _forwardingSmtp = string.Empty;
    [ObservableProperty] private bool _deliverToMailboxAndForward = true;
    [ObservableProperty] private string _currentForwardingDisplay = "(sin configurar)";

    [ObservableProperty] private string _calendarPrincipal = string.Empty;
    [ObservableProperty] private string _calendarAccess = CalendarAccessRights.Reviewer;
    [ObservableProperty] private CalendarPermissionEntry? _selectedCalendarPermission;

    public ObservableCollection<CalendarPermissionEntry> CalendarPermissions { get; } = new();
    public IReadOnlyList<string> CalendarAccessOptions { get; } = CalendarAccessRights.All;

    public AutoReplyState[] AutoReplyStates { get; } =
        new[] { AutoReplyState.Disabled, AutoReplyState.Enabled, AutoReplyState.Scheduled };

    public MailboxRulesViewModel(IMailboxRulesService rules, IUiLogSink log)
    {
        _rules = rules;
        _log = log;
    }

    [RelayCommand]
    private async Task LoadAsync()
    {
        if (string.IsNullOrWhiteSpace(Identity))
        {
            StatusMessage = "Buzón vacío.";
            return;
        }
        EnsureToken();
        IsBusy = true;
        StatusMessage = $"Cargando reglas de {Identity}...";
        try
        {
            var ar = await _rules.GetAutoReplyAsync(Identity.Trim(), _log.Progress, _cts!.Token).ConfigureAwait(true);
            if (ar is not null)
            {
                AutoReplyState = ar.State;
                InternalMessage = ar.InternalMessage ?? string.Empty;
                ExternalMessage = ar.ExternalMessage ?? string.Empty;
                if (ar.StartTime.HasValue) StartTime = ar.StartTime.Value;
                if (ar.EndTime.HasValue) EndTime = ar.EndTime.Value;
            }

            var fwd = await _rules.GetForwardingAsync(Identity.Trim(), _log.Progress, _cts.Token).ConfigureAwait(true);
            if (fwd is not null)
            {
                ForwardingSmtp = fwd.ForwardingSmtpAddress ?? string.Empty;
                DeliverToMailboxAndForward = fwd.DeliverToMailboxAndForward;
                CurrentForwardingDisplay = string.IsNullOrWhiteSpace(fwd.ForwardingSmtpAddress)
                    && string.IsNullOrWhiteSpace(fwd.ForwardingAddress)
                    ? "(sin configurar)"
                    : $"SMTP: {fwd.ForwardingSmtpAddress ?? "—"} · Dir: {fwd.ForwardingAddress ?? "—"} · Deliver: {fwd.DeliverToMailboxAndForward}";
            }

            try
            {
                var cal = await _rules.GetCalendarPermissionsAsync(Identity.Trim(), _log.Progress, _cts.Token).ConfigureAwait(true);
                CalendarPermissions.Clear();
                foreach (var c in cal) CalendarPermissions.Add(c);
            }
            catch (Exception ex)
            {
                _log.Progress.Report(LogEntry.Warn("MailboxRules", "Calendar perms no cargados: " + ex.Message));
            }

            StatusMessage = "Reglas cargadas.";
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("MailboxRules", ex.Message, ex));
        }
        finally
        {
            DisposeToken();
        }
    }

    [RelayCommand]
    private async Task ApplyAutoReplyAsync()
    {
        if (string.IsNullOrWhiteSpace(Identity))
        {
            StatusMessage = "Buzón vacío.";
            return;
        }
        var config = new AutoReplyConfig(
            State: AutoReplyState,
            InternalMessage: string.IsNullOrWhiteSpace(InternalMessage) ? null : InternalMessage,
            ExternalMessage: string.IsNullOrWhiteSpace(ExternalMessage) ? null : ExternalMessage,
            StartTime: AutoReplyState == AutoReplyState.Scheduled ? StartTime : null,
            EndTime: AutoReplyState == AutoReplyState.Scheduled ? EndTime : null);

        EnsureToken();
        IsBusy = true;
        StatusMessage = $"Aplicando AutoReply {AutoReplyState}...";
        try
        {
            await _rules.SetAutoReplyAsync(Identity.Trim(), config, _log.Progress, _cts!.Token).ConfigureAwait(true);
            StatusMessage = $"AutoReply: {AutoReplyState}";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("MailboxRules", ex.Message, ex));
        }
        finally
        {
            DisposeToken();
        }
    }

    [RelayCommand]
    private async Task ApplyForwardingAsync()
    {
        if (string.IsNullOrWhiteSpace(Identity))
        {
            StatusMessage = "Buzón vacío.";
            return;
        }
        var confirm = System.Windows.MessageBox.Show(
            $"Aplicar reenvío de {Identity} hacia {ForwardingSmtp}?\nDeliver a buzón original: {DeliverToMailboxAndForward}",
            "Confirmar reenvío",
            System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Question);
        if (confirm != System.Windows.MessageBoxResult.Yes)
        {
            StatusMessage = "Cancelado por el usuario.";
            return;
        }
        EnsureToken();
        IsBusy = true;
        StatusMessage = $"Aplicando reenvío...";
        try
        {
            await _rules.SetForwardingAsync(Identity.Trim(), ForwardingSmtp.Trim(), DeliverToMailboxAndForward, _log.Progress, _cts!.Token).ConfigureAwait(true);
            CurrentForwardingDisplay = $"SMTP: {ForwardingSmtp} · Deliver: {DeliverToMailboxAndForward}";
            StatusMessage = "Reenvío aplicado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("MailboxRules", ex.Message, ex));
        }
        finally
        {
            DisposeToken();
        }
    }

    [RelayCommand]
    private async Task ClearForwardingAsync()
    {
        if (string.IsNullOrWhiteSpace(Identity))
        {
            StatusMessage = "Buzón vacío.";
            return;
        }
        var confirm = System.Windows.MessageBox.Show(
            $"Quitar reenvío de {Identity}?",
            "Confirmar",
            System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Warning);
        if (confirm != System.Windows.MessageBoxResult.Yes)
        {
            StatusMessage = "Cancelado por el usuario.";
            return;
        }
        EnsureToken();
        IsBusy = true;
        StatusMessage = "Limpiando reenvío...";
        try
        {
            await _rules.ClearForwardingAsync(Identity.Trim(), _log.Progress, _cts!.Token).ConfigureAwait(true);
            ForwardingSmtp = string.Empty;
            DeliverToMailboxAndForward = false;
            CurrentForwardingDisplay = "(sin configurar)";
            StatusMessage = "Reenvío eliminado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("MailboxRules", ex.Message, ex));
        }
        finally
        {
            DisposeToken();
        }
    }

    [RelayCommand]
    private async Task ApplyCalendarPermissionAsync()
    {
        if (string.IsNullOrWhiteSpace(Identity) || string.IsNullOrWhiteSpace(CalendarPrincipal))
        {
            StatusMessage = "Indica buzón y principal.";
            return;
        }
        EnsureToken();
        IsBusy = true;
        StatusMessage = $"Aplicando {CalendarAccess} a {CalendarPrincipal}...";
        try
        {
            await _rules.ApplyCalendarPermissionAsync(Identity.Trim(), CalendarPrincipal.Trim(), CalendarAccess, _log.Progress, _cts!.Token).ConfigureAwait(true);
            var refreshed = await _rules.GetCalendarPermissionsAsync(Identity.Trim(), _log.Progress, _cts.Token).ConfigureAwait(true);
            CalendarPermissions.Clear();
            foreach (var c in refreshed) CalendarPermissions.Add(c);
            StatusMessage = "Permiso calendario aplicado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("MailboxRules", ex.Message, ex));
        }
        finally
        {
            DisposeToken();
        }
    }

    [RelayCommand]
    private async Task RemoveCalendarPermissionAsync()
    {
        if (string.IsNullOrWhiteSpace(Identity) || SelectedCalendarPermission is null)
        {
            StatusMessage = "Selecciona un permiso.";
            return;
        }
        var target = SelectedCalendarPermission;
        var confirm = System.Windows.MessageBox.Show(
            $"Quitar permiso calendario a {target.Principal} ({target.AccessRights})?",
            "Confirmar",
            System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Warning);
        if (confirm != System.Windows.MessageBoxResult.Yes)
        {
            StatusMessage = "Cancelado por el usuario.";
            return;
        }
        EnsureToken();
        IsBusy = true;
        StatusMessage = $"Quitando {target.Principal}...";
        try
        {
            await _rules.RemoveCalendarPermissionAsync(Identity.Trim(), target.Principal, _log.Progress, _cts!.Token).ConfigureAwait(true);
            CalendarPermissions.Remove(target);
            StatusMessage = "Permiso eliminado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("MailboxRules", ex.Message, ex));
        }
        finally
        {
            DisposeToken();
        }
    }

    [RelayCommand(CanExecute = nameof(CanCancel))]
    private void Cancel() => _cts?.Cancel();

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
        CancelCommand.NotifyCanExecuteChanged();
    }
}
