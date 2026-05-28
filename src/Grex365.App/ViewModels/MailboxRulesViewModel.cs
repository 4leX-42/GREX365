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
    private readonly IRbacGuard _rbac;
    private readonly IDialogService _dialogs;
    private CancellationTokenSource? _cts;

    [ObservableProperty] private string _identity = string.Empty;
    [ObservableProperty] private string _statusMessage = L10n.Get("MailboxRules.Status.Initial");
    [ObservableProperty] private bool _isBusy;

    [ObservableProperty] private AutoReplyState _autoReplyState = AutoReplyState.Disabled;
    [ObservableProperty] private string _internalMessage = string.Empty;
    [ObservableProperty] private string _externalMessage = string.Empty;
    [ObservableProperty] private DateTime _startTime = DateTime.Today.AddDays(1);
    [ObservableProperty] private DateTime _endTime = DateTime.Today.AddDays(8);

    [ObservableProperty] private string _forwardingSmtp = string.Empty;
    [ObservableProperty] private bool _deliverToMailboxAndForward = true;
    [ObservableProperty] private string _currentForwardingDisplay = L10n.Get("MailboxRules.Display.NotConfigured");

    [ObservableProperty] private string _calendarPrincipal = string.Empty;
    [ObservableProperty] private string _calendarAccess = CalendarAccessRights.Reviewer;
    [ObservableProperty] private CalendarPermissionEntry? _selectedCalendarPermission;

    // True once a mailbox has been successfully loaded — gates the rule-editing
    // cards (auto-reply / forwarding / calendar) so they stay hidden until there
    // is an actual target mailbox to act on.
    [ObservableProperty] private bool _rulesLoaded;

    public ObservableCollection<CalendarPermissionEntry> CalendarPermissions { get; } = new();
    public IReadOnlyList<string> CalendarAccessOptions { get; } = CalendarAccessRights.All;

    public AutoReplyState[] AutoReplyStates { get; } =
        new[] { AutoReplyState.Disabled, AutoReplyState.Enabled, AutoReplyState.Scheduled };

    public MailboxRulesViewModel(IMailboxRulesService rules, IUiLogSink log, IRbacGuard rbac, IDialogService dialogs)
    {
        _rules = rules;
        _log = log;
        _rbac = rbac;
        _dialogs = dialogs;
    }

    private async Task<bool> RequireAuthorizedAsync(string contextName)
    {
        var decision = await _rbac.EvaluateAsync().ConfigureAwait(true);
        if (decision.Allowed)
        {
            return true;
        }
        StatusMessage = decision.Reason;
        _log.Progress.Report(LogEntry.Warn("RBAC", $"{contextName} bloqueado: {decision.Reason}"));
        return false;
    }

    [RelayCommand]
    private async Task LoadAsync()
    {
        if (string.IsNullOrWhiteSpace(Identity))
        {
            StatusMessage = L10n.Get("MailboxRules.Status.EmptyMailbox");
            return;
        }
        EnsureToken();
        IsBusy = true;
        StatusMessage = L10n.Format("MailboxRules.Status.Loading", Identity);
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
                    ? L10n.Get("MailboxRules.Display.NotConfigured")
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

            RulesLoaded = true;
            StatusMessage = L10n.Get("MailboxRules.Status.RulesLoaded");
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
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
            StatusMessage = L10n.Get("MailboxRules.Status.EmptyMailbox");
            return;
        }
        if (!await RequireAuthorizedAsync("Apply AutoReply").ConfigureAwait(true)) return;
        var config = new AutoReplyConfig(
            State: AutoReplyState,
            InternalMessage: string.IsNullOrWhiteSpace(InternalMessage) ? null : InternalMessage,
            ExternalMessage: string.IsNullOrWhiteSpace(ExternalMessage) ? null : ExternalMessage,
            StartTime: AutoReplyState == AutoReplyState.Scheduled ? StartTime : null,
            EndTime: AutoReplyState == AutoReplyState.Scheduled ? EndTime : null);

        EnsureToken();
        IsBusy = true;
        StatusMessage = L10n.Format("MailboxRules.Status.ApplyingAutoReply", AutoReplyState);
        try
        {
            await _rules.SetAutoReplyAsync(Identity.Trim(), config, _log.Progress, _cts!.Token).ConfigureAwait(true);
            StatusMessage = L10n.Format("MailboxRules.Status.AutoReplyResult", AutoReplyState);
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
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
            StatusMessage = L10n.Get("MailboxRules.Status.EmptyMailbox");
            return;
        }
        if (!await RequireAuthorizedAsync("Apply forwarding").ConfigureAwait(true)) return;
        var ok = await _dialogs.ConfirmAsync(
            L10n.Format("MailboxRules.Confirm.ForwardingBody", Identity, ForwardingSmtp, DeliverToMailboxAndForward),
            L10n.Get("MailboxRules.Confirm.ForwardingTitle")).ConfigureAwait(true);
        if (!ok)
        {
            StatusMessage = L10n.Get("Common.Status.CancelledByUser");
            return;
        }
        EnsureToken();
        IsBusy = true;
        StatusMessage = L10n.Get("MailboxRules.Status.ApplyingForwarding");
        try
        {
            await _rules.SetForwardingAsync(Identity.Trim(), ForwardingSmtp.Trim(), DeliverToMailboxAndForward, _log.Progress, _cts!.Token).ConfigureAwait(true);
            CurrentForwardingDisplay = $"SMTP: {ForwardingSmtp} · Deliver: {DeliverToMailboxAndForward}";
            StatusMessage = L10n.Get("MailboxRules.Status.ForwardingApplied");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
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
            StatusMessage = L10n.Get("MailboxRules.Status.EmptyMailbox");
            return;
        }
        if (!await RequireAuthorizedAsync("Clear forwarding").ConfigureAwait(true)) return;
        var ok = await _dialogs.ConfirmAsync(
            L10n.Format("MailboxRules.Confirm.ClearForwardingBody", Identity),
            L10n.Get("Common.Confirm.Title"),
            DialogIcon.Warning).ConfigureAwait(true);
        if (!ok)
        {
            StatusMessage = L10n.Get("Common.Status.CancelledByUser");
            return;
        }
        EnsureToken();
        IsBusy = true;
        StatusMessage = L10n.Get("MailboxRules.Status.ClearingForwarding");
        try
        {
            await _rules.ClearForwardingAsync(Identity.Trim(), _log.Progress, _cts!.Token).ConfigureAwait(true);
            ForwardingSmtp = string.Empty;
            DeliverToMailboxAndForward = false;
            CurrentForwardingDisplay = L10n.Get("MailboxRules.Display.NotConfigured");
            StatusMessage = L10n.Get("MailboxRules.Status.ForwardingCleared");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
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
            StatusMessage = L10n.Get("MailboxRules.Status.EnterMailboxAndPrincipal");
            return;
        }
        if (!await RequireAuthorizedAsync("Apply calendar perm").ConfigureAwait(true)) return;
        EnsureToken();
        IsBusy = true;
        StatusMessage = L10n.Format("MailboxRules.Status.ApplyingCalendar", CalendarAccess, CalendarPrincipal);
        try
        {
            await _rules.ApplyCalendarPermissionAsync(Identity.Trim(), CalendarPrincipal.Trim(), CalendarAccess, _log.Progress, _cts!.Token).ConfigureAwait(true);
            var refreshed = await _rules.GetCalendarPermissionsAsync(Identity.Trim(), _log.Progress, _cts.Token).ConfigureAwait(true);
            CalendarPermissions.Clear();
            foreach (var c in refreshed) CalendarPermissions.Add(c);
            StatusMessage = L10n.Get("MailboxRules.Status.CalendarApplied");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
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
            StatusMessage = L10n.Get("MailboxRules.Status.SelectPermission");
            return;
        }
        if (!await RequireAuthorizedAsync("Remove calendar perm").ConfigureAwait(true)) return;
        var target = SelectedCalendarPermission;
        var ok = await _dialogs.ConfirmAsync(
            L10n.Format("MailboxRules.Confirm.RemoveCalendarBody", target.Principal, target.AccessRights),
            L10n.Get("Common.Confirm.Title"),
            DialogIcon.Warning).ConfigureAwait(true);
        if (!ok)
        {
            StatusMessage = L10n.Get("Common.Status.CancelledByUser");
            return;
        }
        EnsureToken();
        IsBusy = true;
        StatusMessage = L10n.Format("MailboxRules.Status.RemovingPermission", target.Principal);
        try
        {
            await _rules.RemoveCalendarPermissionAsync(Identity.Trim(), target.Principal, _log.Progress, _cts!.Token).ConfigureAwait(true);
            CalendarPermissions.Remove(target);
            StatusMessage = L10n.Get("MailboxRules.Status.PermissionRemoved");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
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
