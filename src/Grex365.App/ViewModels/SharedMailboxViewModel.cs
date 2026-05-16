using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.App.ViewModels;

public sealed partial class SharedMailboxViewModel : ObservableObject
{
    private readonly ISharedMailboxService _service;
    private readonly IUiLogSink _log;

    [ObservableProperty] private string _mailboxIdentity = string.Empty;
    [ObservableProperty] private MailboxInfo? _mailboxInfo;
    [ObservableProperty] private string _statusMessage = string.Empty;
    [ObservableProperty] private bool _isBusy;

    [ObservableProperty] private string _permPrincipal = string.Empty;
    [ObservableProperty] private string _permAction = "add";
    [ObservableProperty] private string _permPermission = "FullAccess";

    public ObservableCollection<MailboxPermissionResult> PermissionResults { get; } = new();

    public SharedMailboxViewModel(ISharedMailboxService service, IUiLogSink log)
    {
        _service = service;
        _log = log;
    }

    [RelayCommand]
    private async Task LookupAsync()
    {
        if (string.IsNullOrWhiteSpace(MailboxIdentity))
        {
            StatusMessage = "Indica un buzón.";
            return;
        }
        IsBusy = true;
        StatusMessage = "Consultando buzón...";
        try
        {
            MailboxInfo = await _service.GetMailboxAsync(MailboxIdentity.Trim(), _log.Progress).ConfigureAwait(true);
            StatusMessage = MailboxInfo is null ? "Buzón no encontrado." : $"Tipo: {MailboxInfo.RecipientTypeDetails}";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("Mailbox", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private async Task ConvertToRegularAsync()
    {
        if (string.IsNullOrWhiteSpace(MailboxIdentity))
        {
            StatusMessage = "Indica un buzón.";
            return;
        }
        IsBusy = true;
        StatusMessage = "Convirtiendo a UserMailbox...";
        try
        {
            MailboxInfo = await _service.ConvertToRegularAsync(MailboxIdentity.Trim(), _log.Progress).ConfigureAwait(true);
            StatusMessage = MailboxInfo is null
                ? "Sin confirmación."
                : $"Resultado: {MailboxInfo.RecipientTypeDetails}";
            _log.Progress.Report(LogEntry.Ok("Mailbox", StatusMessage));
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("Mailbox", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private async Task ApplyPermissionAsync()
    {
        if (string.IsNullOrWhiteSpace(MailboxIdentity) || string.IsNullOrWhiteSpace(PermPrincipal))
        {
            StatusMessage = "Falta buzón o principal.";
            return;
        }
        IsBusy = true;
        StatusMessage = $"{PermAction} {PermPermission} {MailboxIdentity} ↔ {PermPrincipal}...";
        try
        {
            var r = await _service.ApplyPermissionAsync(
                PermAction, PermPermission, MailboxIdentity.Trim(), PermPrincipal.Trim(), _log.Progress)
                .ConfigureAwait(true);
            PermissionResults.Insert(0, r);
            StatusMessage = $"{r.Status}: {r.Detail}";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("Mailbox", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
        }
    }
}
