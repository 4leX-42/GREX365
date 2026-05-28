using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Csv;
using Grex365.Core.Models;
using Microsoft.Win32;

namespace Grex365.App.ViewModels;

public sealed partial class SharedMailboxViewModel : ObservableObject
{
    private readonly ISharedMailboxService _service;
    private readonly IUiLogSink _log;
    private readonly IRbacGuard _rbac;

    [ObservableProperty] private string _mailboxIdentity = string.Empty;
    [ObservableProperty] private MailboxInfo? _mailboxInfo;
    [ObservableProperty] private string _statusMessage = string.Empty;
    [ObservableProperty] private bool _isBusy;

    [ObservableProperty] private string _permPrincipal = string.Empty;
    [ObservableProperty] private string _permAction = "add";
    [ObservableProperty] private string _permPermission = "FullAccess";

    public ObservableCollection<MailboxPermissionResult> PermissionResults { get; } = new();
    public ObservableCollection<MailboxPermissionEntry> CurrentPermissions { get; } = new();

    public SharedMailboxViewModel(ISharedMailboxService service, IUiLogSink log, IRbacGuard rbac)
    {
        _service = service;
        _log = log;
        _rbac = rbac;
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
    private async Task LookupAsync()
    {
        if (string.IsNullOrWhiteSpace(MailboxIdentity))
        {
            StatusMessage = L10n.Get("SharedMailbox.Status.EnterMailbox");
            return;
        }
        IsBusy = true;
        StatusMessage = L10n.Get("SharedMailbox.Status.Querying");
        try
        {
            MailboxInfo = await _service.GetMailboxAsync(MailboxIdentity.Trim(), _log.Progress).ConfigureAwait(true);
            CurrentPermissions.Clear();
            if (MailboxInfo is not null)
            {
                try
                {
                    var perms = await _service.GetPermissionsAsync(MailboxIdentity.Trim(), _log.Progress).ConfigureAwait(true);
                    foreach (var p in perms)
                    {
                        CurrentPermissions.Add(p);
                    }
                    StatusMessage = L10n.Format("SharedMailbox.Status.LookupSummary", MailboxInfo.RecipientTypeDetails, perms.Count);
                }
                catch (Exception permEx)
                {
                    StatusMessage = L10n.Format("SharedMailbox.Status.LookupPermsError", MailboxInfo.RecipientTypeDetails, permEx.Message);
                }
            }
            else
            {
                StatusMessage = L10n.Get("SharedMailbox.Status.NotFound");
            }
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
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
            StatusMessage = L10n.Get("SharedMailbox.Status.EnterMailbox");
            return;
        }
        if (!await RequireAuthorizedAsync("Convert mailbox").ConfigureAwait(true)) return;
        IsBusy = true;
        StatusMessage = L10n.Get("SharedMailbox.Status.Converting");
        try
        {
            MailboxInfo = await _service.ConvertToRegularAsync(MailboxIdentity.Trim(), _log.Progress).ConfigureAwait(true);
            StatusMessage = MailboxInfo is null
                ? L10n.Get("SharedMailbox.Status.NoConfirmation")
                : L10n.Format("SharedMailbox.Status.ConvertResult", MailboxInfo.RecipientTypeDetails);
            _log.Progress.Report(LogEntry.Ok("Mailbox", StatusMessage));
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Mailbox", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private async Task ImportPermissionsCsvAsync()
    {
        var dlg = new OpenFileDialog
        {
            Title = L10n.Get("SharedMailbox.Dialog.BulkCsv"),
            Filter = "CSV (*.csv)|*.csv|Todos|*.*",
            CheckFileExists = true
        };
        if (dlg.ShowDialog() != true)
        {
            return;
        }

        if (!await RequireAuthorizedAsync("Bulk mailbox perms").ConfigureAwait(true)) return;

        IsBusy = true;
        StatusMessage = L10n.Format("SharedMailbox.Status.Processing", Path.GetFileName(dlg.FileName));
        try
        {
            var rows = FlexibleCsvReader.Read(dlg.FileName);
            if (rows.Count == 0)
            {
                StatusMessage = L10n.Get("SharedMailbox.Status.CsvEmpty");
                return;
            }

            var ok = 0; var err = 0; var inv = 0;
            foreach (var row in rows)
            {
                row.TryGetValue("Action", out var action);
                row.TryGetValue("Permission", out var permission);
                row.TryGetValue("Mailbox", out var mailbox);
                row.TryGetValue("Principal", out var principal);

                var r = await _service.ApplyPermissionAsync(
                    action ?? string.Empty,
                    permission ?? string.Empty,
                    mailbox ?? string.Empty,
                    principal ?? string.Empty,
                    _log.Progress).ConfigureAwait(true);
                PermissionResults.Insert(0, r);
                switch (r.Status)
                {
                    case "OK": ok++; break;
                    case "INVALIDO": inv++; break;
                    default: err++; break;
                }
            }

            StatusMessage = L10n.Format("SharedMailbox.Status.CsvSummary", ok, inv, err);
            _log.Progress.Report(LogEntry.Ok("Mailbox", StatusMessage));
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Mailbox", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private void ExportPermissionResults()
    {
        if (PermissionResults.Count == 0)
        {
            StatusMessage = L10n.Get("Common.Status.NoResultsToExport");
            return;
        }

        var dlg = new SaveFileDialog
        {
            Title = L10n.Get("Common.Dialog.SaveResults"),
            Filter = "CSV (*.csv)|*.csv",
            FileName = $"mailbox_permissions_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };
        if (dlg.ShowDialog() != true)
        {
            return;
        }

        try
        {
            var sb = new StringBuilder();
            sb.AppendLine("Action,Permission,Mailbox,Principal,Status,Detail");
            foreach (var r in PermissionResults)
            {
                sb.Append(Escape(r.Action)).Append(',');
                sb.Append(Escape(r.Permission)).Append(',');
                sb.Append(Escape(r.Mailbox)).Append(',');
                sb.Append(Escape(r.Principal)).Append(',');
                sb.Append(Escape(r.Status)).Append(',');
                sb.Append(Escape(r.Detail)).AppendLine();
            }
            File.WriteAllText(dlg.FileName, sb.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
            StatusMessage = L10n.Format("Common.Status.Exported", Path.GetFileName(dlg.FileName));
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Mailbox", ex.Message, ex));
        }
    }

    private static string Escape(string? value) => Grex365.Core.Csv.CsvEscaper.Escape(value);

    [RelayCommand]
    private async Task ApplyPermissionAsync()
    {
        if (string.IsNullOrWhiteSpace(MailboxIdentity) || string.IsNullOrWhiteSpace(PermPrincipal))
        {
            StatusMessage = L10n.Get("SharedMailbox.Status.MissingMailboxOrPrincipal");
            return;
        }
        if (!await RequireAuthorizedAsync("Apply mailbox perm").ConfigureAwait(true)) return;
        IsBusy = true;
        StatusMessage = L10n.Format("SharedMailbox.Status.ApplyingPerm", PermAction, PermPermission, MailboxIdentity, PermPrincipal);
        try
        {
            var r = await _service.ApplyPermissionAsync(
                PermAction, PermPermission, MailboxIdentity.Trim(), PermPrincipal.Trim(), _log.Progress)
                .ConfigureAwait(true);
            PermissionResults.Insert(0, r);
            StatusMessage = L10n.Format("SharedMailbox.Status.ApplyResult", r.Status, r.Detail);
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Mailbox", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
        }
    }
}
