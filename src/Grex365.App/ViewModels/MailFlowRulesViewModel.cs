using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.App.ViewModels;

public sealed partial class MailFlowRulesViewModel : ObservableObject
{
    private readonly IMailFlowRulesService _service;
    private readonly IUiLogSink _log;
    private readonly IExchangeConnection _exo;

    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private string _statusMessage = L10n.Get("MailFlow.Status.Initial");
    [ObservableProperty] private string _filter = string.Empty;

    private readonly List<TransportRuleSummary> _all = new();
    public ObservableCollection<TransportRuleSummary> Filtered { get; } = new();

    public MailFlowRulesViewModel(
        IMailFlowRulesService service,
        IUiLogSink log,
        IExchangeConnection exo)
    {
        _service = service;
        _log = log;
        _exo = exo;
    }

    [RelayCommand]
    private async Task LoadAsync()
    {
        if (!_exo.IsConnected)
        {
            StatusMessage = L10n.Get("MailFlow.Status.NotConnected");
            return;
        }

        IsBusy = true;
        StatusMessage = L10n.Get("MailFlow.Status.Loading");
        try
        {
            var rules = await _service.GetRulesAsync(_log.Progress).ConfigureAwait(true);
            _all.Clear();
            _all.AddRange(rules);
            ApplyFilter();
            StatusMessage = L10n.Format("MailFlow.Status.Summary", rules.Count, Filtered.Count);
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("MailFlow", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
        }
    }

    partial void OnFilterChanged(string value) => ApplyFilter();

    private void ApplyFilter()
    {
        Filtered.Clear();
        IEnumerable<TransportRuleSummary> q = _all;
        if (!string.IsNullOrWhiteSpace(Filter))
        {
            q = q.Where(r =>
                r.Name.Contains(Filter, StringComparison.OrdinalIgnoreCase) ||
                (r.Description?.Contains(Filter, StringComparison.OrdinalIgnoreCase) ?? false));
        }
        foreach (var r in q)
        {
            Filtered.Add(r);
        }
    }
}
