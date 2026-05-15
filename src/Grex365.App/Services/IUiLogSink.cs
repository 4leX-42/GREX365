using System.Collections.ObjectModel;
using Grex365.Core.Models;

namespace Grex365.App.Services;

public interface IUiLogSink
{
    ObservableCollection<LogEntry> Entries { get; }

    IProgress<LogEntry> Progress { get; }

    void Clear();
}
