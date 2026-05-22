using System.ComponentModel;

namespace Grex365.Core.Abstractions;

public interface IAuditFindingsStore : INotifyPropertyChanged
{
    DateTimeOffset? LastRunAt { get; }
    string? LastAuditName { get; }
    int ErrorCount { get; }
    int WarnCount { get; }
    int InfoCount { get; }

    void Update(string auditName, int errorCount, int warnCount, int infoCount);
}
