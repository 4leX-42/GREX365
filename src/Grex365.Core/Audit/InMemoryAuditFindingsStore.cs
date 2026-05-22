using System.ComponentModel;
using Grex365.Core.Abstractions;

namespace Grex365.Core.Audit;

public sealed class InMemoryAuditFindingsStore : IAuditFindingsStore
{
    public DateTimeOffset? LastRunAt { get; private set; }
    public string? LastAuditName { get; private set; }
    public int ErrorCount { get; private set; }
    public int WarnCount { get; private set; }
    public int InfoCount { get; private set; }

    public event PropertyChangedEventHandler? PropertyChanged;

    public void Update(string auditName, int errorCount, int warnCount, int infoCount)
    {
        LastAuditName = auditName;
        ErrorCount = errorCount;
        WarnCount = warnCount;
        InfoCount = infoCount;
        LastRunAt = DateTimeOffset.UtcNow;
        Raise(nameof(LastAuditName));
        Raise(nameof(ErrorCount));
        Raise(nameof(WarnCount));
        Raise(nameof(InfoCount));
        Raise(nameof(LastRunAt));
    }

    private void Raise(string name) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
