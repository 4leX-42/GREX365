using Grex365.Core.Abstractions;

namespace Grex365.Core.Audit;

public sealed record RecentActivity(
    int TodayCount,
    int TodayErrors,
    IReadOnlyList<AuditRecord> Recent);

// Pure: "what happened today" summary for the Dashboard — counts today's ops (local time)
// and surfaces the most recent N records (any outcome, newest first).
public static class RecentActivityAggregator
{
    public static RecentActivity Compute(
        IEnumerable<AuditRecord> records,
        DateTimeOffset? now = null,
        int take = 5)
    {
        ArgumentNullException.ThrowIfNull(records);
        var list = records as IReadOnlyCollection<AuditRecord> ?? records.ToList();

        var today = (now ?? DateTimeOffset.Now).LocalDateTime.Date;
        var todayCount = 0;
        var todayErrors = 0;
        foreach (var r in list)
        {
            if (r.Timestamp.LocalDateTime.Date != today) continue;
            todayCount++;
            switch (r.Outcome?.Trim().ToUpperInvariant())
            {
                case "ERROR":
                case "ERR":
                case "FATAL":
                    todayErrors++;
                    break;
            }
        }

        var recent = list
            .OrderByDescending(r => r.Timestamp)
            .Take(Math.Max(0, take))
            .ToList();

        return new RecentActivity(todayCount, todayErrors, recent);
    }
}
