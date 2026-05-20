using Grex365.Core.Abstractions;

namespace Grex365.Core.Audit;

public sealed record AuditMetrics(
    int TotalCount,
    int OkCount,
    int WarnCount,
    int ErrorCount,
    double ErrorRate,
    int Last24hCount,
    IReadOnlyList<SourceCount> TopSources,
    IReadOnlyList<AuditRecord> RecentErrors);

public sealed record SourceCount(string Source, int Count);

public static class MetricsAggregator
{
    public static AuditMetrics Compute(
        IEnumerable<AuditRecord> records,
        DateTimeOffset? now = null,
        int topSources = 5,
        int recentErrors = 5)
    {
        ArgumentNullException.ThrowIfNull(records);
        var list = records as IReadOnlyCollection<AuditRecord> ?? records.ToList();

        var ok = 0;
        var warn = 0;
        var error = 0;
        var last24h = 0;
        var threshold = (now ?? DateTimeOffset.Now).AddHours(-24);

        foreach (var r in list)
        {
            switch (r.Outcome?.Trim().ToUpperInvariant())
            {
                case "OK": ok++; break;
                case "WARN":
                case "WARNING": warn++; break;
                case "ERROR":
                case "ERR":
                case "FATAL": error++; break;
            }
            if (r.Timestamp >= threshold)
            {
                last24h++;
            }
        }

        var total = list.Count;
        var errorRate = total == 0 ? 0.0 : Math.Round((double)error / total, 4);

        var top = list
            .GroupBy(r => string.IsNullOrWhiteSpace(r.Source) ? "(sin source)" : r.Source)
            .Select(g => new SourceCount(g.Key, g.Count()))
            .OrderByDescending(s => s.Count)
            .ThenBy(s => s.Source, StringComparer.OrdinalIgnoreCase)
            .Take(Math.Max(0, topSources))
            .ToList();

        var recent = list
            .Where(r => string.Equals(r.Outcome, "ERROR", StringComparison.OrdinalIgnoreCase)
                     || string.Equals(r.Outcome, "ERR", StringComparison.OrdinalIgnoreCase)
                     || string.Equals(r.Outcome, "FATAL", StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(r => r.Timestamp)
            .Take(Math.Max(0, recentErrors))
            .ToList();

        return new AuditMetrics(total, ok, warn, error, errorRate, last24h, top, recent);
    }
}
