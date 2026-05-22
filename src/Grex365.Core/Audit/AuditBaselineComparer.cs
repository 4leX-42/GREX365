using Grex365.Core.Models;

namespace Grex365.Core.Audit;

public static class AuditBaselineComparer
{
    public static AuditBaselineDiff Compare(
        IEnumerable<AuditFinding> baseline,
        IEnumerable<AuditFinding> current)
    {
        var baselineSet = (baseline ?? Array.Empty<AuditFinding>()).ToHashSet(FindingKeyComparer.Instance);
        var currentList = (current ?? Array.Empty<AuditFinding>()).ToList();

        var newFindings = new List<AuditFinding>();
        var persistent = new List<AuditFinding>();
        var seenInCurrent = new HashSet<AuditFinding>(FindingKeyComparer.Instance);

        foreach (var f in currentList)
        {
            if (!seenInCurrent.Add(f))
            {
                continue;
            }
            if (baselineSet.Contains(f))
            {
                persistent.Add(f);
            }
            else
            {
                newFindings.Add(f);
            }
        }

        var resolved = baselineSet
            .Where(b => !seenInCurrent.Contains(b))
            .ToList();

        return new AuditBaselineDiff(newFindings, resolved, persistent);
    }

    private sealed class FindingKeyComparer : IEqualityComparer<AuditFinding>
    {
        public static readonly FindingKeyComparer Instance = new();

        public bool Equals(AuditFinding? x, AuditFinding? y)
        {
            if (ReferenceEquals(x, y)) return true;
            if (x is null || y is null) return false;
            return string.Equals(x.Category, y.Category, StringComparison.Ordinal)
                && string.Equals(x.Identity, y.Identity, StringComparison.Ordinal)
                && string.Equals(x.Detail, y.Detail, StringComparison.Ordinal)
                && string.Equals(x.Severity ?? string.Empty, y.Severity ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        }

        public int GetHashCode(AuditFinding obj)
        {
            return HashCode.Combine(
                obj.Category,
                obj.Identity,
                obj.Detail,
                (obj.Severity ?? string.Empty).ToUpperInvariant());
        }
    }
}

public sealed record AuditBaselineDiff(
    IReadOnlyList<AuditFinding> New,
    IReadOnlyList<AuditFinding> Resolved,
    IReadOnlyList<AuditFinding> Persistent)
{
    public int NewCount => New.Count;
    public int ResolvedCount => Resolved.Count;
    public int PersistentCount => Persistent.Count;
    public bool HasChanges => NewCount > 0 || ResolvedCount > 0;
}
