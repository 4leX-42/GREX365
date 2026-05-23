namespace Grex365.App.ViewModels;

public static class LicenseFilterMatcher
{
    // True iff the card matches the free-text filter (case-insensitive substring on
    // FriendlyName / SkuPartNumber / CategoryLabel). Whitespace-only filter = no filter.
    public static bool Matches(LicenseCard card, string? filter)
    {
        ArgumentNullException.ThrowIfNull(card);
        var q = (filter ?? string.Empty).Trim();
        if (q.Length == 0) return true;
        return (card.FriendlyName?.Contains(q, StringComparison.OrdinalIgnoreCase) ?? false)
            || (card.SkuPartNumber?.Contains(q, StringComparison.OrdinalIgnoreCase) ?? false)
            || (card.CategoryLabel?.Contains(q, StringComparison.OrdinalIgnoreCase) ?? false);
    }
}
