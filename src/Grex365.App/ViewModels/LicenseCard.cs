using Grex365.Core.Models;

namespace Grex365.App.ViewModels;

public sealed record LicenseCard(
    string SkuPartNumber,
    string FriendlyName,
    string CategoryLabel,
    LicenseCategory Category,
    int Priority,
    int Consumed,
    int Enabled,
    int Available,
    double PercentUsed,
    string UtilizationLevel)    // "Low" | "Medium" | "High" | "Critical"
{
    public static LicenseCard From(LicenseSummary summary)
    {
        var info = SkuCatalog.Resolve(summary.SkuPartNumber);
        var pct = summary.Enabled > 0 ? summary.Consumed * 100.0 / summary.Enabled : 0;
        var level = pct switch
        {
            >= 95 => "Critical",
            >= 80 => "High",
            >= 60 => "Medium",
            _ => "Low",
        };
        return new LicenseCard(
            SkuPartNumber: summary.SkuPartNumber,
            FriendlyName: info.FriendlyName,
            CategoryLabel: SkuCatalog.CategoryLabel(info.Category),
            Category: info.Category,
            Priority: info.Priority,
            Consumed: summary.Consumed,
            Enabled: summary.Enabled,
            Available: Math.Max(summary.Enabled - summary.Consumed, 0),
            PercentUsed: pct,
            UtilizationLevel: level);
    }
}
