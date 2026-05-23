using FluentAssertions;
using Grex365.App.ViewModels;
using Grex365.Core.Models;

namespace Grex365.App.Tests;

public class LicenseFilterMatcherTests
{
    private static LicenseCard Card(
        string sku = "SPE_E5",
        string friendly = "Microsoft 365 E5",
        string label = "Enterprise")
        => new(
            SkuPartNumber: sku,
            FriendlyName: friendly,
            CategoryLabel: label,
            Category: LicenseCategory.Enterprise,
            Priority: 10,
            Consumed: 0,
            Enabled: 0,
            Available: 0,
            PercentUsed: 0,
            UtilizationLevel: "Low");

    [Fact]
    public void Matches_NullFilter_ReturnsTrue()
    {
        LicenseFilterMatcher.Matches(Card(), null).Should().BeTrue();
    }

    [Fact]
    public void Matches_EmptyFilter_ReturnsTrue()
    {
        LicenseFilterMatcher.Matches(Card(), string.Empty).Should().BeTrue();
    }

    [Fact]
    public void Matches_WhitespaceOnly_ReturnsTrue()
    {
        LicenseFilterMatcher.Matches(Card(), "   ").Should().BeTrue();
    }

    [Fact]
    public void Matches_OnFriendlyName()
    {
        LicenseFilterMatcher.Matches(Card(friendly: "Microsoft 365 E5"), "365").Should().BeTrue();
    }

    [Fact]
    public void Matches_OnSkuPartNumber()
    {
        LicenseFilterMatcher.Matches(Card(sku: "SPE_E5"), "spe").Should().BeTrue();
    }

    [Fact]
    public void Matches_OnCategoryLabel()
    {
        LicenseFilterMatcher.Matches(Card(label: "Enterprise"), "enter").Should().BeTrue();
    }

    [Fact]
    public void Matches_CaseInsensitive()
    {
        LicenseFilterMatcher.Matches(Card(friendly: "Microsoft 365 E5"), "MICROSOFT").Should().BeTrue();
        LicenseFilterMatcher.Matches(Card(sku: "SPE_E5"), "spe_E5").Should().BeTrue();
    }

    [Fact]
    public void Matches_TrimsFilter()
    {
        LicenseFilterMatcher.Matches(Card(friendly: "Visio Plan 2"), "  Visio  ").Should().BeTrue();
    }

    [Fact]
    public void Matches_NoMatch_ReturnsFalse()
    {
        LicenseFilterMatcher.Matches(Card(friendly: "Microsoft 365 E5", sku: "SPE_E5", label: "Enterprise"),
            "Frontline").Should().BeFalse();
    }

    [Fact]
    public void Matches_AnyFieldMatches_ShortCircuits()
    {
        // Match comes from CategoryLabel even when FriendlyName/Sku miss.
        LicenseFilterMatcher.Matches(Card(friendly: "X", sku: "Y", label: "Business"), "Business")
            .Should().BeTrue();
    }

    [Fact]
    public void Matches_NullCard_Throws()
    {
        Action act = () => LicenseFilterMatcher.Matches(null!, "anything");
        act.Should().Throw<ArgumentNullException>();
    }
}
