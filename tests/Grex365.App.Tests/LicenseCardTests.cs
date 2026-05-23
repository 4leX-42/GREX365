using FluentAssertions;
using Grex365.App.ViewModels;
using Grex365.Core.Models;

namespace Grex365.App.Tests;

public class LicenseCardTests
{
    private static LicenseSummary Summary(string sku = "SPE_E5", int consumed = 50, int enabled = 100)
        => new(SkuPartNumber: sku, SkuId: "sku-id-" + sku, Consumed: consumed, Enabled: enabled, Warning: 0, Suspended: 0);

    [Fact]
    public void From_KnownSku_ResolvesFriendlyNameAndCategory()
    {
        var card = LicenseCard.From(Summary("SPE_E5", 30, 50));

        card.FriendlyName.Should().Be("Microsoft 365 E5");
        card.Category.Should().Be(LicenseCategory.Enterprise);
        card.CategoryLabel.Should().NotBeNullOrWhiteSpace();
    }

    [Fact]
    public void From_UnknownSku_HumanizesAsFallback()
    {
        var card = LicenseCard.From(Summary("CUSTOM_FUTURE_PLAN", 1, 10));

        card.SkuPartNumber.Should().Be("CUSTOM_FUTURE_PLAN");
        card.FriendlyName.Should().NotBeNullOrWhiteSpace();
        card.Category.Should().Be(LicenseCategory.Other);
    }

    [Fact]
    public void From_ZeroEnabled_PercentZero_LevelLow_AvailableZero()
    {
        var card = LicenseCard.From(Summary(consumed: 0, enabled: 0));

        card.PercentUsed.Should().Be(0);
        card.UtilizationLevel.Should().Be("Low");
        card.Available.Should().Be(0);
    }

    [Fact]
    public void From_HalfUsed_LevelLow()
    {
        var card = LicenseCard.From(Summary(consumed: 50, enabled: 100));

        card.PercentUsed.Should().Be(50);
        card.UtilizationLevel.Should().Be("Low");
        card.Available.Should().Be(50);
    }

    [Theory]
    [InlineData(60, 100, "Medium")]
    [InlineData(79, 100, "Medium")]
    [InlineData(80, 100, "High")]
    [InlineData(94, 100, "High")]
    [InlineData(95, 100, "Critical")]
    [InlineData(100, 100, "Critical")]
    public void From_UtilizationLevels_Boundaries(int consumed, int enabled, string expected)
    {
        var card = LicenseCard.From(Summary(consumed: consumed, enabled: enabled));
        card.UtilizationLevel.Should().Be(expected);
    }

    [Fact]
    public void From_OverConsumed_ClampsAvailableToZero()
    {
        var card = LicenseCard.From(Summary(consumed: 120, enabled: 100));

        card.Available.Should().Be(0, because: "Math.Max clamps negative to 0");
        card.PercentUsed.Should().Be(120);
        card.UtilizationLevel.Should().Be("Critical");
    }

    [Fact]
    public void From_PreservesSeatCounts()
    {
        var card = LicenseCard.From(Summary(consumed: 42, enabled: 100));

        card.Consumed.Should().Be(42);
        card.Enabled.Should().Be(100);
        card.Available.Should().Be(58);
    }

    [Fact]
    public void From_PriorityFromSkuInfo_NotZero_ForKnownSku()
    {
        var card = LicenseCard.From(Summary("SPE_E5"));
        card.Priority.Should().BeGreaterThan(0);
    }
}
