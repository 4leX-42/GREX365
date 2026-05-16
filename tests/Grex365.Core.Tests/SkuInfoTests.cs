using FluentAssertions;
using Grex365.Core.Models;

namespace Grex365.Core.Tests;

public class SkuInfoTests
{
    [Fact]
    public void Available_ComputesFreeSeats()
    {
        var s = new SkuInfo(Guid.NewGuid(), "E3", Enabled: 10, Consumed: 7);
        s.Available.Should().Be(3);
    }

    [Fact]
    public void Available_NeverNegative_WhenOverConsumed()
    {
        var s = new SkuInfo(Guid.NewGuid(), "E3", Enabled: 5, Consumed: 8);
        s.Available.Should().Be(0);
    }

    [Fact]
    public void Available_AllConsumed_IsZero()
    {
        var s = new SkuInfo(Guid.NewGuid(), "E3", Enabled: 5, Consumed: 5);
        s.Available.Should().Be(0);
    }

    [Fact]
    public void Display_IncludesPartNumberAndCounts()
    {
        var s = new SkuInfo(Guid.NewGuid(), "ENTERPRISEPACK", Enabled: 100, Consumed: 25);
        s.Display.Should().Contain("ENTERPRISEPACK").And.Contain("75").And.Contain("100");
    }

    [Fact]
    public void Display_FallsBackToGuid_WhenPartNumberEmpty()
    {
        var id = Guid.NewGuid();
        var s = new SkuInfo(id, string.Empty, 0, 0);
        s.Display.Should().Be(id.ToString());
    }

    [Fact]
    public void OrderingByAvailableThenName_Works()
    {
        var skus = new[]
        {
            new SkuInfo(Guid.NewGuid(), "BB", 10, 10),
            new SkuInfo(Guid.NewGuid(), "AA", 10, 3),
            new SkuInfo(Guid.NewGuid(), "CC", 10, 3),
        };
        var ordered = skus
            .OrderByDescending(s => s.Available)
            .ThenBy(s => s.SkuPartNumber)
            .ToList();
        ordered[0].SkuPartNumber.Should().Be("AA");
        ordered[1].SkuPartNumber.Should().Be("CC");
        ordered[2].SkuPartNumber.Should().Be("BB");
    }
}
