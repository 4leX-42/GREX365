using FluentAssertions;
using Grex365.Core.Abstractions;
using Grex365.Core.Audit;

namespace Grex365.Core.Tests;

public class RecentActivityAggregatorTests
{
    private static readonly DateTimeOffset Now = new(2026, 6, 5, 14, 0, 0, TimeSpan.FromHours(2));

    private static AuditRecord R(DateTimeOffset at, string outcome = "OK", string source = "Users") =>
        new(at, "admin", source, outcome, $"op-{at:HHmmss}");

    [Fact]
    public void Empty_ReturnsZeroes()
    {
        var a = RecentActivityAggregator.Compute(Array.Empty<AuditRecord>(), Now);
        a.TodayCount.Should().Be(0);
        a.TodayErrors.Should().Be(0);
        a.Recent.Should().BeEmpty();
    }

    [Fact]
    public void CountsOnlyToday_LocalDate()
    {
        var records = new[]
        {
            R(Now.AddHours(-1)),                    // hoy
            R(Now.AddHours(-2), "ERROR"),           // hoy, error
            R(Now.AddDays(-1)),                     // ayer
            R(Now.AddDays(-10), "ERROR"),           // viejo — no cuenta para hoy
        };
        var a = RecentActivityAggregator.Compute(records, Now);
        a.TodayCount.Should().Be(2);
        a.TodayErrors.Should().Be(1);
    }

    [Theory]
    [InlineData("ERROR")]
    [InlineData("err")]
    [InlineData(" FATAL ")]
    public void TodayErrors_OutcomeVariants(string outcome)
    {
        var a = RecentActivityAggregator.Compute(new[] { R(Now, outcome) }, Now);
        a.TodayErrors.Should().Be(1);
    }

    [Fact]
    public void Recent_NewestFirst_CappedAtTake()
    {
        var records = Enumerable.Range(0, 10).Select(i => R(Now.AddMinutes(-i))).ToList();
        var a = RecentActivityAggregator.Compute(records, Now, take: 5);
        a.Recent.Should().HaveCount(5);
        a.Recent[0].Timestamp.Should().Be(Now);                    // el más nuevo primero
        a.Recent.Should().BeInDescendingOrder(r => r.Timestamp);
    }

    [Fact]
    public void Recent_IncludesAllOutcomes()
    {
        var records = new[] { R(Now, "OK"), R(Now.AddMinutes(-1), "ERROR"), R(Now.AddMinutes(-2), "WARN") };
        var a = RecentActivityAggregator.Compute(records, Now);
        a.Recent.Should().HaveCount(3); // no filtra por outcome (a diferencia de MetricsAggregator.RecentErrors)
    }

    [Fact]
    public void NullRecords_Throws()
    {
        var act = () => RecentActivityAggregator.Compute(null!, Now);
        act.Should().Throw<ArgumentNullException>();
    }
}
