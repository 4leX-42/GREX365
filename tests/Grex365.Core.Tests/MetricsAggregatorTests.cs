using FluentAssertions;
using Grex365.Core.Abstractions;
using Grex365.Core.Audit;

namespace Grex365.Core.Tests;

public class MetricsAggregatorTests
{
    private static AuditRecord Rec(string outcome, string source, DateTimeOffset ts, string message = "x")
        => new(ts, "actor", source, outcome, message);

    [Fact]
    public void Empty_Returns_Zero_Metrics()
    {
        var m = MetricsAggregator.Compute(Array.Empty<AuditRecord>(), DateTimeOffset.Parse("2026-05-20T10:00:00Z"));
        m.TotalCount.Should().Be(0);
        m.OkCount.Should().Be(0);
        m.WarnCount.Should().Be(0);
        m.ErrorCount.Should().Be(0);
        m.ErrorRate.Should().Be(0.0);
        m.Last24hCount.Should().Be(0);
        m.TopSources.Should().BeEmpty();
        m.RecentErrors.Should().BeEmpty();
    }

    [Fact]
    public void Counts_By_Outcome_Are_Correct_And_Case_Insensitive()
    {
        var now = DateTimeOffset.Parse("2026-05-20T10:00:00Z");
        var records = new[]
        {
            Rec("OK", "Users", now),
            Rec("ok", "Users", now),
            Rec("WARN", "Groups", now),
            Rec("Warning", "Groups", now),
            Rec("ERROR", "Audit", now),
            Rec("err", "Audit", now),
            Rec("Fatal", "Audit", now),
        };
        var m = MetricsAggregator.Compute(records, now);
        m.OkCount.Should().Be(2);
        m.WarnCount.Should().Be(2);
        m.ErrorCount.Should().Be(3);
        m.TotalCount.Should().Be(7);
        m.ErrorRate.Should().BeApproximately(3d / 7d, 0.001);
    }

    [Fact]
    public void Last24h_Counts_Recent_Entries_Only()
    {
        var now = DateTimeOffset.Parse("2026-05-20T10:00:00Z");
        var records = new[]
        {
            Rec("OK", "x", now.AddHours(-1)),     // in window
            Rec("OK", "x", now.AddHours(-23)),    // in window
            Rec("OK", "x", now.AddHours(-25)),    // outside
            Rec("ERROR", "x", now.AddDays(-7)),   // outside
        };
        var m = MetricsAggregator.Compute(records, now);
        m.Last24hCount.Should().Be(2);
    }

    [Fact]
    public void Top_Sources_Are_Ordered_By_Count_Desc_Then_Alpha()
    {
        var now = DateTimeOffset.Parse("2026-05-20T10:00:00Z");
        var records = new[]
        {
            Rec("OK", "Users", now),
            Rec("OK", "Users", now),
            Rec("OK", "Users", now),
            Rec("OK", "Groups", now),
            Rec("OK", "Groups", now),
            Rec("OK", "Audit", now),
            Rec("OK", "Connect", now),
        };
        var m = MetricsAggregator.Compute(records, now, topSources: 3);
        m.TopSources.Should().HaveCount(3);
        m.TopSources[0].Source.Should().Be("Users");
        m.TopSources[0].Count.Should().Be(3);
        m.TopSources[1].Source.Should().Be("Groups");
        m.TopSources[1].Count.Should().Be(2);
        m.TopSources[2].Source.Should().Be("Audit"); // alpha vs Connect (same count 1)
    }

    [Fact]
    public void Recent_Errors_Are_Latest_N()
    {
        var now = DateTimeOffset.Parse("2026-05-20T10:00:00Z");
        var records = new[]
        {
            Rec("ERROR", "A", now.AddMinutes(-5), "newest"),
            Rec("ERROR", "B", now.AddMinutes(-15), "mid"),
            Rec("OK", "C", now.AddMinutes(-1)),
            Rec("ERROR", "D", now.AddMinutes(-30), "oldest"),
        };
        var m = MetricsAggregator.Compute(records, now, recentErrors: 2);
        m.RecentErrors.Should().HaveCount(2);
        m.RecentErrors[0].Message.Should().Be("newest");
        m.RecentErrors[1].Message.Should().Be("mid");
    }

    [Fact]
    public void Blank_Source_Bucketed_As_Sin_Source()
    {
        var now = DateTimeOffset.Parse("2026-05-20T10:00:00Z");
        var records = new[]
        {
            Rec("OK", "", now),
            Rec("OK", "  ", now),
            Rec("OK", "Users", now),
        };
        var m = MetricsAggregator.Compute(records, now);
        m.TopSources.Should().Contain(s => s.Source == "(sin source)" && s.Count == 2);
        m.TopSources.Should().Contain(s => s.Source == "Users" && s.Count == 1);
    }
}
