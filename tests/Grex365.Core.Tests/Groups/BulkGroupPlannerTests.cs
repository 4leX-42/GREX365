using FluentAssertions;
using Grex365.Core.Groups;
using Grex365.Core.Models;

namespace Grex365.Core.Tests.Groups;

public class BulkGroupPlannerTests
{
    private static BulkGroupRow Row(string name, string email, string type = "M365")
        => new(name, email, type);

    [Fact]
    public void Plan_NullRows_Throws()
    {
        Action act = () => BulkGroupPlanner.Plan(null!, "Auto");
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void Plan_EmptyRows_ReturnsZeroCounts()
    {
        var plan = BulkGroupPlanner.Plan(Array.Empty<BulkGroupRow>(), "Auto");

        plan.M365Rows.Should().BeEmpty();
        plan.DlRows.Should().BeEmpty();
        plan.DistinctM365GroupCount.Should().Be(0);
        plan.DistinctDlGroupCount.Should().Be(0);
        plan.Breakdown.Should().BeEmpty();
    }

    [Fact]
    public void Plan_AutoChoice_SplitsByRowType()
    {
        var rows = new[]
        {
            Row("G1", "a@x.com", "M365"),
            Row("G1", "b@x.com", "M365"),
            Row("G2", "c@x.com", "DL"),
        };

        var plan = BulkGroupPlanner.Plan(rows, "Auto");

        plan.M365Rows.Should().HaveCount(2);
        plan.DlRows.Should().HaveCount(1);
        plan.DistinctM365GroupCount.Should().Be(1);
        plan.DistinctDlGroupCount.Should().Be(1);
        plan.Breakdown.Should().Be("1 M365 + 1 DL");
        plan.TypeHint.Should().Be(BulkGroupPlanner.TypeHintAuto);
    }

    [Fact]
    public void Plan_ForcedM365_OverridesAllRows()
    {
        var rows = new[]
        {
            Row("G1", "a@x.com", "DL"),
            Row("G2", "b@x.com", "M365"),
        };

        var plan = BulkGroupPlanner.Plan(rows, "M365");

        plan.M365Rows.Should().HaveCount(2);
        plan.DlRows.Should().BeEmpty();
        plan.DistinctM365GroupCount.Should().Be(2);
        plan.TypeHint.Should().Contain("M365").And.Contain("Forzado");
    }

    [Fact]
    public void Plan_ForcedDl_OverridesAllRows()
    {
        var rows = new[]
        {
            Row("G1", "a@x.com", "M365"),
            Row("G2", "b@x.com", "M365"),
        };

        var plan = BulkGroupPlanner.Plan(rows, "DL");

        plan.M365Rows.Should().BeEmpty();
        plan.DlRows.Should().HaveCount(2);
        plan.DistinctDlGroupCount.Should().Be(2);
        plan.TypeHint.Should().Contain("DL").And.Contain("Forzado");
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("  ")]
    [InlineData("auto")]
    [InlineData("AUTO")]
    [InlineData("unknown")]
    public void Plan_NullOrEmptyOrUnknownChoice_TreatedAsAuto(string? choice)
    {
        var rows = new[] { Row("G1", "a@x.com", "M365") };

        var plan = BulkGroupPlanner.Plan(rows, choice!);

        plan.M365Rows.Should().HaveCount(1);
        plan.TypeHint.Should().StartWith(BulkGroupPlanner.TypeHintAuto[..15], because: "auto/unknown choice keeps row-level GroupType");
    }

    [Fact]
    public void Plan_ChoiceCaseInsensitive()
    {
        var rows = new[] { Row("G1", "a@x.com", "DL") };

        var planLower = BulkGroupPlanner.Plan(rows, "m365");
        var planUpper = BulkGroupPlanner.Plan(rows, "M365");

        planLower.M365Rows.Should().HaveCount(1);
        planUpper.M365Rows.Should().HaveCount(1);
    }

    [Fact]
    public void Plan_DistinctGroupCount_IsCaseInsensitive()
    {
        var rows = new[]
        {
            Row("Sales", "a@x.com"),
            Row("SALES", "b@x.com"),
            Row("sales", "c@x.com"),
        };

        var plan = BulkGroupPlanner.Plan(rows, "Auto");

        plan.DistinctM365GroupCount.Should().Be(1);
        plan.M365Rows.Should().HaveCount(3);
    }

    [Fact]
    public void Plan_BreakdownOmitsZeroSides()
    {
        var onlyM365 = new[] { Row("G1", "a@x.com", "M365") };
        var onlyDl = new[] { Row("G1", "a@x.com", "DL") };

        BulkGroupPlanner.Plan(onlyM365, "Auto").Breakdown.Should().Be("1 M365");
        BulkGroupPlanner.Plan(onlyDl, "Auto").Breakdown.Should().Be("1 DL");
    }

    [Fact]
    public void BuildConfirmMessage_IncludesBreakdownRowCountDomainHint()
    {
        var rows = new[]
        {
            Row("G1", "a@x.com", "M365"),
            Row("G2", "b@x.com", "DL"),
        };
        var plan = BulkGroupPlanner.Plan(rows, "Auto");

        var msg = BulkGroupPlanner.BuildConfirmMessage(plan, rows.Length, "contoso.com");

        msg.Should().Contain("1 M365 + 1 DL");
        msg.Should().Contain("(2 miembros)");
        msg.Should().Contain("@contoso.com");
        msg.Should().Contain("¿Continuar?");
    }

    [Fact]
    public void BuildConfirmMessage_TrimsDomainAndStripsLeadingAt()
    {
        var plan = BulkGroupPlanner.Plan(new[] { Row("G1", "a@x.com") }, "Auto");

        var msg = BulkGroupPlanner.BuildConfirmMessage(plan, 1, "  @contoso.com  ");

        msg.Should().Contain("@contoso.com");
        msg.Should().NotContain("@@");
    }

    [Fact]
    public void BuildConfirmMessage_NullPlan_Throws()
    {
        Action act = () => BulkGroupPlanner.BuildConfirmMessage(null!, 0, "x.com");
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void Summarize_NullResults_Throws()
    {
        Action act = () => BulkGroupPlanner.Summarize(null!);
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void Summarize_EmptyResults_AllZero()
    {
        var s = BulkGroupPlanner.Summarize(Array.Empty<BulkGroupResult>());

        s.Should().Be("Grupos: Nuevos=0  YaEstaban=0  Miembros=0  Err=0");
    }

    [Fact]
    public void Summarize_CountsByAction()
    {
        var results = new[]
        {
            new BulkGroupResult("G1", "g1@x.com", "Created", null, string.Empty),
            new BulkGroupResult("G1", "g1@x.com", "MemberAdded", "a@x.com", string.Empty),
            new BulkGroupResult("G1", "g1@x.com", "MemberAdded", "b@x.com", string.Empty),
            new BulkGroupResult("G2", "g2@x.com", "Skipped", null, string.Empty),
            new BulkGroupResult("G3", "g3@x.com", "Error", null, "boom"),
        };

        var s = BulkGroupPlanner.Summarize(results);

        s.Should().Be("Grupos: Nuevos=1  YaEstaban=1  Miembros=2  Err=1");
    }

    [Fact]
    public void Summarize_UnknownActions_Ignored()
    {
        var results = new[]
        {
            new BulkGroupResult("G1", "g1@x.com", "Created", null, string.Empty),
            new BulkGroupResult("G1", "g1@x.com", "WhatIsThis", null, string.Empty),
        };

        var s = BulkGroupPlanner.Summarize(results);

        s.Should().Be("Grupos: Nuevos=1  YaEstaban=0  Miembros=0  Err=0");
    }
}
