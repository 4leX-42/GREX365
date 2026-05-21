using System.Text;
using FluentAssertions;
using Grex365.Core.Audit;

namespace Grex365.Core.Tests;

public class GroupActivityAnalyzerTests
{
    private const string Header =
        "Report Refresh Date,Group Display Name,Is Deleted,Owner Principal Name,Last Activity Date,Group Type," +
        "Member Count,External Member Count,Exchange Received Email Count,SharePoint Active File Count," +
        "Yammer Posted Message Count,Yammer Read Message Count,Yammer Liked Message Count,Exchange Mailbox Total Item Count," +
        "Exchange Mailbox Storage Used (Byte),SharePoint Total File Count,SharePoint Site Storage Used (Byte),Report Period";

    private static readonly DateOnly Today = new(2026, 5, 21);

    private static Stream Csv(params string[] dataRows)
    {
        var sb = new StringBuilder();
        sb.AppendLine(Header);
        foreach (var r in dataRows)
        {
            sb.AppendLine(r);
        }
        return new MemoryStream(Encoding.UTF8.GetBytes(sb.ToString()));
    }

    [Fact]
    public void ParseCsv_EmptyStream_ReturnsEmpty()
    {
        var rows = GroupActivityAnalyzer.ParseCsv(new MemoryStream(Array.Empty<byte>()));
        rows.Should().BeEmpty();
    }

    [Fact]
    public void ParseCsv_ReadsCoreColumns()
    {
        using var s = Csv(
            "2026-05-19,Sales Team,False,boss@a.com,2026-04-12,Microsoft 365,42,2,100,50,0,0,0,9000,50000,500,1024000,180");
        var rows = GroupActivityAnalyzer.ParseCsv(s);

        rows.Should().ContainSingle();
        var r = rows[0];
        r.DisplayName.Should().Be("Sales Team");
        r.IsDeleted.Should().BeFalse();
        r.OwnerPrincipalName.Should().Be("boss@a.com");
        r.LastActivityDate.Should().Be(new DateOnly(2026, 4, 12));
        r.GroupType.Should().Be("Microsoft 365");
        r.MemberCount.Should().Be(42);
        r.ExternalMemberCount.Should().Be(2);
    }

    [Fact]
    public void ParseCsv_HandlesQuotedFieldsWithCommas()
    {
        using var s = Csv("2026-05-19,\"Sales, EMEA\",False,boss@a.com,2026-04-12,Microsoft 365,1,0,0,0,0,0,0,0,0,0,0,180");
        var rows = GroupActivityAnalyzer.ParseCsv(s);
        rows.Single().DisplayName.Should().Be("Sales, EMEA");
    }

    [Fact]
    public void ParseCsv_HandlesMissingLastActivityDate()
    {
        using var s = Csv("2026-05-19,Sales,False,boss@a.com,,Microsoft 365,1,0,0,0,0,0,0,0,0,0,0,180");
        var rows = GroupActivityAnalyzer.ParseCsv(s);
        rows.Single().LastActivityDate.Should().BeNull();
    }

    [Fact]
    public void Analyze_FlagsGroupOlderThanThreshold()
    {
        var rows = new[]
        {
            new GroupActivityRow("Stale", IsDeleted: false, "boss@a.com", new DateOnly(2025, 1, 1), "Microsoft 365", 5, 0),
            new GroupActivityRow("Fresh", IsDeleted: false, "boss@a.com", new DateOnly(2026, 5, 10), "Microsoft 365", 5, 0),
        };
        var findings = GroupActivityAnalyzer.Analyze(rows, Today, 90);
        findings.Should().ContainSingle(f => f.Identity == "Stale");
        findings[0].Severity.Should().Be("WARN");
        findings[0].Category.Should().Be("Inactive M365 group");
    }

    [Fact]
    public void Analyze_FlagsGroupWithNoLastActivity()
    {
        var rows = new[]
        {
            new GroupActivityRow("Never", IsDeleted: false, "boss@a.com", LastActivityDate: null, "Microsoft 365", 3, 0),
        };
        var findings = GroupActivityAnalyzer.Analyze(rows, Today, 30);
        findings.Should().ContainSingle();
        findings[0].Detail.Should().Contain("sin actividad");
    }

    [Fact]
    public void Analyze_SkipsDeletedGroups()
    {
        var rows = new[]
        {
            new GroupActivityRow("Gone", IsDeleted: true, "boss@a.com", new DateOnly(2024, 1, 1), "Microsoft 365", 0, 0),
        };
        var findings = GroupActivityAnalyzer.Analyze(rows, Today, 90);
        findings.Should().BeEmpty();
    }

    [Fact]
    public void Analyze_SkipsRowsWithBlankDisplayName()
    {
        var rows = new[]
        {
            new GroupActivityRow(string.Empty, IsDeleted: false, null, LastActivityDate: null, null, 0, 0),
        };
        GroupActivityAnalyzer.Analyze(rows, Today, 30).Should().BeEmpty();
    }

    [Fact]
    public void Analyze_ThresholdLessThanOne_Throws()
    {
        var act = () => GroupActivityAnalyzer.Analyze(Array.Empty<GroupActivityRow>(), Today, 0);
        act.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void Analyze_BoundaryDayEqualsCutoff_NotFlagged()
    {
        // cutoff = today - 90; group with LastActivity == cutoff should NOT be flagged (strict <)
        var rows = new[]
        {
            new GroupActivityRow("Edge", IsDeleted: false, "boss@a.com", Today.AddDays(-90), "Microsoft 365", 1, 0),
        };
        GroupActivityAnalyzer.Analyze(rows, Today, 90).Should().BeEmpty();
    }
}
