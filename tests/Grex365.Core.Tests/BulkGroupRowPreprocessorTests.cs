using FluentAssertions;
using Grex365.Core.Groups;

namespace Grex365.Core.Tests;

public class BulkGroupRowPreprocessorTests
{
    private static Dictionary<string, string> Row(string? group, string? email, string? type = null)
    {
        var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (group is not null) d["GroupName"] = group;
        if (email is not null) d["Email"] = email;
        if (type is not null) d["GroupType"] = type;
        return d;
    }

    [Theory]
    [InlineData("M365", "M365")]
    [InlineData("Microsoft 365", "M365")]
    [InlineData("unified", "M365")]
    [InlineData("DL", "DL")]
    [InlineData("Distribution", "DL")]
    [InlineData("distributionlist", "DL")]
    [InlineData("Distribution List", "DL")]
    [InlineData("Exchange", "DL")]
    [InlineData("nonsense", "M365")]
    [InlineData("", "M365")]
    public void NormalizeType_MapsAliases(string input, string expected) =>
        BulkGroupRowPreprocessor.NormalizeType(input).Should().Be(expected);

    [Fact]
    public void Normalize_DetectsGroupType_FromColumn()
    {
        var raw = new[]
        {
            Row("Ventas",  "a@x.com", "M365"),
            Row("Soporte", "b@x.com", "DL"),
            Row("Otros",   "c@x.com"),
        };
        var rows = BulkGroupRowPreprocessor.Normalize(raw);
        rows.Should().HaveCount(3);
        rows[0].GroupType.Should().Be("M365");
        rows[1].GroupType.Should().Be("DL");
        rows[2].GroupType.Should().Be("M365");
    }

    [Fact]
    public void Normalize_ForwardFillsGroupType_AlongWithName()
    {
        var raw = new[]
        {
            Row("Soporte", "a@x.com", "DL"),
            Row("",        "b@x.com"),
            Row("",        "c@x.com"),
        };
        var rows = BulkGroupRowPreprocessor.Normalize(raw);
        rows.Should().HaveCount(3);
        rows.Select(r => r.GroupType).Should().AllBe("DL");
    }

    [Fact]
    public void ForwardFills_GroupName_FromPreviousNonEmpty()
    {
        var raw = new[]
        {
            Row("Ventas", "a@x.com"),
            Row("",       "b@x.com"),
            Row("  ",     "c@x.com"),
        };
        var rows = BulkGroupRowPreprocessor.Normalize(raw);
        rows.Should().HaveCount(3);
        rows.Select(r => r.GroupName).Should().AllBe("Ventas");
        rows.Select(r => r.Email).Should().Equal("a@x.com", "b@x.com", "c@x.com");
    }

    [Fact]
    public void TransitionsToNewGroupName_WhenNonEmpty()
    {
        var raw = new[]
        {
            Row("A", "1@x.com"),
            Row("",  "2@x.com"),
            Row("B", "3@x.com"),
            Row("",  "4@x.com"),
        };
        var rows = BulkGroupRowPreprocessor.Normalize(raw);
        rows.Select(r => r.GroupName).Should().Equal("A", "A", "B", "B");
    }

    [Fact]
    public void Skips_Rows_WithNoGroupNameYet()
    {
        var raw = new[]
        {
            Row("",  "orphan@x.com"),
            Row("A", "1@x.com"),
        };
        var rows = BulkGroupRowPreprocessor.Normalize(raw);
        rows.Should().HaveCount(1);
        rows[0].GroupName.Should().Be("A");
    }

    [Fact]
    public void Skips_Rows_WithEmptyEmail()
    {
        var raw = new[]
        {
            Row("A", ""),
            Row("A", "  "),
            Row("A", "ok@x.com"),
        };
        var rows = BulkGroupRowPreprocessor.Normalize(raw);
        rows.Should().HaveCount(1);
        rows[0].Email.Should().Be("ok@x.com");
    }

    [Fact]
    public void Trims_Whitespace()
    {
        var raw = new[] { Row("  Ventas  ", "  a@x.com  ") };
        var rows = BulkGroupRowPreprocessor.Normalize(raw);
        rows[0].GroupName.Should().Be("Ventas");
        rows[0].Email.Should().Be("a@x.com");
    }

    [Theory]
    [InlineData("user@example.com", true)]
    [InlineData("a@b.co", true)]
    [InlineData("invalid", false)]
    [InlineData("@example.com", false)]
    [InlineData("user@", false)]
    [InlineData("user@example", false)]
    [InlineData("user@.com", false)]
    [InlineData("", false)]
    public void IsEmail_BasicShape(string input, bool expected)
    {
        BulkGroupRowPreprocessor.IsEmail(input).Should().Be(expected);
    }
}
