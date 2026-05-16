using FluentAssertions;
using Grex365.Core.Groups;

namespace Grex365.Core.Tests;

public class BulkGroupRowPreprocessorTests
{
    private static Dictionary<string, string> Row(string? group, string? email)
    {
        var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (group is not null) d["GroupName"] = group;
        if (email is not null) d["Email"] = email;
        return d;
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
