using System.Text;
using FluentAssertions;
using Grex365.Core.Csv;

namespace Grex365.Core.Tests;

public class FlexibleCsvReaderTests
{
    private static Stream MakeStream(string content) =>
        new MemoryStream(Encoding.UTF8.GetBytes(content));

    [Fact]
    public void ReadsCommaDelimited()
    {
        var rows = FlexibleCsvReader.Read(MakeStream("Email,Id\nuser@a.com,abc\nuser2@a.com,def"));
        rows.Should().HaveCount(2);
        rows[0]["Email"].Should().Be("user@a.com");
        rows[0]["Id"].Should().Be("abc");
        rows[1]["Email"].Should().Be("user2@a.com");
    }

    [Fact]
    public void ReadsSemicolonDelimited()
    {
        var rows = FlexibleCsvReader.Read(MakeStream("Action;Permission;Mailbox;Principal\nadd;FullAccess;m@a;p@a"));
        rows.Should().HaveCount(1);
        rows[0]["Action"].Should().Be("add");
        rows[0]["Permission"].Should().Be("FullAccess");
        rows[0]["Mailbox"].Should().Be("m@a");
        rows[0]["Principal"].Should().Be("p@a");
    }

    [Fact]
    public void HandlesQuotedFieldsWithDelimiterInside()
    {
        var rows = FlexibleCsvReader.Read(MakeStream("Name,Email\n\"Doe, Jane\",jane@a.com"));
        rows.Should().HaveCount(1);
        rows[0]["Name"].Should().Be("Doe, Jane");
        rows[0]["Email"].Should().Be("jane@a.com");
    }

    [Fact]
    public void HandlesEscapedQuotes()
    {
        var rows = FlexibleCsvReader.Read(MakeStream("Note\n\"She said \"\"hi\"\"\""));
        rows[0]["Note"].Should().Be("She said \"hi\"");
    }

    [Fact]
    public void SkipsBlankLines()
    {
        var rows = FlexibleCsvReader.Read(MakeStream("Email\nuser@a\n\n\nuser2@a"));
        rows.Should().HaveCount(2);
    }

    [Fact]
    public void EmptyStreamReturnsEmpty()
    {
        var rows = FlexibleCsvReader.Read(MakeStream(string.Empty));
        rows.Should().BeEmpty();
    }

    [Fact]
    public void CaseInsensitiveColumnAccess()
    {
        var rows = FlexibleCsvReader.Read(MakeStream("Email\nuser@a"));
        rows[0]["email"].Should().Be("user@a");
        rows[0]["EMAIL"].Should().Be("user@a");
    }

    [Fact]
    public void MissingFieldsBecomeEmpty()
    {
        var rows = FlexibleCsvReader.Read(MakeStream("A,B,C\n1,2"));
        rows[0]["C"].Should().Be(string.Empty);
    }
}
