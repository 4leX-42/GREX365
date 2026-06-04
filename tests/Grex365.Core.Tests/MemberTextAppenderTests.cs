using FluentAssertions;
using Grex365.Core.Groups;

namespace Grex365.Core.Tests;

public class MemberTextAppenderTests
{
    [Fact]
    public void Append_EmptyExisting_ReturnsEntry()
    {
        MemberTextAppender.Append("", "a@x.com").Should().Be("a@x.com");
    }

    [Fact]
    public void Append_NullExisting_ReturnsEntry()
    {
        MemberTextAppender.Append(null, "a@x.com").Should().Be("a@x.com");
    }

    [Fact]
    public void Append_NewEntry_AppendsOnNewLine()
    {
        MemberTextAppender.Append("a@x.com", "b@x.com")
            .Should().Be("a@x.com" + Environment.NewLine + "b@x.com");
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    public void Append_BlankEntry_LeavesTextUnchanged(string? entry)
    {
        MemberTextAppender.Append("a@x.com", entry).Should().Be("a@x.com");
    }

    [Fact]
    public void Append_Duplicate_CaseInsensitive_Unchanged()
    {
        MemberTextAppender.Append("A@X.com", "a@x.com").Should().Be("A@X.com");
    }

    [Fact]
    public void Append_DuplicateAcrossCommaDelimiter_Unchanged()
    {
        // GroupsViewModel splits on , ; and newlines — a dupe must be caught regardless of delimiter.
        MemberTextAppender.Append("a@x.com, b@x.com", "b@x.com").Should().Be("a@x.com, b@x.com");
    }

    [Fact]
    public void Append_TrailingNewline_NoDoubleBlankLine()
    {
        MemberTextAppender.Append("a@x.com\n", "b@x.com").Should().Be("a@x.com\nb@x.com");
    }

    [Fact]
    public void Append_TrimsEntry()
    {
        MemberTextAppender.Append("a@x.com", "  b@x.com  ")
            .Should().Be("a@x.com" + Environment.NewLine + "b@x.com");
    }
}
