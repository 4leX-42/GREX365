using FluentAssertions;
using Grex365.Core.Csv;

namespace Grex365.Core.Tests;

public class CsvEscaperTests
{
    [Fact]
    public void Escape_Null_ReturnsEmpty()
    {
        CsvEscaper.Escape(null).Should().Be(string.Empty);
    }

    [Fact]
    public void Escape_Empty_ReturnsEmpty()
    {
        CsvEscaper.Escape(string.Empty).Should().Be(string.Empty);
    }

    [Fact]
    public void Escape_SimpleString_ReturnedAsIs()
    {
        CsvEscaper.Escape("hello").Should().Be("hello");
    }

    [Fact]
    public void Escape_WithSpaces_NotQuoted()
    {
        CsvEscaper.Escape("hello world").Should().Be("hello world");
    }

    [Fact]
    public void Escape_WithComma_Quoted()
    {
        CsvEscaper.Escape("a,b").Should().Be("\"a,b\"");
    }

    [Fact]
    public void Escape_WithDoubleQuote_QuotedAndEscaped()
    {
        CsvEscaper.Escape("she said \"hi\"").Should().Be("\"she said \"\"hi\"\"\"");
    }

    [Fact]
    public void Escape_WithNewline_Quoted()
    {
        CsvEscaper.Escape("line1\nline2").Should().Be("\"line1\nline2\"");
    }

    [Fact]
    public void Escape_WithCarriageReturn_Quoted()
    {
        CsvEscaper.Escape("line1\rline2").Should().Be("\"line1\rline2\"");
    }

    [Fact]
    public void Escape_WithCRLF_Quoted()
    {
        CsvEscaper.Escape("line1\r\nline2").Should().Be("\"line1\r\nline2\"");
    }

    [Fact]
    public void Escape_AllSpecials_QuotedAndDoubled()
    {
        CsvEscaper.Escape("a,b\"c\nd").Should().Be("\"a,b\"\"c\nd\"");
    }

    [Theory]
    [InlineData(",")]
    [InlineData("\"")]
    [InlineData("\n")]
    [InlineData("\r")]
    public void Escape_OnlySpecialChar_Quoted(string s)
    {
        var result = CsvEscaper.Escape(s);
        result.Should().StartWith("\"").And.EndWith("\"");
    }

    [Fact]
    public void Escape_QuoteOnly_DoubledAndWrapped()
    {
        CsvEscaper.Escape("\"").Should().Be("\"\"\"\"");
    }

    [Fact]
    public void Escape_TabIsNotSpecial()
    {
        CsvEscaper.Escape("a\tb").Should().Be("a\tb");
    }
}
