using FluentAssertions;
using Grex365.Core.Models;
using Grex365.Core.Users;

namespace Grex365.Core.Tests;

public class BulkUserActionParserTests
{
    [Theory]
    [InlineData("enable", BulkUserActionKind.Enable)]
    [InlineData("Enable", BulkUserActionKind.Enable)]
    [InlineData("ENABLE", BulkUserActionKind.Enable)]
    [InlineData("disable", BulkUserActionKind.Disable)]
    [InlineData("remove-licenses", BulkUserActionKind.RemoveLicenses)]
    public void Recognizes_Known_Actions(string input, BulkUserActionKind expected)
    {
        BulkUserActionParser.Parse(input).Kind.Should().Be(expected);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    [InlineData("set-password")]
    public void Unknown_Or_Empty_Maps_To_Unknown(string? input)
    {
        BulkUserActionParser.Parse(input).Kind.Should().Be(BulkUserActionKind.Unknown);
    }

    [Theory]
    [InlineData("assign:ENTERPRISEPACK", "ENTERPRISEPACK")]
    [InlineData("assign:e3", "E3")]
    [InlineData("  assign:E5  ", "E5")]
    public void Parses_Assign_With_Sku(string input, string expectedSku)
    {
        var p = BulkUserActionParser.Parse(input);
        p.Kind.Should().Be(BulkUserActionKind.AssignLicense);
        p.SkuPartNumber.Should().Be(expectedSku);
    }

    [Theory]
    [InlineData("assign:")]
    [InlineData("assign:  ")]
    public void Assign_Without_Sku_Has_Null_PartNumber(string input)
    {
        var p = BulkUserActionParser.Parse(input);
        p.Kind.Should().Be(BulkUserActionKind.AssignLicense);
        p.SkuPartNumber.Should().BeNull();
    }

    [Fact]
    public void FindByPartNumber_Hits_CaseInsensitive()
    {
        var skus = new[]
        {
            new SkuInfo(Guid.NewGuid(), "ENTERPRISEPACK", 10, 5),
            new SkuInfo(Guid.NewGuid(), "FLOW_FREE", 100, 1),
        };
        BulkUserActionParser.FindByPartNumber(skus, "enterprisepack")
            .Should().NotBeNull().And.Match<SkuInfo>(s => s.SkuPartNumber == "ENTERPRISEPACK");
    }

    [Fact]
    public void FindByPartNumber_Miss_Returns_Null()
    {
        var skus = new[] { new SkuInfo(Guid.NewGuid(), "E3", 1, 0) };
        BulkUserActionParser.FindByPartNumber(skus, "E5").Should().BeNull();
    }

    [Fact]
    public void FindByPartNumber_NullOrEmpty_Returns_Null()
    {
        var skus = new[] { new SkuInfo(Guid.NewGuid(), "E3", 1, 0) };
        BulkUserActionParser.FindByPartNumber(skus, null).Should().BeNull();
        BulkUserActionParser.FindByPartNumber(skus, "  ").Should().BeNull();
    }
}
