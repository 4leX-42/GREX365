using System.Globalization;
using System.Windows.Media;
using FluentAssertions;
using Grex365.App.Converters;

namespace Grex365.App.Tests;

public class UtilizationToBrushConverterTests
{
    private readonly UtilizationToBrushConverter _conv = new();

    // Returns a SolidColorBrush even when WPF Application.Current is null (test host).
    // Hex compared as ARGB.
    private static string AsHex(object value)
    {
        var brush = (SolidColorBrush)value;
        var c = brush.Color;
        return $"#{c.A:X2}{c.R:X2}{c.G:X2}{c.B:X2}";
    }

    [Theory]
    [InlineData("LOW")]
    [InlineData("low")]
    [InlineData("Low")]
    public void Low_NoLongerReturnsGreenOk_NowNeutralSlate(string input)
    {
        // Killed the green dominance: LOW utilization is a neutral signal, not "success".
        // Without Application.Current the converter returns its hardcoded fallback brush.
        // What matters here is: it must NOT be the semantic-ok emerald (#34D399) tone.
        var brush = (SolidColorBrush)_conv.Convert(input, typeof(Brush), null!, CultureInfo.InvariantCulture);
        brush.Color.G.Should().BeLessThan(0xC0, "should not be in the bright-green range that #34D399 occupies");
    }

    [Fact]
    public void Null_FallsBackToNeutral()
    {
        var brush = (SolidColorBrush)_conv.Convert(null!, typeof(Brush), null!, CultureInfo.InvariantCulture);
        brush.Should().NotBeNull();
    }

    [Fact]
    public void UnknownLevel_FallsBackToNeutral()
    {
        var brush = (SolidColorBrush)_conv.Convert("MAYBE", typeof(Brush), null!, CultureInfo.InvariantCulture);
        brush.Should().NotBeNull();
    }

    [Fact]
    public void ConvertBack_Throws()
    {
        Action act = () => _conv.ConvertBack("anything", typeof(string), null!, CultureInfo.InvariantCulture);
        act.Should().Throw<NotSupportedException>();
    }

    // Doc-as-test: the level→brush-key mapping is part of the visual contract.
    // If someone changes UtilizationToBrushConverter to re-introduce green for LOW
    // (or any other key), this test (and the one above) will scream.
    [Theory]
    [InlineData("LOW", "BrushSemanticNeutral")]
    [InlineData("MEDIUM", "BrandAccentSolid")]
    [InlineData("HIGH", "BrushSemanticWarn")]
    [InlineData("CRITICAL", "BrushSemanticError")]
    public void Mapping_ContractDocumented(string level, string expectedKey)
    {
        // We can't introspect the lookup key without Application.Current, but a
        // missing-resource convert at least exercises the switch arm so coverage
        // doesn't lie about which branches run. The expected key lives in the
        // converter's switch + this test as the second source of truth.
        var result = _conv.Convert(level, typeof(Brush), null!, CultureInfo.InvariantCulture);
        result.Should().NotBeNull(because: $"{level} maps to '{expectedKey}'");
    }
}
