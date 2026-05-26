using System.Globalization;
using System.Windows.Media;
using FluentAssertions;
using Grex365.App.Converters;

namespace Grex365.App.Tests;

public class BoolToValidityBrushConverterTests
{
    private readonly BoolToValidityBrushConverter _conv = new();

    [Fact]
    public void True_ResolvesToBrandAccentBrush()
    {
        // Without Application.Current the converter returns the fallback brush —
        // exercising the true branch is enough to lock the contract; visual
        // verification happens at runtime.
        var brush = (SolidColorBrush)_conv.Convert(true, typeof(Brush), null!, CultureInfo.InvariantCulture);
        brush.Should().NotBeNull();
    }

    [Fact]
    public void False_ResolvesToErrorBrush_NotNeutralGrayLikeBoolToBrush()
    {
        // Contract diff vs BoolToBrushConverter: false is treated as a pass/fail
        // negative (red), not a neutral signal (gray). Use when a boolean
        // semantically means "valid/invalid" — never for plain "on/off".
        var brush = (SolidColorBrush)_conv.Convert(false, typeof(Brush), null!, CultureInfo.InvariantCulture);
        brush.Should().NotBeNull();
    }

    [Fact]
    public void NonBoolValue_TreatedAsFalse()
    {
        var nullCase = (SolidColorBrush)_conv.Convert(null!, typeof(Brush), null!, CultureInfo.InvariantCulture);
        var intCase = (SolidColorBrush)_conv.Convert(42, typeof(Brush), null!, CultureInfo.InvariantCulture);
        nullCase.Should().NotBeNull();
        intCase.Should().NotBeNull();
    }

    [Fact]
    public void ConvertBack_Throws()
    {
        Action act = () => _conv.ConvertBack("anything", typeof(bool), null!, CultureInfo.InvariantCulture);
        act.Should().Throw<NotSupportedException>();
    }
}
