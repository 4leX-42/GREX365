using FluentAssertions;
using Grex365.App;
using Grex365.App.Xaml;

namespace Grex365.App.Tests;

[Collection("L10n")]
public class L10nExtensionTests : IDisposable
{
    public L10nExtensionTests() => L10n.Reset();
    public void Dispose() => L10n.Reset();

    [Fact]
    public void ProvideValue_KnownKey_ReturnsTranslation()
    {
        L10n.Configure(new Dictionary<string, string> { ["Settings.Title"] = "Ajustes" });
        var ext = new L10nExtension("Settings.Title");

        var result = ext.ProvideValue(null!);

        result.Should().Be("Ajustes");
    }

    [Fact]
    public void ProvideValue_NullKey_ReturnsEmpty()
    {
        L10n.Initialize("es");
        var ext = new L10nExtension();

        ext.ProvideValue(null!).Should().Be(string.Empty);
    }

    [Fact]
    public void ProvideValue_EmptyKey_ReturnsEmpty()
    {
        L10n.Initialize("es");
        var ext = new L10nExtension(string.Empty);

        ext.ProvideValue(null!).Should().Be(string.Empty);
    }

    [Fact]
    public void ProvideValue_UnknownKey_ReturnsKey()
    {
        L10n.Configure(new Dictionary<string, string>());
        var ext = new L10nExtension("Missing.X");

        ext.ProvideValue(null!).Should().Be("Missing.X");
    }

    [Fact]
    public void Constructor_KeyProperty_Settable()
    {
        L10n.Configure(new Dictionary<string, string> { ["A"] = "alpha" });
        var ext = new L10nExtension { Key = "A" };

        ext.ProvideValue(null!).Should().Be("alpha");
    }

    [Fact]
    public void Constructor_DefaultsToNullKey()
    {
        L10n.Initialize("es");
        var ext = new L10nExtension();

        ext.Key.Should().BeNull();
        ext.ProvideValue(null!).Should().Be(string.Empty);
    }
}
