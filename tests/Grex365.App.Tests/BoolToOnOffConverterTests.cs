using System.Globalization;
using FluentAssertions;
using Grex365.App;
using Grex365.App.Converters;

namespace Grex365.App.Tests;

[Collection("L10n")]
public class BoolToOnOffConverterTests : IDisposable
{
    private readonly BoolToOnOffConverter _conv = new();

    public BoolToOnOffConverterTests() => L10n.Reset();
    public void Dispose() => L10n.Reset();

    [Fact]
    public void NoParameter_TrueValue_ReturnsCommonConnectedFromL10n()
    {
        L10n.Initialize("es");
        var result = _conv.Convert(true, typeof(string), null!, CultureInfo.InvariantCulture);
        result.Should().Be("conectado");
    }

    [Fact]
    public void NoParameter_FalseValue_ReturnsCommonDisconnectedFromL10n()
    {
        L10n.Initialize("es");
        var result = _conv.Convert(false, typeof(string), null!, CultureInfo.InvariantCulture);
        result.Should().Be("desconectado");
    }

    [Fact]
    public void NoParameter_English_ReturnsConnectedDisconnectedInEnglish()
    {
        L10n.Initialize("en");
        _conv.Convert(true, typeof(string), null!, CultureInfo.InvariantCulture).Should().Be("connected");
        _conv.Convert(false, typeof(string), null!, CultureInfo.InvariantCulture).Should().Be("disconnected");
    }

    [Fact]
    public void DottedParameter_BothPartsResolvedAsL10nKeys_Spanish()
    {
        L10n.Initialize("es");
        _conv.Convert(true, typeof(string), "Status.Enabled/Status.Disabled", CultureInfo.InvariantCulture)
            .Should().Be("Habilitado");
        _conv.Convert(false, typeof(string), "Status.Enabled/Status.Disabled", CultureInfo.InvariantCulture)
            .Should().Be("Deshabilitado");
    }

    [Fact]
    public void DottedParameter_BothPartsResolvedAsL10nKeys_English()
    {
        L10n.Initialize("en");
        _conv.Convert(true, typeof(string), "Status.Enabled/Status.Disabled", CultureInfo.InvariantCulture)
            .Should().Be("Enabled");
        _conv.Convert(false, typeof(string), "Status.Enabled/Status.Disabled", CultureInfo.InvariantCulture)
            .Should().Be("Disabled");
    }

    [Fact]
    public void DottedParameter_UserTypeGuestMember_Spanish()
    {
        L10n.Initialize("es");
        _conv.Convert(true, typeof(string), "UserType.Guest/UserType.Member", CultureInfo.InvariantCulture)
            .Should().Be("Invitado");
        _conv.Convert(false, typeof(string), "UserType.Guest/UserType.Member", CultureInfo.InvariantCulture)
            .Should().Be("Miembro");
    }

    [Fact]
    public void PlainParameter_NoDot_TreatedAsLiteralStrings()
    {
        // Backwards compat: ad-hoc captions without L10n keys still render literally.
        L10n.Initialize("es");
        _conv.Convert(true, typeof(string), "ON/OFF", CultureInfo.InvariantCulture).Should().Be("ON");
        _conv.Convert(false, typeof(string), "ON/OFF", CultureInfo.InvariantCulture).Should().Be("OFF");
    }

    [Fact]
    public void MixedParameter_DottedAndPlain_DottedResolvedPlainLiteral()
    {
        L10n.Initialize("es");
        // Edge case: one side is a key, the other is a literal.
        _conv.Convert(true, typeof(string), "Status.Enabled/raw-off", CultureInfo.InvariantCulture)
            .Should().Be("Habilitado");
        _conv.Convert(false, typeof(string), "Status.Enabled/raw-off", CultureInfo.InvariantCulture)
            .Should().Be("raw-off");
    }

    [Fact]
    public void UnknownKey_FallsBackToKeyItself()
    {
        // L10n.Get returns the key if not found — convert preserves that contract.
        L10n.Initialize("es");
        _conv.Convert(true, typeof(string), "Unknown.Key/Other.Key", CultureInfo.InvariantCulture)
            .Should().Be("Unknown.Key");
    }

    [Fact]
    public void NonBoolValue_TreatedAsFalse()
    {
        L10n.Initialize("es");
        _conv.Convert(null!, typeof(string), null!, CultureInfo.InvariantCulture).Should().Be("desconectado");
        _conv.Convert(42, typeof(string), null!, CultureInfo.InvariantCulture).Should().Be("desconectado");
        _conv.Convert("string", typeof(string), null!, CultureInfo.InvariantCulture).Should().Be("desconectado");
    }

    [Fact]
    public void ConvertBack_Throws()
    {
        Action act = () => _conv.ConvertBack("anything", typeof(bool), null!, CultureInfo.InvariantCulture);
        act.Should().Throw<NotSupportedException>();
    }
}
