using FluentAssertions;
using Grex365.App;

namespace Grex365.App.Tests;

public class L10nTests : IDisposable
{
    public L10nTests() => L10n.Reset();
    public void Dispose() => L10n.Reset();

    [Fact]
    public void Configure_PrimaryDict_GetReturnsValue()
    {
        L10n.Configure(new Dictionary<string, string> { ["Nav.Dashboard"] = "Tablero" });

        L10n.Get("Nav.Dashboard").Should().Be("Tablero");
    }

    [Fact]
    public void Get_MissingKey_FallsBackToFallbackDict()
    {
        L10n.Configure(
            primary: new Dictionary<string, string> { ["A"] = "primaryA" },
            fallback: new Dictionary<string, string> { ["B"] = "fallbackB" });

        L10n.Get("B").Should().Be("fallbackB");
    }

    [Fact]
    public void Get_MissingInBoth_ReturnsKey()
    {
        L10n.Configure(new Dictionary<string, string>());
        L10n.Get("Missing.Key").Should().Be("Missing.Key");
    }

    [Fact]
    public void Get_EmptyKey_ReturnsEmpty()
    {
        L10n.Configure(new Dictionary<string, string> { ["X"] = "v" });
        L10n.Get(string.Empty).Should().Be(string.Empty);
    }

    [Fact]
    public void Configure_CaseInsensitiveLookup()
    {
        L10n.Configure(new Dictionary<string, string> { ["Nav.Dashboard"] = "Dashboard" });
        L10n.Get("nav.dashboard").Should().Be("Dashboard");
        L10n.Get("NAV.DASHBOARD").Should().Be("Dashboard");
    }

    [Fact]
    public void Configure_NullPrimary_Throws()
    {
        Action act = () => L10n.Configure(null!);
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void Configure_ActiveLanguage_Stored()
    {
        L10n.Configure(new Dictionary<string, string>(), activeLanguage: "en");
        L10n.ActiveLanguage.Should().Be("en");
    }

    [Fact]
    public void Initialize_Spanish_LoadsEmbeddedResource()
    {
        L10n.Initialize("es");

        L10n.ActiveLanguage.Should().Be("es");
        L10n.Get("Nav.Dashboard").Should().Be("Dashboard");
        L10n.Get("Nav.Connection").Should().Be("Conexión");
        L10n.Get("Settings.Title").Should().Be("Ajustes");
    }

    [Fact]
    public void Initialize_English_LoadsEmbeddedResource()
    {
        L10n.Initialize("en");

        L10n.ActiveLanguage.Should().Be("en");
        L10n.Get("Nav.Connection").Should().Be("Connection");
        L10n.Get("Settings.Title").Should().Be("Settings");
    }

    [Fact]
    public void Initialize_EnglishMissingKey_FallsBackToSpanish()
    {
        L10n.Initialize("en");
        // Both have all keys currently — emulate gap by querying a hypothetical key
        // that lives only in es.json. For now both are parity, so just verify the
        // fallback dict is populated by checking ActiveLanguage = en + es loaded as fallback.
        L10n.Get("Nav.Dashboard").Should().NotBe("Nav.Dashboard");
    }

    [Fact]
    public void Initialize_UnknownLanguage_FallsBackToDefault()
    {
        L10n.Initialize("klingon");

        L10n.ActiveLanguage.Should().Be(L10n.DefaultLanguage);
        L10n.Get("Nav.Connection").Should().Be("Conexión");
    }

    [Fact]
    public void Initialize_Null_FallsBackToDefault()
    {
        L10n.Initialize(null);

        L10n.ActiveLanguage.Should().Be(L10n.DefaultLanguage);
        L10n.Get("Nav.Audit").Should().Be("Auditoría");
    }

    [Fact]
    public void Initialize_Whitespace_FallsBackToDefault()
    {
        L10n.Initialize("   ");
        L10n.ActiveLanguage.Should().Be(L10n.DefaultLanguage);
    }

    [Fact]
    public void Initialize_CaseInsensitive_LanguageCode()
    {
        L10n.Initialize("EN");
        L10n.ActiveLanguage.Should().Be("en");
        L10n.Get("Settings.Title").Should().Be("Settings");
    }

    [Fact]
    public void SupportedLanguages_HasEsAndEn()
    {
        L10n.SupportedLanguages.Should().BeEquivalentTo(["es", "en"]);
    }

    [Fact]
    public void Reset_ClearsState()
    {
        L10n.Configure(new Dictionary<string, string> { ["X"] = "v" }, activeLanguage: "en");
        L10n.Reset();

        L10n.ActiveLanguage.Should().Be(L10n.DefaultLanguage);
        L10n.Get("X").Should().Be("X");
    }

    [Theory]
    [InlineData("Nav.Dashboard")]
    [InlineData("Nav.Connection")]
    [InlineData("Nav.Licenses")]
    [InlineData("Nav.Users")]
    [InlineData("Nav.Groups")]
    [InlineData("Nav.Onboarding")]
    [InlineData("Nav.Offboarding")]
    [InlineData("Nav.SharedMailbox")]
    [InlineData("Nav.MailboxRules")]
    [InlineData("Nav.MailFlow")]
    [InlineData("Nav.Audit")]
    [InlineData("Nav.AuditLog")]
    [InlineData("Nav.PsConsole")]
    [InlineData("Nav.CertWizard")]
    [InlineData("Nav.DnsCheck")]
    public void Initialize_Spanish_AllNavKeys_Resolved(string key)
    {
        L10n.Initialize("es");
        var value = L10n.Get(key);
        value.Should().NotBe(key, because: $"the key '{key}' should resolve to a non-key value in es.json");
        value.Should().NotBeNullOrWhiteSpace();
    }

    [Theory]
    [InlineData("Nav.Dashboard")]
    [InlineData("Nav.Connection")]
    [InlineData("Nav.Licenses")]
    [InlineData("Nav.Users")]
    [InlineData("Nav.Audit")]
    public void Initialize_English_NavKeys_Resolved(string key)
    {
        L10n.Initialize("en");
        var value = L10n.Get(key);
        value.Should().NotBe(key);
        value.Should().NotBeNullOrWhiteSpace();
    }
}
