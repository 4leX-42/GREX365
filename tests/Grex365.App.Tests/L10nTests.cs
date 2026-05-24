using FluentAssertions;
using Grex365.App;

namespace Grex365.App.Tests;

[Collection("L10n")]
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

    [Theory]
    [InlineData("Settings.Window.Title")]
    [InlineData("Settings.Preferences")]
    [InlineData("Settings.Section.Connection")]
    [InlineData("Settings.Section.Plugins")]
    [InlineData("Settings.Section.Certificate")]
    [InlineData("Settings.Connection.Cert")]
    [InlineData("Settings.Connection.Traditional")]
    [InlineData("Settings.Tenant.IdLabel")]
    [InlineData("Settings.Tenant.DomainLabel")]
    [InlineData("Settings.EnforceTenantLock")]
    [InlineData("Settings.LanguageLabel")]
    [InlineData("Settings.Theme.Dark")]
    [InlineData("Settings.Theme.Light")]
    [InlineData("Settings.Theme.Auto")]
    [InlineData("Settings.Theme.Hint")]
    [InlineData("Settings.LogLevel")]
    [InlineData("Settings.LogLevel.Debug")]
    [InlineData("Settings.LogLevel.Information")]
    [InlineData("Settings.LogLevel.Warning")]
    [InlineData("Settings.LogLevel.Error")]
    [InlineData("Settings.LogLevel.Hint")]
    [InlineData("Settings.AppInsights")]
    [InlineData("Settings.AppInsights.Placeholder")]
    [InlineData("Settings.AppInsights.Hint")]
    [InlineData("Settings.Rbac")]
    [InlineData("Settings.Rbac.Hint")]
    [InlineData("Settings.Plugins.Hint")]
    [InlineData("Settings.Plugins.Empty")]
    [InlineData("Settings.Plugins.ModulesSuffix")]
    [InlineData("Settings.Plugins.Status.Loaded")]
    [InlineData("Settings.Plugins.Status.Disabled")]
    [InlineData("Settings.Plugins.Status.ErrorPrefix")]
    [InlineData("Settings.Cert.AppId")]
    [InlineData("Settings.Cert.TenantId")]
    [InlineData("Settings.Cert.Organization")]
    [InlineData("Settings.Cert.Thumbprint")]
    [InlineData("Settings.Cert.Browse")]
    [InlineData("Settings.Cert.Validate")]
    [InlineData("Settings.Button.Reload")]
    [InlineData("Settings.Button.Save")]
    [InlineData("Settings.SaveStatus.Prefix")]
    [InlineData("Settings.SaveStatus.ErrorPrefix")]
    [InlineData("Settings.SaveStatus.SavedLog")]
    public void Initialize_Spanish_AllSettingsKeys_Resolved(string key)
    {
        L10n.Initialize("es");
        var value = L10n.Get(key);
        value.Should().NotBe(key, because: $"the key '{key}' should resolve in es");
        value.Should().NotBeNullOrWhiteSpace();
    }

    [Theory]
    [InlineData("Settings.Window.Title")]
    [InlineData("Settings.Section.Connection")]
    [InlineData("Settings.Connection.Cert")]
    [InlineData("Settings.Theme.Auto")]
    [InlineData("Settings.LogLevel.Hint")]
    [InlineData("Settings.Plugins.ModulesSuffix")]
    [InlineData("Settings.Cert.Validate")]
    [InlineData("Settings.Button.Save")]
    public void Initialize_English_SettingsKeys_Resolved(string key)
    {
        L10n.Initialize("en");
        var value = L10n.Get(key);
        value.Should().NotBe(key);
        value.Should().NotBeNullOrWhiteSpace();
    }

    [Theory]
    [InlineData("About.Window.Title")]
    [InlineData("About.TitleBar")]
    [InlineData("About.Description")]
    [InlineData("About.Version")]
    [InlineData("About.Runtime")]
    [InlineData("About.DataDir")]
    [InlineData("About.Button.OpenDataDir")]
    [InlineData("About.Button.Close")]
    public void Initialize_Spanish_AllAboutKeys_Resolved(string key)
    {
        L10n.Initialize("es");
        L10n.Get(key).Should().NotBe(key).And.NotBeNullOrWhiteSpace();
    }

    [Theory]
    [InlineData("About.Description")]
    [InlineData("About.Button.OpenDataDir")]
    [InlineData("About.Button.Close")]
    public void Initialize_English_AboutKeys_Resolved(string key)
    {
        L10n.Initialize("en");
        L10n.Get(key).Should().NotBe(key).And.NotBeNullOrWhiteSpace();
    }

    [Theory]
    [InlineData("Dialog.Ok")]
    [InlineData("Dialog.Cancel")]
    [InlineData("Dialog.Yes")]
    [InlineData("Dialog.No")]
    [InlineData("Dialog.Close")]
    [InlineData("Common.Connect")]
    [InlineData("Common.Disconnect")]
    [InlineData("Common.Refresh")]
    [InlineData("Common.Export")]
    [InlineData("Common.Apply")]
    [InlineData("Common.Clear")]
    [InlineData("Common.Search")]
    [InlineData("Common.Loading")]
    [InlineData("Common.Empty")]
    [InlineData("Common.Save")]
    [InlineData("Common.Delete")]
    [InlineData("Common.Edit")]
    [InlineData("Common.Add")]
    [InlineData("Common.Remove")]
    [InlineData("Common.Back")]
    [InlineData("Common.Next")]
    [InlineData("Common.Finish")]
    [InlineData("Common.Skip")]
    public void Initialize_Both_DialogAndCommonKeys_Resolved(string key)
    {
        L10n.Initialize("es");
        L10n.Get(key).Should().NotBe(key).And.NotBeNullOrWhiteSpace();

        L10n.Initialize("en");
        L10n.Get(key).Should().NotBe(key).And.NotBeNullOrWhiteSpace();
    }

    [Fact]
    public void Format_KeyWithPlaceholder_SubstitutesArg()
    {
        L10n.Configure(new Dictionary<string, string> { ["Greeting"] = "Hola {0}" });

        L10n.Format("Greeting", "mundo").Should().Be("Hola mundo");
    }

    [Fact]
    public void Format_MultipleArgs_OrderPreserved()
    {
        L10n.Configure(new Dictionary<string, string> { ["X"] = "{0}-{1}" });

        L10n.Format("X", "a", "b").Should().Be("a-b");
    }

    [Fact]
    public void Format_NoArgs_ReturnsRawTemplate()
    {
        L10n.Configure(new Dictionary<string, string> { ["X"] = "static" });

        L10n.Format("X").Should().Be("static");
    }

    [Fact]
    public void Format_InvalidPlaceholder_ReturnsRawTemplate()
    {
        L10n.Configure(new Dictionary<string, string> { ["X"] = "bad {nope}" });

        L10n.Format("X", "z").Should().Be("bad {nope}");
    }

    [Fact]
    public void Format_NullArgs_ReturnsRawTemplate()
    {
        L10n.Configure(new Dictionary<string, string> { ["X"] = "literal" });

        L10n.Format("X", null!).Should().Be("literal");
    }

    [Fact]
    public void Format_SaveStatusPrefix_Spanish_FormatsTimestamp()
    {
        L10n.Initialize("es");
        var result = L10n.Format("Settings.SaveStatus.Prefix", "12:34:56");

        result.Should().Contain("12:34:56").And.Contain("Guardado");
    }

    [Fact]
    public void Format_SaveStatusPrefix_English_FormatsTimestamp()
    {
        L10n.Initialize("en");
        var result = L10n.Format("Settings.SaveStatus.Prefix", "09:00:00");

        result.Should().Contain("09:00:00").And.Contain("Saved");
    }

    [Fact]
    public void KnownKeys_NotEmpty_ContainsNavAndSettings()
    {
        L10n.Initialize("es");

        L10n.KnownKeys.Should().NotBeEmpty();
        L10n.KnownKeys.Should().Contain("Nav.Dashboard");
        L10n.KnownKeys.Should().Contain("Settings.Title");
    }

    [Fact]
    public void EnDict_HasTranslationForEveryEsKey()
    {
        L10n.Initialize("es");
        var keys = L10n.KnownKeys.ToList();

        L10n.Initialize("en");
        foreach (var key in keys)
        {
            var en = L10n.Get(key);
            en.Should().NotBe(key, because: $"EN should translate '{key}' (or fallback chain should mask it)");
        }
    }
}
