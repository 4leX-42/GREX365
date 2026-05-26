using FluentAssertions;
using Grex365.Core.Models;
using Grex365.Core.Preferences;

namespace Grex365.Core.Tests;

public class PreferencesStoreTests : IDisposable
{
    private readonly string _tempDir;

    public PreferencesStoreTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "grex365-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    [Fact]
    public async Task Load_NonExistent_Returns_Defaults()
    {
        var store = new JsonPreferencesStore(_tempDir);
        var prefs = await store.LoadAsync();
        prefs.Should().NotBeNull();
        prefs.ConnectionMethod.Should().BeNull();
        prefs.Role.Should().Be("operator");
    }

    [Fact]
    public async Task Save_Then_Load_Roundtrip()
    {
        var store = new JsonPreferencesStore(_tempDir);
        var prefs = new UserPreferences
        {
            ConnectionMethod = "cert",
            EnforceTenantLock = true,
            ExpectedTenantId = "abc-123"
        };
        await store.SaveAsync(prefs);

        var loaded = await store.LoadAsync();
        loaded.ConnectionMethod.Should().Be("cert");
        loaded.EnforceTenantLock.Should().BeTrue();
        loaded.ExpectedTenantId.Should().Be("abc-123");
    }

    [Fact]
    public async Task DisabledPluginAssemblies_Roundtrip()
    {
        var store = new JsonPreferencesStore(_tempDir);
        var prefs = new UserPreferences
        {
            DisabledPluginAssemblies = new List<string> { "PluginA.dll", "PluginB.dll" }
        };
        await store.SaveAsync(prefs);

        var loaded = await store.LoadAsync();
        loaded.DisabledPluginAssemblies.Should().BeEquivalentTo("PluginA.dll", "PluginB.dll");
    }

    [Fact]
    public async Task DisabledPluginAssemblies_Defaults_To_Empty_When_Missing()
    {
        var path = Path.Combine(_tempDir, "user_preferences.json");
        await File.WriteAllTextAsync(path, "{\"Theme\":\"Light\"}");

        var store = new JsonPreferencesStore(_tempDir);
        var loaded = await store.LoadAsync();
        loaded.Theme.Should().Be("Light");
        loaded.DisabledPluginAssemblies.Should().NotBeNull().And.BeEmpty();
    }

    [Fact]
    public async Task LogLevel_Roundtrip()
    {
        var store = new JsonPreferencesStore(_tempDir);
        var prefs = new UserPreferences { LogLevel = "Debug" };
        await store.SaveAsync(prefs);

        var loaded = await store.LoadAsync();
        loaded.LogLevel.Should().Be("Debug");
    }

    [Fact]
    public async Task LogLevel_Defaults_To_Information_When_Missing()
    {
        var path = Path.Combine(_tempDir, "user_preferences.json");
        await File.WriteAllTextAsync(path, "{\"Theme\":\"Dark\"}");

        var store = new JsonPreferencesStore(_tempDir);
        var loaded = await store.LoadAsync();
        loaded.LogLevel.Should().Be("Information");
    }

    [Fact]
    public async Task CertConfig_Roundtrip()
    {
        var store = new JsonCertConfigStore(_tempDir);
        var cfg = new CertConfig("app-1", "tenant-1", "org.onmicrosoft.com", "ABCDEF");
        await store.SaveAsync(cfg);

        var loaded = await store.LoadAsync();
        loaded.Should().NotBeNull();
        loaded!.AppId.Should().Be("app-1");
        loaded.CertThumbprint.Should().Be("ABCDEF");
    }

    [Fact]
    public async Task Load_CorruptJson_ReturnsDefaults()
    {
        var path = Path.Combine(_tempDir, "user_preferences.json");
        await File.WriteAllTextAsync(path, "{this is not valid json");

        var store = new JsonPreferencesStore(_tempDir);
        var loaded = await store.LoadAsync();

        loaded.Should().NotBeNull();
        loaded.ConnectionMethod.Should().BeNull();
        loaded.Role.Should().Be("operator");
    }

    [Fact]
    public async Task Load_CorruptJson_QuarantinesOriginalFile()
    {
        var path = Path.Combine(_tempDir, "user_preferences.json");
        var corruptContent = "{\"Theme\":\"Dark\"  truncated";
        await File.WriteAllTextAsync(path, corruptContent);

        var store = new JsonPreferencesStore(_tempDir);
        _ = await store.LoadAsync();

        File.Exists(path).Should().BeFalse("the corrupt file should have been moved aside");

        var backup = Directory.EnumerateFiles(_tempDir, "user_preferences.json.corrupted-*.bak").Single();
        var preserved = await File.ReadAllTextAsync(backup);
        preserved.Should().Be(corruptContent, "quarantine preserves bytes for user recovery");
    }

    [Fact]
    public async Task Load_CorruptThenSave_OverwritesFreshFile()
    {
        var path = Path.Combine(_tempDir, "user_preferences.json");
        await File.WriteAllTextAsync(path, "{broken");

        var store = new JsonPreferencesStore(_tempDir);
        var loaded = await store.LoadAsync();
        loaded.Theme = "Light";
        await store.SaveAsync(loaded);

        var reloaded = await store.LoadAsync();
        reloaded.Theme.Should().Be("Light");
    }

    [Fact]
    public async Task CertConfig_Load_CorruptJson_ReturnsNull_AndQuarantines()
    {
        var path = Path.Combine(_tempDir, "exo-app-params.json");
        await File.WriteAllTextAsync(path, "{not valid");

        var store = new JsonCertConfigStore(_tempDir);
        var loaded = await store.LoadAsync();

        loaded.Should().BeNull();
        File.Exists(path).Should().BeFalse();
        Directory.EnumerateFiles(_tempDir, "exo-app-params.json.corrupted-*.bak").Should().HaveCount(1);
    }

    [Fact]
    public async Task Load_EmptyFile_ReturnsDefaults_NoQuarantine()
    {
        // Empty file is treated as JsonException (empty input) — should fall through
        // to defaults like any other corrupt content.
        var path = Path.Combine(_tempDir, "user_preferences.json");
        await File.WriteAllTextAsync(path, "");

        var store = new JsonPreferencesStore(_tempDir);
        var loaded = await store.LoadAsync();

        loaded.Should().NotBeNull();
        loaded.Role.Should().Be("operator");
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(_tempDir, recursive: true);
        }
        catch
        {
            // ignore
        }
    }
}
