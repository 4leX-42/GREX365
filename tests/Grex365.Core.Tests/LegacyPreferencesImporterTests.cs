using System.Text.Json;
using FluentAssertions;
using Grex365.Core.Models;
using Grex365.Core.Preferences;

namespace Grex365.Core.Tests;

public class LegacyPreferencesImporterTests : IDisposable
{
    private readonly string _root;
    private readonly string _legacy;
    private readonly string _target;

    public LegacyPreferencesImporterTests()
    {
        _root = Path.Combine(Path.GetTempPath(), "grex365-legacy-" + Guid.NewGuid().ToString("N"));
        _legacy = Path.Combine(_root, "legacy");
        _target = Path.Combine(_root, "target");
        Directory.CreateDirectory(_legacy);
        Directory.CreateDirectory(_target);
    }

    [Fact]
    public async Task Imports_Both_Files_When_Present_And_Target_Empty()
    {
        await File.WriteAllTextAsync(
            Path.Combine(_legacy, "user_preferences.json"),
            JsonSerializer.Serialize(new UserPreferences { ConnectionMethod = "cert", Role = "admin" }));
        await File.WriteAllTextAsync(
            Path.Combine(_legacy, "exo-app-params.json"),
            JsonSerializer.Serialize(new CertConfig("app-1", "tid", "org", "ABC")));

        var sut = new LegacyPreferencesImporter(_target);
        var result = await sut.TryImportAsync([_legacy]);

        result.PreferencesImported.Should().BeTrue();
        result.CertConfigImported.Should().BeTrue();
        File.Exists(Path.Combine(_target, "user_preferences.json")).Should().BeTrue();
        File.Exists(Path.Combine(_target, "exo-app-params.json")).Should().BeTrue();
    }

    [Fact]
    public async Task Does_Not_Overwrite_Existing_Target_Files()
    {
        await File.WriteAllTextAsync(
            Path.Combine(_target, "user_preferences.json"),
            JsonSerializer.Serialize(new UserPreferences { Role = "operator" }));
        await File.WriteAllTextAsync(
            Path.Combine(_legacy, "user_preferences.json"),
            JsonSerializer.Serialize(new UserPreferences { Role = "admin" }));

        var sut = new LegacyPreferencesImporter(_target);
        var result = await sut.TryImportAsync([_legacy]);

        result.PreferencesImported.Should().BeFalse();

        await using var stream = File.OpenRead(Path.Combine(_target, "user_preferences.json"));
        var loaded = await JsonSerializer.DeserializeAsync<UserPreferences>(stream);
        loaded!.Role.Should().Be("operator");
    }

    [Fact]
    public async Task Returns_False_When_Legacy_Files_Missing()
    {
        var sut = new LegacyPreferencesImporter(_target);
        var result = await sut.TryImportAsync([_legacy]);

        result.PreferencesImported.Should().BeFalse();
        result.CertConfigImported.Should().BeFalse();
    }

    public void Dispose()
    {
        try { Directory.Delete(_root, recursive: true); } catch { }
    }
}
