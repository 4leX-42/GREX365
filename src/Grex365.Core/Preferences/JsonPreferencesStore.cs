using System.Text.Json;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.Core.Preferences;

public sealed class JsonPreferencesStore : IPreferencesStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = null
    };

    private readonly string _filePath;

    public JsonPreferencesStore(string configDirectory)
    {
        Directory.CreateDirectory(configDirectory);
        _filePath = Path.Combine(configDirectory, "user_preferences.json");
    }

    public async Task<UserPreferences> LoadAsync(CancellationToken cancellationToken = default)
    {
        if (!File.Exists(_filePath))
        {
            return new UserPreferences();
        }

        await using var stream = File.OpenRead(_filePath);
        var prefs = await JsonSerializer.DeserializeAsync<UserPreferences>(stream, JsonOptions, cancellationToken).ConfigureAwait(false);
        return prefs ?? new UserPreferences();
    }

    public async Task SaveAsync(UserPreferences preferences, CancellationToken cancellationToken = default)
    {
        preferences.LastUpdated = DateTimeOffset.Now;
        await using var stream = File.Create(_filePath);
        await JsonSerializer.SerializeAsync(stream, preferences, JsonOptions, cancellationToken).ConfigureAwait(false);
    }
}

public sealed class JsonCertConfigStore : ICertConfigStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = null
    };

    private readonly string _filePath;

    public JsonCertConfigStore(string configDirectory)
    {
        Directory.CreateDirectory(configDirectory);
        _filePath = Path.Combine(configDirectory, "exo-app-params.json");
    }

    public async Task<CertConfig?> LoadAsync(CancellationToken cancellationToken = default)
    {
        if (!File.Exists(_filePath))
        {
            return null;
        }

        await using var stream = File.OpenRead(_filePath);
        return await JsonSerializer.DeserializeAsync<CertConfig>(stream, JsonOptions, cancellationToken).ConfigureAwait(false);
    }

    public async Task SaveAsync(CertConfig config, CancellationToken cancellationToken = default)
    {
        await using var stream = File.Create(_filePath);
        await JsonSerializer.SerializeAsync(stream, config, JsonOptions, cancellationToken).ConfigureAwait(false);
    }

    public Task DeleteAsync(CancellationToken cancellationToken = default)
    {
        if (File.Exists(_filePath))
        {
            File.Delete(_filePath);
        }
        return Task.CompletedTask;
    }
}
