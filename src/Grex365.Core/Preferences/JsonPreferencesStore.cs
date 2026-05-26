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

        try
        {
            await using var stream = File.OpenRead(_filePath);
            var prefs = await JsonSerializer.DeserializeAsync<UserPreferences>(stream, JsonOptions, cancellationToken).ConfigureAwait(false);
            return prefs ?? new UserPreferences();
        }
        catch (JsonException)
        {
            // Corrupt JSON (manual edit / disk truncation / version drift): quarantine the
            // bad file so the user can recover values manually, then start fresh with defaults
            // rather than throwing on app startup.
            CorruptFileQuarantine.MoveAside(_filePath);
            return new UserPreferences();
        }
        catch (IOException)
        {
            return new UserPreferences();
        }
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

        try
        {
            await using var stream = File.OpenRead(_filePath);
            return await JsonSerializer.DeserializeAsync<CertConfig>(stream, JsonOptions, cancellationToken).ConfigureAwait(false);
        }
        catch (JsonException)
        {
            CorruptFileQuarantine.MoveAside(_filePath);
            return null;
        }
        catch (IOException)
        {
            return null;
        }
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

internal static class CorruptFileQuarantine
{
    // Renames a corrupt config file to "<name>.corrupted-yyyyMMddHHmmss.bak" so the
    // user can still inspect/recover values manually. Best-effort: silently swallows
    // any I/O failure (target locked, perm denied) — the caller has already decided
    // to proceed with defaults.
    public static string? MoveAside(string filePath)
    {
        try
        {
            var stamp = DateTime.UtcNow.ToString("yyyyMMddHHmmss");
            var backup = filePath + $".corrupted-{stamp}.bak";
            File.Move(filePath, backup, overwrite: true);
            return backup;
        }
        catch (IOException)
        {
            return null;
        }
        catch (UnauthorizedAccessException)
        {
            return null;
        }
    }
}
