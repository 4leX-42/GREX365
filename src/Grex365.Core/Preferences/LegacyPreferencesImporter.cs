using System.Text.Json;
using Grex365.Core.Models;

namespace Grex365.Core.Preferences;

/// Imports user_preferences.json and exo-app-params.json from a legacy
/// GREX365 PowerShell toolkit checkout when found next to the executable.
/// Safe to call repeatedly: only copies when target file does not exist.
public sealed class LegacyPreferencesImporter
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = null
    };

    private readonly string _targetConfigDir;

    public LegacyPreferencesImporter(string targetConfigDir)
    {
        _targetConfigDir = targetConfigDir;
    }

    public async Task<ImportResult> TryImportAsync(IEnumerable<string> candidateLegacyDirs, CancellationToken ct = default)
    {
        Directory.CreateDirectory(_targetConfigDir);

        var importedPrefs = false;
        var importedCert = false;

        foreach (var dir in candidateLegacyDirs)
        {
            var legacyPrefs = Path.Combine(dir, "user_preferences.json");
            var targetPrefs = Path.Combine(_targetConfigDir, "user_preferences.json");
            if (!importedPrefs && File.Exists(legacyPrefs) && !File.Exists(targetPrefs))
            {
                try
                {
                    await using var src = File.OpenRead(legacyPrefs);
                    var parsed = await JsonSerializer.DeserializeAsync<UserPreferences>(src, JsonOptions, ct).ConfigureAwait(false);
                    if (parsed is not null)
                    {
                        await using var dst = File.Create(targetPrefs);
                        await JsonSerializer.SerializeAsync(dst, parsed, JsonOptions, ct).ConfigureAwait(false);
                        importedPrefs = true;
                    }
                }
                catch
                {
                    // ignore corrupt legacy file
                }
            }

            var legacyCert = Path.Combine(dir, "exo-app-params.json");
            var targetCert = Path.Combine(_targetConfigDir, "exo-app-params.json");
            if (!importedCert && File.Exists(legacyCert) && !File.Exists(targetCert))
            {
                try
                {
                    await using var src = File.OpenRead(legacyCert);
                    var parsed = await JsonSerializer.DeserializeAsync<CertConfig>(src, JsonOptions, ct).ConfigureAwait(false);
                    if (parsed is not null)
                    {
                        await using var dst = File.Create(targetCert);
                        await JsonSerializer.SerializeAsync(dst, parsed, JsonOptions, ct).ConfigureAwait(false);
                        importedCert = true;
                    }
                }
                catch
                {
                    // ignore
                }
            }

            if (importedPrefs && importedCert)
            {
                break;
            }
        }

        return new ImportResult(importedPrefs, importedCert);
    }

    public sealed record ImportResult(bool PreferencesImported, bool CertConfigImported);
}
