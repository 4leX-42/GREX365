using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace Grex365.Core.Operations;

// A plan id proves that a dry-run of exactly this operation with exactly these arguments was
// produced recently on this installation. It is an HMAC over (operation, arguments, expiry)
// with a per-installation key, so it survives across processes (CLI dry-run, then CLI execute)
// and cannot be reused for different arguments or after it expires.
public sealed class PlanTokenService
{
    public static readonly TimeSpan Lifetime = TimeSpan.FromMinutes(15);

    private readonly byte[] _key;
    private readonly TimeProvider _time;

    public PlanTokenService(byte[] key, TimeProvider? time = null)
    {
        if (key is null || key.Length < 16)
        {
            throw new ArgumentException("La clave de planes debe tener al menos 16 bytes.", nameof(key));
        }

        _key = key;
        _time = time ?? TimeProvider.System;
    }

    // Loads the installation key, creating it on first use.
    public static PlanTokenService FromKeyFile(string path, TimeProvider? time = null)
    {
        byte[] key;
        if (File.Exists(path))
        {
            key = File.ReadAllBytes(path);
        }
        else
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            key = RandomNumberGenerator.GetBytes(32);
            File.WriteAllBytes(path, key);
        }

        return new PlanTokenService(key, time);
    }

    public string Issue(string operationId, OperationArguments arguments)
    {
        var expires = _time.GetUtcNow().Add(Lifetime).ToUnixTimeSeconds();
        return $"{expires.ToString(CultureInfo.InvariantCulture)}.{Sign(operationId, arguments, expires)}";
    }

    public bool Validate(string? planId, string operationId, OperationArguments arguments)
    {
        if (string.IsNullOrWhiteSpace(planId))
        {
            return false;
        }

        var parts = planId.Trim().Split('.');
        if (parts.Length != 2
            || !long.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out var expires)
            || expires < _time.GetUtcNow().ToUnixTimeSeconds())
        {
            return false;
        }

        var expected = Encoding.ASCII.GetBytes(Sign(operationId, arguments, expires));
        var actual = Encoding.ASCII.GetBytes(parts[1]);
        return CryptographicOperations.FixedTimeEquals(expected, actual);
    }

    private string Sign(string operationId, OperationArguments arguments, long expires)
    {
        var payload = $"{operationId.ToLowerInvariant()}\n{expires.ToString(CultureInfo.InvariantCulture)}\n{arguments.Canonical()}";
        var hash = HMACSHA256.HashData(_key, Encoding.UTF8.GetBytes(payload));
        return Convert.ToHexStringLower(hash.AsSpan(0, 16));
    }
}
