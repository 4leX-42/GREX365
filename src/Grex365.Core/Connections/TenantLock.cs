using Grex365.Core.Abstractions;

namespace Grex365.Core.Connections;

public sealed class TenantLock : ITenantLock
{
    private readonly IPreferencesStore _preferences;

    public TenantLock(IPreferencesStore preferences)
    {
        _preferences = preferences;
    }

    public async Task EnforceAsync(string actualTenantId, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(actualTenantId))
        {
            return;
        }

        var prefs = await _preferences.LoadAsync(cancellationToken).ConfigureAwait(false);
        if (!prefs.EnforceTenantLock || string.IsNullOrWhiteSpace(prefs.ExpectedTenantId))
        {
            return;
        }

        if (!string.Equals(prefs.ExpectedTenantId, actualTenantId, StringComparison.OrdinalIgnoreCase))
        {
            throw new TenantLockViolationException(prefs.ExpectedTenantId, actualTenantId);
        }
    }
}
