using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

public interface ITenantLock
{
    Task EnforceAsync(string actualTenantId, CancellationToken cancellationToken = default);
}

public sealed class TenantLockViolationException : InvalidOperationException
{
    public TenantLockViolationException(string expected, string actual)
        : base($"Tenant lock violado · esperado={expected} · actual={actual}")
    {
        Expected = expected;
        Actual = actual;
    }

    public string Expected { get; }
    public string Actual { get; }
}
