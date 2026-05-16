namespace Grex365.Core.Models;

public sealed class UserPreferences
{
    public string? ConnectionMethod { get; set; }
    public string? TraditionalAdminUpn { get; set; }
    public string? Organization { get; set; }
    public bool FirstRunCompleted { get; set; }
    public string? ExpectedTenantId { get; set; }
    public string? ExpectedTenantDomain { get; set; }
    public bool EnforceTenantLock { get; set; }
    public string Role { get; set; } = "operator";
    public string UIMode { get; set; } = "support";
    public string Theme { get; set; } = "Dark";
    public DateTimeOffset LastUpdated { get; set; } = DateTimeOffset.Now;
}
