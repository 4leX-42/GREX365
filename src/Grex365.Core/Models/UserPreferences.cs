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
    public string LogLevel { get; set; } = "Information";
    public string? LastSelectedNavigation { get; set; }
    public List<string> DisabledPluginAssemblies { get; set; } = new();
    public string? ApplicationInsightsConnectionString { get; set; }
    public string? AuthorizationGroupId { get; set; }
    public bool LogPanelVisible { get; set; } = false;

    // Window state restored on next launch. Null = use defaults.
    public double? WindowWidth { get; set; }
    public double? WindowHeight { get; set; }
    public double? WindowLeft { get; set; }
    public double? WindowTop { get; set; }
    public bool WindowMaximized { get; set; }

    public DateTimeOffset LastUpdated { get; set; } = DateTimeOffset.Now;
}
