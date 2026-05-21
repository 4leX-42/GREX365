using System.Management.Automation.Runspaces;

namespace Grex365.PowerShell;

public sealed class RunspacePoolHost : IDisposable
{
    private readonly RunspacePool _pool;
    private bool _disposed;

    public RunspacePoolHost(int minRunspaces = 1, int maxRunspaces = 4)
    {
        var iss = InitialSessionState.CreateDefault2();
        iss.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.Bypass;
        iss.Variables.Add(new SessionStateVariableEntry("ConfirmPreference", "None", string.Empty));
        iss.Variables.Add(new SessionStateVariableEntry("ProgressPreference", "SilentlyContinue", string.Empty));
        iss.Variables.Add(new SessionStateVariableEntry("WarningPreference", "SilentlyContinue", string.Empty));
        iss.Variables.Add(new SessionStateVariableEntry("ErrorActionPreference", "Continue", string.Empty));

        // Filter PSModulePath: drop WindowsApps store paths because their ACLs deny load of
        // Microsoft.PackageManagement.dll from embedded runspaces. Keep CurrentUser + ProgramFiles + System32.
        var sanitizedModulePath = SanitizeModulePath(Environment.GetEnvironmentVariable("PSModulePath") ?? string.Empty);
        iss.Variables.Add(new SessionStateVariableEntry("env:PSModulePath", sanitizedModulePath, string.Empty));
        iss.EnvironmentVariables.Add(new SessionStateVariableEntry("PSModulePath", sanitizedModulePath, string.Empty));

        _pool = RunspaceFactory.CreateRunspacePool(iss);
        _pool.SetMinRunspaces(minRunspaces);
        _pool.SetMaxRunspaces(maxRunspaces);
        _pool.ApartmentState = System.Threading.ApartmentState.MTA;
        _pool.ThreadOptions = PSThreadOptions.Default;
        _pool.Open();
    }

    public RunspacePool Pool => _pool;

    private static string SanitizeModulePath(string original)
    {
        if (string.IsNullOrEmpty(original))
        {
            return BuildSafeDefault();
        }
        var parts = original.Split(System.IO.Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries);
        var kept = new List<string>(parts.Length);
        foreach (var p in parts)
        {
            if (p.IndexOf("WindowsApps", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                continue;
            }
            kept.Add(p);
        }
        // Ensure CurrentUser modules path is included.
        var userModules = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "PowerShell", "Modules");
        if (!kept.Any(p => string.Equals(p, userModules, StringComparison.OrdinalIgnoreCase)))
        {
            kept.Insert(0, userModules);
        }
        return string.Join(System.IO.Path.PathSeparator, kept);
    }

    private static string BuildSafeDefault()
    {
        var docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var user = System.IO.Path.Combine(docs, "PowerShell", "Modules");
        var prog = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "PowerShell", "Modules");
        var sys = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "WindowsPowerShell", "v1.0", "Modules");
        return string.Join(System.IO.Path.PathSeparator, new[] { user, prog, sys });
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }
        _disposed = true;
        try
        {
            _pool.Close();
        }
        catch
        {
            // ignore
        }
        _pool.Dispose();
    }
}
