namespace Grex365.PowerShell;

public static class PSModulePathSanitizer
{
    // Drops WindowsApps store paths because their ACLs deny load of
    // Microsoft.PackageManagement.dll from embedded runspaces. Keeps CurrentUser
    // + ProgramFiles + System32 modules. Always ensures the CurrentUser modules
    // path is present (prepended if missing).
    public static string Sanitize(string original) => Sanitize(original, ResolveUserModulesPath());

    // Test-friendly overload: caller supplies the user modules path (avoids Environment dep).
    public static string Sanitize(string original, string userModulesPath)
    {
        if (string.IsNullOrEmpty(original))
        {
            return BuildSafeDefault(userModulesPath);
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
        if (!kept.Any(p => string.Equals(p, userModulesPath, StringComparison.OrdinalIgnoreCase)))
        {
            kept.Insert(0, userModulesPath);
        }
        return string.Join(System.IO.Path.PathSeparator, kept);
    }

    public static string BuildSafeDefault() => BuildSafeDefault(ResolveUserModulesPath());

    public static string BuildSafeDefault(string userModulesPath)
    {
        var prog = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
            "PowerShell", "Modules");
        var sys = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.System),
            "WindowsPowerShell", "v1.0", "Modules");
        return string.Join(System.IO.Path.PathSeparator, new[] { userModulesPath, prog, sys });
    }

    public static string ResolveUserModulesPath()
        => System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "PowerShell", "Modules");
}
