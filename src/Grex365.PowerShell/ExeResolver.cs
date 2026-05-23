namespace Grex365.PowerShell;

public static class ExeResolver
{
    // Walks the PATH environment looking for the first candidate exe that exists.
    // Returns fallback when nothing matches (caller can choose: bare name for shell
    // resolution at runtime, or an explicit absolute fallback).
    public static string Resolve(
        IEnumerable<string> candidateNames,
        string? pathEnv,
        Func<string, bool> fileExists,
        string fallback)
    {
        ArgumentNullException.ThrowIfNull(candidateNames);
        ArgumentNullException.ThrowIfNull(fileExists);
        ArgumentNullException.ThrowIfNull(fallback);

        if (string.IsNullOrEmpty(pathEnv))
        {
            return fallback;
        }

        var segments = pathEnv.Split(System.IO.Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries);
        foreach (var name in candidateNames)
        {
            foreach (var seg in segments)
            {
                var candidate = System.IO.Path.Combine(seg, name);
                if (fileExists(candidate))
                {
                    return candidate;
                }
            }
        }
        return fallback;
    }

    // Default impl using process environment + File.Exists.
    public static string ResolvePwsh()
        => Resolve(
            candidateNames: ["pwsh.exe", "powershell.exe"],
            pathEnv: Environment.GetEnvironmentVariable("PATH"),
            fileExists: System.IO.File.Exists,
            fallback: "powershell.exe");
}
