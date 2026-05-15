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

        _pool = RunspaceFactory.CreateRunspacePool(iss);
        _pool.SetMinRunspaces(minRunspaces);
        _pool.SetMaxRunspaces(maxRunspaces);
        _pool.ApartmentState = System.Threading.ApartmentState.MTA;
        _pool.ThreadOptions = PSThreadOptions.Default;
        _pool.Open();
    }

    public RunspacePool Pool => _pool;

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
