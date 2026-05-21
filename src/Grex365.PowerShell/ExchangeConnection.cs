using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.PowerShell;

public sealed class ExchangeConnection : IExchangeConnection
{
    private const string ModuleName = "ExchangeOnlineManagement";

    private readonly IPowerShellRunner _runner;
    private bool _connected;
    private string? _tenantId;
    private string? _organization;

    public ExchangeConnection(IPowerShellRunner runner)
    {
        _runner = runner;
    }

    public bool IsConnected => _connected;

    public string? TenantId => _tenantId;

    public string? Organization => _organization;

    public async Task ConnectByCertificateAsync(
        CertConfig config,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        progress?.Report(LogEntry.Info("EXO", "Asegurando módulo ExchangeOnlineManagement..."));
        await EnsureModuleAsync(progress, cancellationToken).ConfigureAwait(false);

        progress?.Report(LogEntry.Info("EXO", $"Connect-ExchangeOnline (cert) tenant={config.TenantId}"));
        const string script = """
            param([string]$AppId, [string]$Thumbprint, [string]$Organization)
            Import-Module ExchangeOnlineManagement -ErrorAction Stop
            Connect-ExchangeOnline `
                -AppId $AppId `
                -CertificateThumbprint $Thumbprint `
                -Organization $Organization `
                -ShowBanner:$false `
                -ErrorAction Stop
            $info = Get-ConnectionInformation | Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1
            if ($info) {
                [PSCustomObject]@{
                    TenantId     = [string]$info.TenantId
                    Organization = [string]$info.Organization
                }
            }
            """;

        var result = await _runner.RunAsync(
            script,
            new Dictionary<string, object?>
            {
                ["AppId"] = config.AppId,
                ["Thumbprint"] = config.CertThumbprint,
                ["Organization"] = config.Organization
            },
            progress,
            cancellationToken).ConfigureAwait(false);

        if (!result.Success)
        {
            throw new InvalidOperationException(
                "Connect-ExchangeOnline falló: " + string.Join("; ", result.Errors));
        }

        _connected = true;
        _tenantId = config.TenantId;
        _organization = config.Organization;

        progress?.Report(LogEntry.Ok("EXO", "Exchange Online conectado."));
    }

    public Task<bool> CheckLiveAsync(CancellationToken cancellationToken = default)
    {
        // EXO sesion vive en el runspace donde Connect-ExchangeOnline corrio.
        // Probar Get-ConnectionInformation desde un runspace nuevo del pool flippea a false
        // (la sesion no es compartida). Confiamos en el flag _connected hasta Disconnect explicito.
        return Task.FromResult(_connected);
    }

    public async Task<ExoModuleStatus> ProbeModuleAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        const string script = """
            param([string]$Name)
            $loaded = Get-Module -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
            if ($loaded) {
                [PSCustomObject]@{ Installed = $true; Version = [string]$loaded.Version; Detail = 'loaded' }
                return
            }
            $available = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
            if ($available) {
                [PSCustomObject]@{ Installed = $true; Version = [string]$available.Version; Detail = 'available' }
            } else {
                [PSCustomObject]@{ Installed = $false; Version = $null; Detail = 'not installed' }
            }
            """;

        var result = await _runner.RunAsync(
            script,
            new Dictionary<string, object?> { ["Name"] = ModuleName },
            progress,
            cancellationToken).ConfigureAwait(false);

        if (!result.Success || result.Output.Count == 0)
        {
            return new ExoModuleStatus(false, null, "probe failed: " + string.Join("; ", result.Errors));
        }

        if (result.Output[0] is System.Management.Automation.PSObject ps)
        {
            var installed = ps.Properties["Installed"]?.Value is bool b && b;
            var version = ps.Properties["Version"]?.Value?.ToString();
            var detail = ps.Properties["Detail"]?.Value?.ToString();
            return new ExoModuleStatus(installed, version, detail);
        }

        return new ExoModuleStatus(false, null, "unknown probe output");
    }

    public async Task<ExoModuleStatus> InstallModuleAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        progress?.Report(LogEntry.Info("EXO", $"Instalando {ModuleName} en CurrentUser (proceso externo)..."));

        // Run Install-Module in a SEPARATE pwsh.exe process to bypass the WindowsApps
        // PackageManagement.dll access-denied issue that hits embedded runspaces.
        var psi = new System.Diagnostics.ProcessStartInfo
        {
            FileName = ResolvePwshExe(),
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
        };
        psi.ArgumentList.Add("-NoLogo");
        psi.ArgumentList.Add("-NoProfile");
        psi.ArgumentList.Add("-NonInteractive");
        psi.ArgumentList.Add("-Command");
        psi.ArgumentList.Add(
            $"try {{ " +
            $"  if (-not (Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue) -or (Get-PSRepository -Name PSGallery).InstallationPolicy -ne 'Trusted') {{ " +
            $"    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue " +
            $"  }}; " +
            $"  Install-Module -Name {ModuleName} -Scope CurrentUser -Force -AllowClobber -Confirm:$false -ErrorAction Stop; " +
            $"  Write-Host 'INSTALLED' " +
            $"}} catch {{ Write-Host 'ERROR:' $_.Exception.Message; exit 1 }}");

        using var proc = new System.Diagnostics.Process { StartInfo = psi };
        proc.OutputDataReceived += (_, e) =>
        {
            if (!string.IsNullOrEmpty(e.Data))
            {
                progress?.Report(LogEntry.Info("EXO-Install", e.Data));
            }
        };
        proc.ErrorDataReceived += (_, e) =>
        {
            if (!string.IsNullOrEmpty(e.Data))
            {
                progress?.Report(LogEntry.Warn("EXO-Install", e.Data));
            }
        };

        proc.Start();
        proc.BeginOutputReadLine();
        proc.BeginErrorReadLine();

        await using (cancellationToken.Register(() =>
        {
            try { if (!proc.HasExited) proc.Kill(entireProcessTree: true); } catch { }
        }).ConfigureAwait(false))
        {
            await proc.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
        }

        if (proc.ExitCode != 0)
        {
            return new ExoModuleStatus(false, null, $"pwsh exit code {proc.ExitCode}");
        }

        progress?.Report(LogEntry.Ok("EXO-Install", $"{ModuleName} instalado. Reprobando..."));
        return await ProbeModuleAsync(progress, cancellationToken).ConfigureAwait(false);
    }

    private static string ResolvePwshExe()
    {
        // Prefer pwsh.exe on PATH; fallback to legacy powershell.exe.
        foreach (var name in new[] { "pwsh.exe", "powershell.exe" })
        {
            var path = Environment.GetEnvironmentVariable("PATH")?
                .Split(Path.PathSeparator)
                .Select(p => Path.Combine(p, name))
                .FirstOrDefault(File.Exists);
            if (!string.IsNullOrEmpty(path)) return path;
        }
        // Last resort: rely on shell resolution
        return "powershell.exe";
    }

    public async Task DisconnectAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        if (!_connected)
        {
            return;
        }

        const string script = """
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            """;

        await _runner.RunAsync(script, parameters: null, progress, cancellationToken).ConfigureAwait(false);
        _connected = false;
        _tenantId = null;
        _organization = null;
        progress?.Report(LogEntry.Info("EXO", "Exchange Online desconectado."));
    }

    private async Task EnsureModuleAsync(IProgress<LogEntry>? progress, CancellationToken ct)
    {
        // Probe-only: NO Install-Module desde la app (PackageManagement bajo WindowsApps
        // tiene ACLs restrictivas y rompe en hosts embebidos). Si falta, instruir al user.
        const string probeScript = """
            param([string]$Name)
            $loaded = Get-Module -Name $Name
            if ($loaded) {
                [PSCustomObject]@{ Loaded = $true; Available = $true; Version = [string]$loaded.Version }
                return
            }
            $available = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
            if (-not $available) {
                [PSCustomObject]@{ Loaded = $false; Available = $false; Version = $null }
                return
            }
            Import-Module $Name -ErrorAction Stop -Verbose:$false
            [PSCustomObject]@{ Loaded = $true; Available = $true; Version = [string]$available.Version }
            """;

        var result = await _runner.RunAsync(
            probeScript,
            new Dictionary<string, object?> { ["Name"] = ModuleName },
            progress,
            ct).ConfigureAwait(false);

        if (!result.Success)
        {
            throw new InvalidOperationException(
                $"No se pudo importar {ModuleName}: " + string.Join("; ", result.Errors));
        }

        var available = false;
        if (result.Output.Count > 0 && result.Output[0] is System.Management.Automation.PSObject ps)
        {
            available = ps.Properties["Available"]?.Value is bool b && b;
        }

        if (!available)
        {
            throw new InvalidOperationException(
                $"Modulo {ModuleName} no esta instalado. Abre PowerShell (no admin) y ejecuta: " +
                $"Install-Module {ModuleName} -Scope CurrentUser -Force; despues vuelve a Conectar.");
        }
    }
}
