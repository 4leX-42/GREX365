using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.PowerShell;

public sealed class ExchangeConnection : IExchangeConnection
{
    private const string ModuleName = "ExchangeOnlineManagement";
    private static readonly TimeSpan ProbeCacheTtl = TimeSpan.FromSeconds(10);

    private readonly IPowerShellRunner _runner;
    private bool _connected;
    private string? _tenantId;
    private string? _organization;
    private DateTimeOffset _lastProbe = DateTimeOffset.MinValue;
    private bool _lastProbeResult;

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
        _lastProbeResult = true;
        _lastProbe = DateTimeOffset.Now;

        progress?.Report(LogEntry.Ok("EXO", "Exchange Online conectado."));
    }

    public async Task<bool> CheckLiveAsync(CancellationToken cancellationToken = default)
    {
        if (!_connected)
        {
            return false;
        }

        var now = DateTimeOffset.Now;
        if (now - _lastProbe < ProbeCacheTtl)
        {
            return _lastProbeResult;
        }

        try
        {
            const string script = """
                $info = Get-ConnectionInformation -ErrorAction SilentlyContinue |
                    Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1
                [bool]$info
                """;
            var result = await _runner.RunAsync(script, parameters: null, progress: null, cancellationToken).ConfigureAwait(false);
            _lastProbeResult = result.Success
                && result.Output.Count > 0
                && result.Output[0] is bool b
                && b;
        }
        catch
        {
            _lastProbeResult = false;
        }

        _lastProbe = now;
        return _lastProbeResult;
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
        _lastProbeResult = false;
        _lastProbe = DateTimeOffset.MinValue;
        progress?.Report(LogEntry.Info("EXO", "Exchange Online desconectado."));
    }

    private async Task EnsureModuleAsync(IProgress<LogEntry>? progress, CancellationToken ct)
    {
        const string script = """
            if (-not (Get-Module -Name $Name)) {
                if (-not (Get-Module -ListAvailable -Name $Name)) {
                    try {
                        $repo = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
                        if ($repo -and $repo.InstallationPolicy -ne 'Trusted') {
                            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
                        }
                    } catch {}
                    Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -Confirm:$false -ErrorAction Stop
                }
                Import-Module $Name -ErrorAction Stop -Verbose:$false
            }
            """;

        var result = await _runner.RunAsync(
            script,
            new Dictionary<string, object?> { ["Name"] = ModuleName },
            progress,
            ct).ConfigureAwait(false);

        if (!result.Success)
        {
            throw new InvalidOperationException(
                $"No se pudo cargar el módulo {ModuleName}: " + string.Join("; ", result.Errors));
        }
    }
}
