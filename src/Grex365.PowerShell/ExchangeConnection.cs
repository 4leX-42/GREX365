using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.PowerShell;

public sealed class ExchangeConnection : IExchangeConnection
{
    private const string ModuleName = "ExchangeOnlineManagement";
    private readonly IPowerShellRunner _runner;
    private bool _connected;

    public ExchangeConnection(IPowerShellRunner runner)
    {
        _runner = runner;
    }

    public bool IsConnected => _connected;

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
            Get-ConnectionInformation | Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1
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
        progress?.Report(LogEntry.Ok("EXO", "Exchange Online conectado."));
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
