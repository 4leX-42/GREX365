using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.PowerShell;

public sealed class SharedMailboxService : ISharedMailboxService
{
    private readonly IPowerShellRunner _runner;

    public SharedMailboxService(IPowerShellRunner runner)
    {
        _runner = runner;
    }

    public async Task<MailboxInfo?> GetMailboxAsync(string identity, IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default)
    {
        const string script = """
            $m = Get-Mailbox -Identity $Identity -ErrorAction Stop
            [PSCustomObject]@{
                Identity            = [string]$m.Identity
                DisplayName         = [string]$m.DisplayName
                PrimarySmtpAddress  = [string]$m.PrimarySmtpAddress
                RecipientTypeDetails = [string]$m.RecipientTypeDetails
            }
            """;

        var result = await _runner.RunAsync(
            script,
            new Dictionary<string, object?> { ["Identity"] = identity },
            progress,
            cancellationToken).ConfigureAwait(false);

        if (!result.Success || result.Output.Count == 0)
        {
            return null;
        }

        return Map(result.Output[0]);
    }

    public async Task<MailboxInfo?> ConvertToSharedAsync(string identity, IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default)
    {
        const string script = """
            $current = Get-Mailbox -Identity $Identity -ErrorAction Stop
            if ($current.RecipientTypeDetails -eq 'SharedMailbox') {
                [PSCustomObject]@{
                    Identity = [string]$current.Identity
                    DisplayName = [string]$current.DisplayName
                    PrimarySmtpAddress = [string]$current.PrimarySmtpAddress
                    RecipientTypeDetails = [string]$current.RecipientTypeDetails
                }
                return
            }

            Set-Mailbox -Identity $Identity -Type Shared -ErrorAction Stop
            Write-Information "Set-Mailbox -Type Shared aplicado."

            $deadline = (Get-Date).AddSeconds(120)
            do {
                Start-Sleep -Seconds 5
                try {
                    $check = Get-Mailbox -Identity $Identity -ErrorAction Stop
                } catch {
                    $check = $null
                }
                if ($check -and $check.RecipientTypeDetails -eq 'SharedMailbox') {
                    break
                }
            } while ((Get-Date) -lt $deadline)

            if (-not $check) {
                throw "No se pudo confirmar la conversion: Get-Mailbox devolvio vacio."
            }

            [PSCustomObject]@{
                Identity            = [string]$check.Identity
                DisplayName         = [string]$check.DisplayName
                PrimarySmtpAddress  = [string]$check.PrimarySmtpAddress
                RecipientTypeDetails = [string]$check.RecipientTypeDetails
            }
            """;

        var result = await _runner.RunAsync(
            script,
            new Dictionary<string, object?> { ["Identity"] = identity },
            progress,
            cancellationToken).ConfigureAwait(false);

        if (!result.Success)
        {
            throw new InvalidOperationException("Conversion a Shared fallo: " + string.Join("; ", result.Errors));
        }

        return result.Output.Count > 0 ? Map(result.Output[0]) : null;
    }

    public async Task<MailboxInfo?> ConvertToRegularAsync(string identity, IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default)
    {
        const string script = """
            $current = Get-Mailbox -Identity $Identity -ErrorAction Stop
            if ($current.RecipientTypeDetails -ne 'SharedMailbox') {
                [PSCustomObject]@{
                    Identity = [string]$current.Identity
                    DisplayName = [string]$current.DisplayName
                    PrimarySmtpAddress = [string]$current.PrimarySmtpAddress
                    RecipientTypeDetails = [string]$current.RecipientTypeDetails
                }
                return
            }

            Set-Mailbox -Identity $Identity -Type Regular -ErrorAction Stop
            Write-Information "Set-Mailbox aplicado. Esperando propagación..."

            $deadline = (Get-Date).AddSeconds(120)
            do {
                Start-Sleep -Seconds 5
                try {
                    $check = Get-Mailbox -Identity $Identity -ErrorAction Stop
                } catch {
                    $check = $null
                }
                if ($check -and $check.RecipientTypeDetails -ne 'SharedMailbox') {
                    break
                }
            } while ((Get-Date) -lt $deadline)

            if (-not $check) {
                throw "No se pudo confirmar la conversión: Get-Mailbox devolvió vacío."
            }

            [PSCustomObject]@{
                Identity            = [string]$check.Identity
                DisplayName         = [string]$check.DisplayName
                PrimarySmtpAddress  = [string]$check.PrimarySmtpAddress
                RecipientTypeDetails = [string]$check.RecipientTypeDetails
            }
            """;

        var result = await _runner.RunAsync(
            script,
            new Dictionary<string, object?> { ["Identity"] = identity },
            progress,
            cancellationToken).ConfigureAwait(false);

        if (!result.Success)
        {
            throw new InvalidOperationException("Conversión falló: " + string.Join("; ", result.Errors));
        }

        return result.Output.Count > 0 ? Map(result.Output[0]) : null;
    }

    public async Task<MailboxPermissionResult> ApplyPermissionAsync(
        string action,
        string permission,
        string mailbox,
        string principal,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var a = (action ?? string.Empty).Trim();
        var p = (permission ?? string.Empty).Trim();
        var m = mailbox ?? string.Empty;
        var pr = principal ?? string.Empty;
        var actionLower = a.ToLowerInvariant();

        if (actionLower != "add" && actionLower != "remove")
        {
            return new MailboxPermissionResult(a, p, m, pr, "INVALIDO", "Action debe ser add|remove");
        }
        if (p is not ("FullAccess" or "SendAs" or "SendOnBehalf"))
        {
            return new MailboxPermissionResult(a, p, m, pr, "INVALIDO", "Permission no soportada");
        }
        if (string.IsNullOrWhiteSpace(m) || string.IsNullOrWhiteSpace(pr))
        {
            return new MailboxPermissionResult(a, p, m, pr, "INVALIDO", "Mailbox/Principal vacío");
        }

        var script = BuildPermissionScript(actionLower, p);

        var result = await _runner.RunAsync(
            script,
            new Dictionary<string, object?>
            {
                ["Mailbox"] = m,
                ["Principal"] = pr
            },
            progress,
            cancellationToken).ConfigureAwait(false);

        if (!result.Success)
        {
            return new MailboxPermissionResult(a, p, m, pr, "ERROR", string.Join("; ", result.Errors));
        }

        return new MailboxPermissionResult(a, p, m, pr, "OK", "Aplicado");
    }

    private static string BuildPermissionScript(string action, string permission)
    {
        var add = action == "add";

        return permission switch
        {
            "FullAccess" => add
                ? "Add-MailboxPermission -Identity $Mailbox -User $Principal -AccessRights FullAccess -InheritanceType All -AutoMapping:$true -Confirm:$false -ErrorAction Stop | Out-Null"
                : "Remove-MailboxPermission -Identity $Mailbox -User $Principal -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop | Out-Null",

            "SendAs" => add
                ? "Add-RecipientPermission -Identity $Mailbox -Trustee $Principal -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null"
                : "Remove-RecipientPermission -Identity $Mailbox -Trustee $Principal -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null",

            "SendOnBehalf" => add
                ? "Set-Mailbox -Identity $Mailbox -GrantSendOnBehalfTo @{Add=$Principal} -ErrorAction Stop | Out-Null"
                : "Set-Mailbox -Identity $Mailbox -GrantSendOnBehalfTo @{Remove=$Principal} -ErrorAction Stop | Out-Null",

            _ => throw new InvalidOperationException("Permission no soportada: " + permission)
        };
    }

    public async Task<IReadOnlyList<MailboxPermissionEntry>> GetPermissionsAsync(
        string mailbox,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        const string script = """
            $full = @(Get-MailboxPermission -Identity $Mailbox -ErrorAction Stop |
                Where-Object { $_.AccessRights -contains 'FullAccess' -and -not $_.IsInherited -and $_.User -notlike 'NT AUTHORITY\\SELF' })
            $send = @(Get-RecipientPermission -Identity $Mailbox -ErrorAction SilentlyContinue |
                Where-Object { $_.AccessRights -contains 'SendAs' -and $_.Trustee -notlike 'NT AUTHORITY\\SELF' })
            $mbx  = Get-Mailbox -Identity $Mailbox -ErrorAction Stop
            $onBehalf = @($mbx.GrantSendOnBehalfTo)

            $out = New-Object System.Collections.Generic.List[object]
            foreach ($f in $full)  { $out.Add([PSCustomObject]@{ Permission='FullAccess';   Principal=[string]$f.User;    Detail=[string]$f.AccessRights }) }
            foreach ($s in $send)  { $out.Add([PSCustomObject]@{ Permission='SendAs';       Principal=[string]$s.Trustee; Detail=[string]$s.AccessRights }) }
            foreach ($o in $onBehalf) { $out.Add([PSCustomObject]@{ Permission='SendOnBehalf'; Principal=[string]$o; Detail='From Set-Mailbox' }) }
            $out
            """;

        var result = await _runner.RunAsync(
            script,
            new Dictionary<string, object?> { ["Mailbox"] = mailbox },
            progress,
            cancellationToken).ConfigureAwait(false);

        if (!result.Success)
        {
            throw new InvalidOperationException("Get-permissions falló: " + string.Join("; ", result.Errors));
        }

        var entries = new List<MailboxPermissionEntry>(result.Output.Count);
        foreach (var raw in result.Output)
        {
            entries.Add(MapPermission(raw));
        }
        return entries;
    }

    private static MailboxPermissionEntry MapPermission(object? raw)
    {
        if (raw is System.Management.Automation.PSObject ps)
        {
            return new MailboxPermissionEntry(
                Permission: ps.Properties["Permission"]?.Value?.ToString() ?? string.Empty,
                Principal: ps.Properties["Principal"]?.Value?.ToString() ?? string.Empty,
                Detail: ps.Properties["Detail"]?.Value?.ToString() ?? string.Empty);
        }
        var t = raw?.GetType();
        if (t is null)
        {
            return new MailboxPermissionEntry(string.Empty, string.Empty, string.Empty);
        }
        string Get(string name) => t.GetProperty(name)?.GetValue(raw)?.ToString() ?? string.Empty;
        return new MailboxPermissionEntry(Get("Permission"), Get("Principal"), Get("Detail"));
    }

    private static MailboxInfo Map(object? raw)
    {
        if (raw is System.Management.Automation.PSObject ps)
        {
            return new MailboxInfo(
                Identity: ps.Properties["Identity"]?.Value?.ToString() ?? string.Empty,
                DisplayName: ps.Properties["DisplayName"]?.Value?.ToString() ?? string.Empty,
                PrimarySmtpAddress: ps.Properties["PrimarySmtpAddress"]?.Value?.ToString() ?? string.Empty,
                RecipientTypeDetails: ps.Properties["RecipientTypeDetails"]?.Value?.ToString() ?? string.Empty);
        }

        var t = raw?.GetType();
        if (t is null)
        {
            return new MailboxInfo(string.Empty, string.Empty, string.Empty, string.Empty);
        }

        string Get(string name) => t.GetProperty(name)?.GetValue(raw)?.ToString() ?? string.Empty;
        return new MailboxInfo(Get("Identity"), Get("DisplayName"), Get("PrimarySmtpAddress"), Get("RecipientTypeDetails"));
    }
}
