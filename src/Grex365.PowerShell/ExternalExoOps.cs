using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.Json;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.PowerShell;

public sealed class ExternalExoOps : IExternalExoOps
{
    private const string JsonMarker = "###GREX-JSON###";
    private const string ErrMarker = "###GREX-ERR###";
    private readonly ICertConfigStore _certStore;

    public ExternalExoOps(ICertConfigStore certStore)
    {
        _certStore = certStore;
    }

    public async Task<MailboxInfo?> GetMailboxFactsAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var cfg = await RequireConfigAsync(cancellationToken).ConfigureAwait(false);
        var id = Lit(identity);
        var body = $$"""
            $id = {{id}}
            $m = Get-Mailbox -Identity $id -ErrorAction Stop
            $bytes = $null
            try {
                $s = Get-EXOMailboxStatistics -Identity $id -Properties TotalItemSize -ErrorAction Stop
                $t = [string]$s.TotalItemSize
                if ($t -match '\(([\d,]+) bytes\)') { $bytes = [int64]($matches[1] -replace ',','') }
            } catch { }
            $holds = 0; if ($m.InPlaceHolds) { $holds = @($m.InPlaceHolds).Count }
            $arch = $false
            if ($m.ArchiveStatus -and [string]$m.ArchiveStatus -ne 'None') { $arch = $true }
            $o = [PSCustomObject]@{
                Identity             = [string]$m.Identity
                DisplayName          = [string]$m.DisplayName
                PrimarySmtpAddress   = [string]$m.PrimarySmtpAddress
                RecipientTypeDetails = [string]$m.RecipientTypeDetails
                LitigationHoldEnabled= [bool]$m.LitigationHoldEnabled
                InPlaceHoldCount     = [int]$holds
                ArchiveEnabled       = [bool]$arch
                TotalItemBytes       = $bytes
            }
            Write-Output ('{{JsonMarker}}' + ($o | ConvertTo-Json -Compress))
            """;

        try
        {
            var json = await RunAsync(cfg, body, progress, cancellationToken).ConfigureAwait(false);
            return Parse(json);
        }
        catch (Exception ex) when (IsMailboxNotFound(ex.Message))
        {
            // Mailbox doesn't exist — a normal "not found", not a failure. Return null so the
            // Shared Mailbox lookup / offboarding pre-checks report it gracefully (no red error).
            progress?.Report(LogEntry.Info("EXO", $"Buzón no encontrado: {identity}"));
            return null;
        }
    }

    // EXO "the object couldn't be found", across locales. The accented text can arrive mojibake'd,
    // so match on the stable fragment. Public for unit testing.
    public static bool IsMailboxNotFound(string? message)
    {
        if (string.IsNullOrEmpty(message)) return false;
        return message.Contains("no se encontr", StringComparison.OrdinalIgnoreCase)
            || message.Contains("couldn't be found", StringComparison.OrdinalIgnoreCase)
            || message.Contains("wasn't found", StringComparison.OrdinalIgnoreCase)
            || message.Contains("can't be found", StringComparison.OrdinalIgnoreCase)
            || message.Contains("ManagementObjectNotFound", StringComparison.OrdinalIgnoreCase);
    }

    public async Task<MailboxInfo?> ConvertToSharedAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var cfg = await RequireConfigAsync(cancellationToken).ConfigureAwait(false);
        var id = Lit(identity);
        var body = $$"""
            $id = {{id}}
            $cur = Get-Mailbox -Identity $id -ErrorAction Stop
            if ($cur.RecipientTypeDetails -ne 'SharedMailbox') {
                Set-Mailbox -Identity $id -Type Shared -ErrorAction Stop
                Write-Output 'Set-Mailbox -Type Shared aplicado; esperando propagacion...'
            }
            $final = [string]$cur.RecipientTypeDetails
            $deadline = (Get-Date).AddSeconds(150)
            while ($final -ne 'SharedMailbox' -and (Get-Date) -lt $deadline) {
                Start-Sleep -Seconds 8
                try { $final = [string](Get-Mailbox -Identity $id -ErrorAction Stop).RecipientTypeDetails } catch { }
            }
            $m = Get-Mailbox -Identity $id -ErrorAction Stop
            $o = [PSCustomObject]@{
                Identity             = [string]$m.Identity
                DisplayName          = [string]$m.DisplayName
                PrimarySmtpAddress   = [string]$m.PrimarySmtpAddress
                RecipientTypeDetails = [string]$m.RecipientTypeDetails
            }
            Write-Output ('{{JsonMarker}}' + ($o | ConvertTo-Json -Compress))
            """;

        var json = await RunAsync(cfg, body, progress, cancellationToken).ConfigureAwait(false);
        return Parse(json);
    }

    public async Task SetAutoReplyAsync(
        string identity,
        string message,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var cfg = await RequireConfigAsync(cancellationToken).ConfigureAwait(false);
        var body = $$"""
            $id = {{Lit(identity)}}
            $msg = {{Lit(message)}}
            Set-MailboxAutoReplyConfiguration -Identity $id -AutoReplyState Enabled -ExternalAudience All -InternalMessage $msg -ExternalMessage $msg -ErrorAction Stop
            Write-Output ('{{JsonMarker}}' + (([PSCustomObject]@{ Note = 'auto-reply Enabled' }) | ConvertTo-Json -Compress))
            """;
        await RunAsync(cfg, body, progress, cancellationToken).ConfigureAwait(false);
    }

    public async Task SetForwardingAsync(
        string identity,
        string forwardTo,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var cfg = await RequireConfigAsync(cancellationToken).ConfigureAwait(false);
        var body = $$"""
            $id = {{Lit(identity)}}
            Set-Mailbox -Identity $id -ForwardingSmtpAddress {{Lit(forwardTo)}} -DeliverToMailboxAndForward $true -ErrorAction Stop
            Write-Output ('{{JsonMarker}}' + (([PSCustomObject]@{ Note = 'forward configurado' }) | ConvertTo-Json -Compress))
            """;
        await RunAsync(cfg, body, progress, cancellationToken).ConfigureAwait(false);
    }

    public async Task<string> HideFromGalAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var cfg = await RequireConfigAsync(cancellationToken).ConfigureAwait(false);
        // Hybrid objects synced from on-prem AD can't be modified in EXO — detect that specific
        // failure and report it as a SKIP-with-guidance instead of a hard error (legacy parity).
        var body = $$"""
            $id = {{Lit(identity)}}
            try {
                Set-Mailbox -Identity $id -HiddenFromAddressListsEnabled $true -ErrorAction Stop
                $note = 'HiddenFromAddressListsEnabled=true'
            } catch {
                $m = $_.Exception.Message
                if ($m -match 'sincroniz|synchroniz|on-prem|on premises|out of.*write scope|write scope|.mbito de escritura|organizaci.n (local|interna)|local organization|cannot be performed.*synchron') {
                    $note = 'objeto hibrido sincronizado desde AD on-prem: aplicar msExchHideFromAddressLists=TRUE en AD local y esperar a Entra Connect'
                } else { throw }
            }
            Write-Output ('{{JsonMarker}}' + (([PSCustomObject]@{ Note = $note }) | ConvertTo-Json -Compress))
            """;
        var json = await RunAsync(cfg, body, progress, cancellationToken).ConfigureAwait(false);
        return ParseNote(json) ?? "aplicado";
    }

    public async Task<string> GrantDelegateAsync(
        string mailbox,
        string delegateUpn,
        bool sendAs,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var cfg = await RequireConfigAsync(cancellationToken).ConfigureAwait(false);
        var sendAsLit = sendAs ? "$true" : "$false";
        var body = $$"""
            $id = {{Lit(mailbox)}}
            $d = {{Lit(delegateUpn)}}
            Add-MailboxPermission -Identity $id -User $d -AccessRights FullAccess -InheritanceType All -AutoMapping:$false -Confirm:$false -ErrorAction Stop | Out-Null
            $note = 'FullAccess'
            if ({{sendAsLit}}) {
                Add-RecipientPermission -Identity $id -Trustee $d -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                $note = 'FullAccess + SendAs'
            }
            Write-Output ('{{JsonMarker}}' + (([PSCustomObject]@{ Note = ($note + ' -> ' + $d) }) | ConvertTo-Json -Compress))
            """;
        var json = await RunAsync(cfg, body, progress, cancellationToken).ConfigureAwait(false);
        return ParseNote(json) ?? "delegado";
    }

    public async Task<MailboxInfo?> ConvertToRegularAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var cfg = await RequireConfigAsync(cancellationToken).ConfigureAwait(false);
        var id = Lit(identity);
        var body = $$"""
            $id = {{id}}
            $cur = Get-Mailbox -Identity $id -ErrorAction Stop
            if ($cur.RecipientTypeDetails -eq 'SharedMailbox') {
                Set-Mailbox -Identity $id -Type Regular -ErrorAction Stop
                Write-Output 'Set-Mailbox -Type Regular aplicado; esperando propagacion...'
            }
            $final = [string]$cur.RecipientTypeDetails
            $deadline = (Get-Date).AddSeconds(150)
            while ($final -eq 'SharedMailbox' -and (Get-Date) -lt $deadline) {
                Start-Sleep -Seconds 8
                try { $final = [string](Get-Mailbox -Identity $id -ErrorAction Stop).RecipientTypeDetails } catch { }
            }
            $m = Get-Mailbox -Identity $id -ErrorAction Stop
            $o = [PSCustomObject]@{
                Identity             = [string]$m.Identity
                DisplayName          = [string]$m.DisplayName
                PrimarySmtpAddress   = [string]$m.PrimarySmtpAddress
                RecipientTypeDetails = [string]$m.RecipientTypeDetails
            }
            Write-Output ('{{JsonMarker}}' + ($o | ConvertTo-Json -Compress))
            """;
        var json = await RunAsync(cfg, body, progress, cancellationToken).ConfigureAwait(false);
        return Parse(json);
    }

    // Pure: maps (action, permission) → the EXO cmdlet operating on $m (mailbox) / $p (principal).
    // Returns null for an unsupported combination. Kept public+static so it stays unit-testable.
    public static string? BuildPermissionCmdlet(string action, string permission)
    {
        var add = string.Equals(action, "add", StringComparison.OrdinalIgnoreCase);
        return permission switch
        {
            "FullAccess" => add
                ? "Add-MailboxPermission -Identity $m -User $p -AccessRights FullAccess -InheritanceType All -AutoMapping:$true -Confirm:$false -ErrorAction Stop | Out-Null"
                : "Remove-MailboxPermission -Identity $m -User $p -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop | Out-Null",
            "SendAs" => add
                ? "Add-RecipientPermission -Identity $m -Trustee $p -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null"
                : "Remove-RecipientPermission -Identity $m -Trustee $p -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null",
            "SendOnBehalf" => add
                ? "Set-Mailbox -Identity $m -GrantSendOnBehalfTo @{Add=$p} -ErrorAction Stop | Out-Null"
                : "Set-Mailbox -Identity $m -GrantSendOnBehalfTo @{Remove=$p} -ErrorAction Stop | Out-Null",
            _ => null,
        };
    }

    public async Task<MailboxPermissionResult> ApplyPermissionAsync(
        string action,
        string permission,
        string mailbox,
        string principal,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var cmdlet = BuildPermissionCmdlet(action, permission);
        if (cmdlet is null)
        {
            return new MailboxPermissionResult(action, permission, mailbox, principal, "INVALIDO", "Permission no soportada");
        }
        var cfg = await RequireConfigAsync(cancellationToken).ConfigureAwait(false);
        var body = $$"""
            $m = {{Lit(mailbox)}}
            $p = {{Lit(principal)}}
            {{cmdlet}}
            Write-Output ('{{JsonMarker}}' + (([PSCustomObject]@{ Note = 'OK' }) | ConvertTo-Json -Compress))
            """;
        try
        {
            await RunAsync(cfg, body, progress, cancellationToken).ConfigureAwait(false);
            return new MailboxPermissionResult(action, permission, mailbox, principal, "OK", "Aplicado");
        }
        catch (Exception ex)
        {
            return new MailboxPermissionResult(action, permission, mailbox, principal, "ERROR", ex.Message);
        }
    }

    public async Task<IReadOnlyList<MailboxPermissionEntry>> GetPermissionsAsync(
        string mailbox,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var cfg = await RequireConfigAsync(cancellationToken).ConfigureAwait(false);
        var body = $$"""
            $mbx = {{Lit(mailbox)}}
            $full = @(Get-MailboxPermission -Identity $mbx -ErrorAction Stop |
                Where-Object { $_.AccessRights -contains 'FullAccess' -and -not $_.IsInherited -and $_.User -notlike 'NT AUTHORITY\SELF' })
            $send = @(Get-RecipientPermission -Identity $mbx -ErrorAction SilentlyContinue |
                Where-Object { $_.AccessRights -contains 'SendAs' -and $_.Trustee -notlike 'NT AUTHORITY\SELF' })
            $m = Get-Mailbox -Identity $mbx -ErrorAction Stop
            $out = New-Object System.Collections.Generic.List[object]
            foreach ($f in $full) { $out.Add([PSCustomObject]@{ Permission='FullAccess';   Principal=[string]$f.User;    Detail=[string]$f.AccessRights }) }
            foreach ($s in $send) { $out.Add([PSCustomObject]@{ Permission='SendAs';       Principal=[string]$s.Trustee; Detail=[string]$s.AccessRights }) }
            foreach ($o in @($m.GrantSendOnBehalfTo)) { $out.Add([PSCustomObject]@{ Permission='SendOnBehalf'; Principal=[string]$o; Detail='From Set-Mailbox' }) }
            Write-Output ('{{JsonMarker}}' + ($out | ConvertTo-Json -Compress -AsArray))
            """;
        var json = await RunAsync(cfg, body, progress, cancellationToken).ConfigureAwait(false);
        return ParsePermissions(json);
    }

    // ---- Mailbox rules: auto-reply / forwarding / calendar (external pwsh) ----

    public async Task<AutoReplyConfig?> GetAutoReplyAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var cfg = await RequireConfigAsync(cancellationToken).ConfigureAwait(false);
        var body = $$"""
            $id = {{Lit(identity)}}
            $c = Get-MailboxAutoReplyConfiguration -Identity $id -ErrorAction Stop
            $o = [PSCustomObject]@{
                State           = [string]$c.AutoReplyState
                InternalMessage = [string]$c.InternalMessage
                ExternalMessage = [string]$c.ExternalMessage
                StartTime       = if ($c.StartTime) { ([datetime]$c.StartTime).ToString('o') } else { '' }
                EndTime         = if ($c.EndTime) { ([datetime]$c.EndTime).ToString('o') } else { '' }
            }
            Write-Output ('{{JsonMarker}}' + ($o | ConvertTo-Json -Compress))
            """;
        var json = await RunAsync(cfg, body, progress, cancellationToken).ConfigureAwait(false);
        return ParseAutoReply(json);
    }

    public async Task SetAutoReplyConfigAsync(
        string identity,
        AutoReplyConfig config,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var cfg = await RequireConfigAsync(cancellationToken).ConfigureAwait(false);
        var body = $$"""
            $id = {{Lit(identity)}}
            $params = @{ Identity = $id; AutoReplyState = {{Lit(config.State.ToString())}}; ErrorAction = 'Stop' }
            $int = {{Lit(config.InternalMessage)}}
            $ext = {{Lit(config.ExternalMessage)}}
            $start = {{Lit(config.StartTime?.ToString("o", CultureInfo.InvariantCulture))}}
            $end = {{Lit(config.EndTime?.ToString("o", CultureInfo.InvariantCulture))}}
            if ($int)   { $params['InternalMessage'] = $int }
            if ($ext)   { $params['ExternalMessage'] = $ext }
            if ($start) { $params['StartTime'] = [datetime]::Parse($start, [System.Globalization.CultureInfo]::InvariantCulture) }
            if ($end)   { $params['EndTime']   = [datetime]::Parse($end,   [System.Globalization.CultureInfo]::InvariantCulture) }
            Set-MailboxAutoReplyConfiguration @params | Out-Null
            Write-Output ('{{JsonMarker}}' + (([PSCustomObject]@{ Note = ('auto-reply ' + {{Lit(config.State.ToString())}}) }) | ConvertTo-Json -Compress))
            """;
        await RunAsync(cfg, body, progress, cancellationToken).ConfigureAwait(false);
    }

    public async Task<ForwardingConfig?> GetForwardingAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var cfg = await RequireConfigAsync(cancellationToken).ConfigureAwait(false);
        var body = $$"""
            $id = {{Lit(identity)}}
            $m = Get-Mailbox -Identity $id -ErrorAction Stop
            $o = [PSCustomObject]@{
                ForwardingAddress          = [string]$m.ForwardingAddress
                ForwardingSmtpAddress      = [string]$m.ForwardingSmtpAddress
                DeliverToMailboxAndForward = [bool]$m.DeliverToMailboxAndForward
            }
            Write-Output ('{{JsonMarker}}' + ($o | ConvertTo-Json -Compress))
            """;
        var json = await RunAsync(cfg, body, progress, cancellationToken).ConfigureAwait(false);
        return ParseForwarding(json);
    }

    public async Task ConfigureForwardingAsync(
        string identity,
        string forwardingSmtpAddress,
        bool deliverToMailboxAndForward,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var cfg = await RequireConfigAsync(cancellationToken).ConfigureAwait(false);
        var deliverLit = deliverToMailboxAndForward ? "$true" : "$false";
        var body = $$"""
            $id = {{Lit(identity)}}
            Set-Mailbox -Identity $id -ForwardingSmtpAddress {{Lit(forwardingSmtpAddress)}} -DeliverToMailboxAndForward:{{deliverLit}} -ErrorAction Stop | Out-Null
            Write-Output ('{{JsonMarker}}' + (([PSCustomObject]@{ Note = 'forwarding configurado' }) | ConvertTo-Json -Compress))
            """;
        await RunAsync(cfg, body, progress, cancellationToken).ConfigureAwait(false);
    }

    public async Task ClearForwardingAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var cfg = await RequireConfigAsync(cancellationToken).ConfigureAwait(false);
        var body = $$"""
            $id = {{Lit(identity)}}
            Set-Mailbox -Identity $id -ForwardingAddress $null -ForwardingSmtpAddress $null -DeliverToMailboxAndForward:$false -ErrorAction Stop | Out-Null
            Write-Output ('{{JsonMarker}}' + (([PSCustomObject]@{ Note = 'forwarding limpiado' }) | ConvertTo-Json -Compress))
            """;
        await RunAsync(cfg, body, progress, cancellationToken).ConfigureAwait(false);
    }

    public async Task<IReadOnlyList<CalendarPermissionEntry>> GetCalendarPermissionsAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var cfg = await RequireConfigAsync(cancellationToken).ConfigureAwait(false);
        // Filter the noise (Default/None, anonymous, empties) in PowerShell so the JSON is clean.
        var body = $$"""
            $id = {{Lit(identity)}}
            $folder = "{0}:\Calendar" -f $id
            $perms = @(Get-MailboxFolderPermission -Identity $folder -ErrorAction Stop)
            $out = New-Object System.Collections.Generic.List[object]
            foreach ($p in $perms) {
                $principal = [string]$p.User
                $rights = ([string]::Join(',', @($p.AccessRights)))
                if (-not $principal -or -not $rights) { continue }
                if ($principal -eq 'Default' -and $rights -eq 'None') { continue }
                $out.Add([PSCustomObject]@{ Principal = $principal; AccessRights = $rights })
            }
            Write-Output ('{{JsonMarker}}' + ($out | ConvertTo-Json -Compress -AsArray))
            """;
        var json = await RunAsync(cfg, body, progress, cancellationToken).ConfigureAwait(false);
        return ParseCalendarPermissions(json);
    }

    public async Task ApplyCalendarPermissionAsync(
        string identity,
        string principal,
        string accessRights,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var cfg = await RequireConfigAsync(cancellationToken).ConfigureAwait(false);
        var body = $$"""
            $id = {{Lit(identity)}}
            $p = {{Lit(principal)}}
            $rights = {{Lit(accessRights)}}
            $folder = "{0}:\Calendar" -f $id
            $existing = Get-MailboxFolderPermission -Identity $folder -User $p -ErrorAction SilentlyContinue
            if ($existing) {
                Set-MailboxFolderPermission -Identity $folder -User $p -AccessRights $rights -ErrorAction Stop | Out-Null
                $note = ('actualizado: ' + $p + ' -> ' + $rights)
            } else {
                Add-MailboxFolderPermission -Identity $folder -User $p -AccessRights $rights -ErrorAction Stop | Out-Null
                $note = ('añadido: ' + $p + ' -> ' + $rights)
            }
            Write-Output ('{{JsonMarker}}' + (([PSCustomObject]@{ Note = $note }) | ConvertTo-Json -Compress))
            """;
        await RunAsync(cfg, body, progress, cancellationToken).ConfigureAwait(false);
    }

    public async Task RemoveCalendarPermissionAsync(
        string identity,
        string principal,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var cfg = await RequireConfigAsync(cancellationToken).ConfigureAwait(false);
        var body = $$"""
            $id = {{Lit(identity)}}
            $p = {{Lit(principal)}}
            $folder = "{0}:\Calendar" -f $id
            Remove-MailboxFolderPermission -Identity $folder -User $p -Confirm:$false -ErrorAction Stop | Out-Null
            Write-Output ('{{JsonMarker}}' + (([PSCustomObject]@{ Note = ('eliminado: ' + $p) }) | ConvertTo-Json -Compress))
            """;
        await RunAsync(cfg, body, progress, cancellationToken).ConfigureAwait(false);
    }

    private static IReadOnlyList<MailboxPermissionEntry> ParsePermissions(string? json)
    {
        var list = new List<MailboxPermissionEntry>();
        if (string.IsNullOrWhiteSpace(json)) return list;
        using var doc = JsonDocument.Parse(json);
        if (doc.RootElement.ValueKind != JsonValueKind.Array) return list;
        foreach (var el in doc.RootElement.EnumerateArray())
        {
            string S(string n) => el.TryGetProperty(n, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString()! : string.Empty;
            list.Add(new MailboxPermissionEntry(S("Permission"), S("Principal"), S("Detail")));
        }
        return list;
    }

    private static AutoReplyConfig? ParseAutoReply(string? json)
    {
        if (string.IsNullOrWhiteSpace(json)) return null;
        using var doc = JsonDocument.Parse(json);
        var r = doc.RootElement;
        string S(string n) => r.TryGetProperty(n, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString()! : string.Empty;
        var state = S("State") switch
        {
            "Enabled" => AutoReplyState.Enabled,
            "Scheduled" => AutoReplyState.Scheduled,
            _ => AutoReplyState.Disabled,
        };
        DateTime? D(string n)
        {
            var s = S(n);
            return DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var dt) ? dt : null;
        }
        static string? NullIfEmpty(string s) => string.IsNullOrEmpty(s) ? null : s;
        return new AutoReplyConfig(state, NullIfEmpty(S("InternalMessage")), NullIfEmpty(S("ExternalMessage")), D("StartTime"), D("EndTime"));
    }

    private static ForwardingConfig? ParseForwarding(string? json)
    {
        if (string.IsNullOrWhiteSpace(json)) return null;
        using var doc = JsonDocument.Parse(json);
        var r = doc.RootElement;
        string S(string n) => r.TryGetProperty(n, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString()! : string.Empty;
        bool B(string n) => r.TryGetProperty(n, out var v) && (v.ValueKind == JsonValueKind.True || (v.ValueKind == JsonValueKind.String && bool.TryParse(v.GetString(), out var b) && b));
        static string? NullIfEmpty(string s) => string.IsNullOrEmpty(s) ? null : s;
        return new ForwardingConfig(NullIfEmpty(S("ForwardingAddress")), NullIfEmpty(S("ForwardingSmtpAddress")), B("DeliverToMailboxAndForward"));
    }

    private static IReadOnlyList<CalendarPermissionEntry> ParseCalendarPermissions(string? json)
    {
        var list = new List<CalendarPermissionEntry>();
        if (string.IsNullOrWhiteSpace(json)) return list;
        using var doc = JsonDocument.Parse(json);
        if (doc.RootElement.ValueKind != JsonValueKind.Array) return list;
        foreach (var el in doc.RootElement.EnumerateArray())
        {
            string S(string n) => el.TryGetProperty(n, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString()! : string.Empty;
            var principal = S("Principal");
            var rights = S("AccessRights");
            if (string.IsNullOrEmpty(principal) || string.IsNullOrEmpty(rights)) continue;
            list.Add(new CalendarPermissionEntry(principal, rights));
        }
        return list;
    }

    private static string? ParseNote(string? json)
    {
        if (string.IsNullOrWhiteSpace(json)) return null;
        using var doc = JsonDocument.Parse(json);
        return doc.RootElement.TryGetProperty("Note", out var v) && v.ValueKind == JsonValueKind.String
            ? v.GetString()
            : null;
    }

    private async Task<CertConfig> RequireConfigAsync(CancellationToken ct)
    {
        var cfg = await _certStore.LoadAsync(ct).ConfigureAwait(false);
        if (cfg is null || string.IsNullOrWhiteSpace(cfg.CertThumbprint))
        {
            throw new InvalidOperationException(
                "No hay configuración de certificado para Exchange Online. Conéctate por certificado primero.");
        }
        return cfg;
    }

    // Wraps the body in connect/disconnect and runs it in an external pwsh process.
    private async Task<string?> RunAsync(
        CertConfig cfg,
        string body,
        IProgress<LogEntry>? progress,
        CancellationToken cancellationToken)
    {
        var full = $$"""
            $ErrorActionPreference = 'Stop'
            $ProgressPreference = 'SilentlyContinue'
            $InformationPreference = 'SilentlyContinue'
            $WarningPreference = 'SilentlyContinue'
            [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
            Import-Module ExchangeOnlineManagement -ErrorAction Stop
            Connect-ExchangeOnline -AppId {{Lit(cfg.AppId)}} -CertificateThumbprint {{Lit(cfg.CertThumbprint)}} -Organization {{Lit(cfg.Organization)}} -ShowBanner:$false -InformationAction SilentlyContinue -ErrorAction Stop | Out-Null
            try {
            {{body}}
            }
            catch {
                Write-Output ('{{ErrMarker}}' + $_.Exception.Message)
            }
            finally {
                Disconnect-ExchangeOnline -Confirm:$false -InformationAction SilentlyContinue -ErrorAction SilentlyContinue *> $null
            }
            """;

        var encoded = Convert.ToBase64String(Encoding.Unicode.GetBytes(full));

        var psi = new ProcessStartInfo
        {
            FileName = ExeResolver.ResolvePwsh(),
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
            CreateNoWindow = true,
        };
        psi.ArgumentList.Add("-NoLogo");
        psi.ArgumentList.Add("-NoProfile");
        psi.ArgumentList.Add("-NonInteractive");
        psi.ArgumentList.Add("-EncodedCommand");
        psi.ArgumentList.Add(encoded);

        using var proc = new Process { StartInfo = psi };
        string? jsonLine = null;
        string? scriptError = null;
        var errors = new StringBuilder();

        proc.OutputDataReceived += (_, e) =>
        {
            if (string.IsNullOrEmpty(e.Data)) return;
            // The script wraps the body in try/catch and emits the real exception message on a
            // marked line — capture it so failures surface the actual EXO error instead of a
            // bare "pwsh exit 1".
            var errIdx = e.Data.IndexOf(ErrMarker, StringComparison.Ordinal);
            if (errIdx >= 0)
            {
                scriptError = e.Data[(errIdx + ErrMarker.Length)..];
                return;
            }
            var idx = e.Data.IndexOf(JsonMarker, StringComparison.Ordinal);
            if (idx >= 0)
            {
                jsonLine = e.Data[(idx + JsonMarker.Length)..];
            }
            else
            {
                progress?.Report(LogEntry.Info("EXO", e.Data));
            }
        };
        proc.ErrorDataReceived += (_, e) =>
        {
            if (string.IsNullOrEmpty(e.Data)) return;
            // The EXO module serialises non-text stream records to stderr as CLIXML noise
            // ("#< CLIXML", "<Objs ...>"). It is not an error — drop it so the live log
            // stays clean instead of showing scary XML warnings.
            var t = e.Data.TrimStart();
            if (t.StartsWith("#< CLIXML", StringComparison.Ordinal) ||
                t.StartsWith("<Objs", StringComparison.Ordinal) ||
                t.StartsWith("<", StringComparison.Ordinal))
            {
                return;
            }
            errors.AppendLine(e.Data);
            progress?.Report(LogEntry.Warn("EXO", e.Data));
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

        if (scriptError is not null)
        {
            throw new InvalidOperationException("Operación Exchange Online falló: " + scriptError.Trim());
        }
        if (proc.ExitCode != 0 && jsonLine is null)
        {
            throw new InvalidOperationException(
                "Operación Exchange Online falló: " + (errors.Length > 0 ? errors.ToString().Trim() : $"pwsh exit {proc.ExitCode}"));
        }
        return jsonLine;
    }

    private static MailboxInfo? Parse(string? json)
    {
        if (string.IsNullOrWhiteSpace(json)) return null;
        using var doc = JsonDocument.Parse(json);
        var r = doc.RootElement;
        string S(string n) => r.TryGetProperty(n, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString()! : string.Empty;
        bool B(string n) => r.TryGetProperty(n, out var v) && (v.ValueKind == JsonValueKind.True || (v.ValueKind == JsonValueKind.String && bool.TryParse(v.GetString(), out var b) && b));
        int I(string n) => r.TryGetProperty(n, out var v) && v.TryGetInt32(out var i) ? i : 0;
        long? L(string n) => r.TryGetProperty(n, out var v) && v.ValueKind == JsonValueKind.Number && v.TryGetInt64(out var l) ? l : null;

        return new MailboxInfo(
            Identity: S("Identity"),
            DisplayName: S("DisplayName"),
            PrimarySmtpAddress: S("PrimarySmtpAddress"),
            RecipientTypeDetails: S("RecipientTypeDetails"),
            LitigationHoldEnabled: B("LitigationHoldEnabled"),
            InPlaceHoldCount: I("InPlaceHoldCount"),
            ArchiveEnabled: B("ArchiveEnabled"),
            TotalItemBytes: L("TotalItemBytes"));
    }

    // PowerShell single-quoted literal (doubles embedded quotes) — safe for injection.
    private static string Lit(string? value) => "'" + (value ?? string.Empty).Replace("'", "''") + "'";
}
