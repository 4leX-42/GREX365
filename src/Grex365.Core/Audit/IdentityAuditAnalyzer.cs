using Grex365.Core.Models;

namespace Grex365.Core.Audit;

public sealed class IdentityAuditAnalyzer
{
    public static readonly TimeSpan MemberStaleAfter = TimeSpan.FromDays(180);
    public static readonly TimeSpan GuestStaleAfter = TimeSpan.FromDays(90);

    private readonly DateTimeOffset _now;
    private readonly bool _signInActivityAvailable;
    private readonly IReadOnlyDictionary<Guid, string>? _skuPartNumbers;
    private readonly List<AuditFinding> _findings = new();
    private readonly Totals _totals = new();

    // signInActivityAvailable=false when the tenant lacks AuditLog.Read.All: last-sign-in
    // data is then null for EVERY user, so we must NOT flag stale (it would falsely mark
    // the whole directory). skuPartNumbers maps assigned licence GUIDs to SKU part numbers
    // so the Disabled+License finding can name the actual licences.
    public IdentityAuditAnalyzer(
        DateTimeOffset now,
        bool signInActivityAvailable = true,
        IReadOnlyDictionary<Guid, string>? skuPartNumbers = null)
    {
        _now = now;
        _signInActivityAvailable = signInActivityAvailable;
        _skuPartNumbers = skuPartNumbers;
    }

    public IReadOnlyList<AuditFinding> Findings => _findings;

    public AuditSummary BuildSummary() => new(
        UsersTotal: _totals.UsersTotal,
        UsersEnabled: _totals.UsersEnabled,
        UsersDisabled: _totals.UsersDisabled,
        Guests: _totals.Guests,
        StaleMembers: _totals.StaleMembers,
        StaleGuests: _totals.StaleGuests,
        DisabledWithLicense: _totals.DisabledWithLicense);

    public void Visit(UserSnapshot user)
    {
        _totals.UsersTotal++;
        var upn = user.UserPrincipalName ?? user.Id ?? "(desconocido)";
        if (user.IsGuest) _totals.Guests++;
        if (user.AccountEnabled) _totals.UsersEnabled++; else _totals.UsersDisabled++;

        if (!user.AccountEnabled && user.AssignedLicenseCount > 0)
        {
            _totals.DisabledWithLicense++;
            _findings.Add(new AuditFinding(
                AuditFinding.DisabledWithLicenseCategory, upn,
                $"Deshabilitado · {user.AssignedLicenseCount} licencias{DescribeLicenses(user)}", "WARN"));
        }

        // Stale detection requires real last-sign-in data. Skip entirely when unavailable
        // (AuditLog.Read.All missing) — otherwise every enabled user would be flagged.
        if (user.AccountEnabled && _signInActivityAvailable)
        {
            var cutoff = user.IsGuest ? _now - GuestStaleAfter : _now - MemberStaleAfter;
            if (user.LastSignIn is null || user.LastSignIn < cutoff)
            {
                var lastTxt = user.LastSignIn?.ToString("yyyy-MM-dd") ?? "nunca";
                var threshold = user.IsGuest ? "90d" : "180d";
                if (user.IsGuest)
                {
                    _totals.StaleGuests++;
                    _findings.Add(new AuditFinding("Stale guest", upn,
                        $"último login: {lastTxt} (>{threshold})", "WARN"));
                }
                else
                {
                    _totals.StaleMembers++;
                    _findings.Add(new AuditFinding("Stale member", upn,
                        $"último login: {lastTxt} (>{threshold})", "WARN"));
                }
            }
        }
    }

    // ": Microsoft 365 E5, Power BI Pro" when SKU names can be resolved, else "".
    private string DescribeLicenses(UserSnapshot user)
    {
        if (_skuPartNumbers is null || user.AssignedSkuIds is null || user.AssignedSkuIds.Count == 0)
        {
            return " asignadas";
        }
        var names = user.AssignedSkuIds
            .Select(id => _skuPartNumbers.TryGetValue(id, out var part) ? part : id.ToString())
            .Select(part => SkuCatalog.Resolve(part).FriendlyName)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
            .ToList();
        return names.Count == 0 ? " asignadas" : ": " + string.Join(", ", names);
    }

    private sealed class Totals
    {
        public int UsersTotal;
        public int UsersEnabled;
        public int UsersDisabled;
        public int Guests;
        public int StaleMembers;
        public int StaleGuests;
        public int DisabledWithLicense;
    }
}

public sealed record UserSnapshot(
    string? Id,
    string? UserPrincipalName,
    bool AccountEnabled,
    bool IsGuest,
    int AssignedLicenseCount,
    DateTimeOffset? LastSignIn,
    IReadOnlyList<Guid>? AssignedSkuIds = null);
