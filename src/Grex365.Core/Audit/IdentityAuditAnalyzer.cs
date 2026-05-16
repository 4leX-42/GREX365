using Grex365.Core.Models;

namespace Grex365.Core.Audit;

public sealed class IdentityAuditAnalyzer
{
    public static readonly TimeSpan MemberStaleAfter = TimeSpan.FromDays(180);
    public static readonly TimeSpan GuestStaleAfter = TimeSpan.FromDays(90);

    private readonly DateTimeOffset _now;
    private readonly List<AuditFinding> _findings = new();
    private readonly Totals _totals = new();

    public IdentityAuditAnalyzer(DateTimeOffset now)
    {
        _now = now;
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
                "Disabled+License", upn,
                $"Deshabilitado con {user.AssignedLicenseCount} licencias asignadas", "WARN"));
        }

        if (user.AccountEnabled)
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
    DateTimeOffset? LastSignIn);
