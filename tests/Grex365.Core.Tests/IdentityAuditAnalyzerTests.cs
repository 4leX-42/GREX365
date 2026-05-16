using FluentAssertions;
using Grex365.Core.Audit;

namespace Grex365.Core.Tests;

public class IdentityAuditAnalyzerTests
{
    private static readonly DateTimeOffset Now = new(2026, 5, 16, 0, 0, 0, TimeSpan.Zero);

    [Fact]
    public void EnabledMember_NoSignIn_IsStale()
    {
        var a = new IdentityAuditAnalyzer(Now);
        a.Visit(new UserSnapshot("id", "u@a", true, false, 0, null));
        a.Findings.Should().ContainSingle(f => f.Category == "Stale member");
        a.BuildSummary().StaleMembers.Should().Be(1);
    }

    [Fact]
    public void EnabledMember_RecentSignIn_NotStale()
    {
        var a = new IdentityAuditAnalyzer(Now);
        a.Visit(new UserSnapshot("id", "u@a", true, false, 0, Now.AddDays(-30)));
        a.Findings.Should().BeEmpty();
        a.BuildSummary().StaleMembers.Should().Be(0);
    }

    [Fact]
    public void EnabledMember_OldSignIn_IsStale()
    {
        var a = new IdentityAuditAnalyzer(Now);
        a.Visit(new UserSnapshot("id", "u@a", true, false, 0, Now.AddDays(-200)));
        a.Findings.Should().ContainSingle(f => f.Category == "Stale member");
    }

    [Fact]
    public void EnabledGuest_OlderThan90Days_IsStale()
    {
        var a = new IdentityAuditAnalyzer(Now);
        a.Visit(new UserSnapshot("id", "guest#EXT#@a", true, true, 0, Now.AddDays(-100)));
        a.Findings.Should().ContainSingle(f => f.Category == "Stale guest");
        a.BuildSummary().StaleGuests.Should().Be(1);
    }

    [Fact]
    public void EnabledGuest_Within90Days_NotStale()
    {
        var a = new IdentityAuditAnalyzer(Now);
        a.Visit(new UserSnapshot("id", "guest#EXT#@a", true, true, 0, Now.AddDays(-89)));
        a.Findings.Should().BeEmpty();
    }

    [Fact]
    public void DisabledWithLicense_Flagged()
    {
        var a = new IdentityAuditAnalyzer(Now);
        a.Visit(new UserSnapshot("id", "u@a", false, false, 3, null));
        a.Findings.Should().ContainSingle(f => f.Category == "Disabled+License" && f.Detail.Contains("3 licencias"));
        a.BuildSummary().DisabledWithLicense.Should().Be(1);
    }

    [Fact]
    public void DisabledWithoutLicense_NotFlagged_AndDoesNotCountAsStale()
    {
        var a = new IdentityAuditAnalyzer(Now);
        a.Visit(new UserSnapshot("id", "u@a", false, false, 0, null));
        a.Findings.Should().BeEmpty();
        var s = a.BuildSummary();
        s.UsersDisabled.Should().Be(1);
        s.StaleMembers.Should().Be(0);
    }

    [Fact]
    public void Totals_AccumulateCorrectly()
    {
        var a = new IdentityAuditAnalyzer(Now);
        a.Visit(new UserSnapshot("1", "a@a", true, false, 1, Now.AddDays(-1)));
        a.Visit(new UserSnapshot("2", "b@a", false, false, 2, null));
        a.Visit(new UserSnapshot("3", "g@a", true, true, 0, Now.AddDays(-200)));
        var s = a.BuildSummary();
        s.UsersTotal.Should().Be(3);
        s.UsersEnabled.Should().Be(2);
        s.UsersDisabled.Should().Be(1);
        s.Guests.Should().Be(1);
        s.StaleGuests.Should().Be(1);
        s.DisabledWithLicense.Should().Be(1);
    }

    [Fact]
    public void UsesUserPrincipalName_FallsBackToId()
    {
        var a = new IdentityAuditAnalyzer(Now);
        a.Visit(new UserSnapshot("only-id", null, false, false, 2, null));
        a.Findings[0].Identity.Should().Be("only-id");
    }
}
