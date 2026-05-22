using FluentAssertions;
using Grex365.Core.Audit;

namespace Grex365.Core.Tests;

public class TenantDefaultsAnalyzerTests
{
    private static AuthorizationPolicySnapshot Safe() => new(
        AllowedToSignUpEmailBasedSubscriptions: false,
        AllowedToUseSspr: true,
        AllowEmailVerifiedUsersToJoinOrganization: false,
        AllowInvitesFrom: "adminsAndGuestInviters",
        DefaultUserCanCreateApps: false,
        DefaultUserCanCreateSecurityGroups: false,
        DefaultUserCanCreateTenants: false,
        DefaultUserCanReadOtherUsers: true);

    [Fact]
    public void SafeBaseline_NoFindings()
    {
        var (summary, findings) = TenantDefaultsAnalyzer.Analyze(Safe(), securityDefaultsEnabled: false);
        summary.Findings.Should().Be(0);
        findings.Should().BeEmpty();
    }

    [Fact]
    public void UsersCanCreateApps_FlaggedWarn()
    {
        var p = Safe() with { DefaultUserCanCreateApps = true };
        var (_, findings) = TenantDefaultsAnalyzer.Analyze(p, false);
        findings.Should().ContainSingle();
        findings[0].Category.Should().Contain("create apps");
        findings[0].Severity.Should().Be("WARN");
    }

    [Fact]
    public void UsersCanCreateTenants_FlaggedWarn()
    {
        var p = Safe() with { DefaultUserCanCreateTenants = true };
        var (_, findings) = TenantDefaultsAnalyzer.Analyze(p, false);
        findings.Should().Contain(f => f.Category.Contains("create tenants") && f.Severity == "WARN");
    }

    [Fact]
    public void EmailVerifiedJoin_FlaggedWarn()
    {
        var p = Safe() with { AllowEmailVerifiedUsersToJoinOrganization = true };
        var (_, findings) = TenantDefaultsAnalyzer.Analyze(p, false);
        findings.Should().Contain(f => f.Category.Contains("self-join") && f.Severity == "WARN");
    }

    [Fact]
    public void InvitesFromEveryone_FlaggedWarn()
    {
        var p = Safe() with { AllowInvitesFrom = "everyone" };
        var (_, findings) = TenantDefaultsAnalyzer.Analyze(p, false);
        findings.Should().Contain(f => f.Category.Contains("invite guests") && f.Severity == "WARN");
    }

    [Fact]
    public void InvitesFromAdminsOnly_NotFlagged()
    {
        var p = Safe() with { AllowInvitesFrom = "adminsOnly" };
        var (_, findings) = TenantDefaultsAnalyzer.Analyze(p, false);
        findings.Should().BeEmpty();
    }

    [Fact]
    public void SsprDisabled_FlaggedInfo()
    {
        var p = Safe() with { AllowedToUseSspr = false };
        var (_, findings) = TenantDefaultsAnalyzer.Analyze(p, false);
        findings.Should().Contain(f => f.Category.Contains("SSPR") && f.Severity == "INFO");
    }

    [Fact]
    public void EmailBasedSubs_FlaggedInfo()
    {
        var p = Safe() with { AllowedToSignUpEmailBasedSubscriptions = true };
        var (_, findings) = TenantDefaultsAnalyzer.Analyze(p, false);
        findings.Should().Contain(f => f.Category.Contains("email-based subscriptions") && f.Severity == "INFO");
    }

    [Fact]
    public void SecurityDefaultsEnabled_FlaggedInfo()
    {
        var (summary, findings) = TenantDefaultsAnalyzer.Analyze(Safe(), securityDefaultsEnabled: true);
        summary.SecurityDefaultsEnabled.Should().BeTrue();
        findings.Should().Contain(f => f.Category.Contains("Security Defaults") && f.Severity == "INFO");
    }

    [Fact]
    public void InvitesFromCaseInsensitive()
    {
        var p = Safe() with { AllowInvitesFrom = "EveryOne" };
        var (_, findings) = TenantDefaultsAnalyzer.Analyze(p, false);
        findings.Should().Contain(f => f.Category.Contains("invite guests"));
    }
}
