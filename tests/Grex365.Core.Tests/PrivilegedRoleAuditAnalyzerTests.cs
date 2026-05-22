using FluentAssertions;
using Grex365.Core.Audit;

namespace Grex365.Core.Tests;

public class PrivilegedRoleAuditAnalyzerTests
{
    private const string GaTemplate = PrivilegedRoleAuditAnalyzer.GlobalAdministratorTemplateId;
    private const string OtherTemplate = "11111111-2222-3333-4444-555555555555";

    private static PrivilegedRoleAssignment Make(
        string memberId,
        string roleName = "Global Administrator",
        string? template = GaTemplate,
        string? upn = null,
        string? userType = "Member",
        string? memberType = "User",
        bool enabled = true,
        string? display = null) =>
        new(
            RoleName: roleName,
            RoleTemplateId: template,
            MemberId: memberId,
            MemberDisplayName: display ?? memberId,
            MemberUpn: upn ?? $"{memberId}@a.com",
            MemberType: memberType,
            MemberUserType: userType,
            MemberAccountEnabled: enabled);

    [Fact]
    public void EmptyInput_FlaggedNoGA()
    {
        var (summary, findings) = PrivilegedRoleAuditAnalyzer.Analyze(Array.Empty<PrivilegedRoleAssignment>());
        summary.GlobalAdmins.Should().Be(0);
        summary.TotalAssignments.Should().Be(0);
        findings.Should().ContainSingle(f => f.Category == "No Global Administrators" && f.Severity == "ERROR");
    }

    [Fact]
    public void SingleGA_WarnsAboutBackup()
    {
        var (summary, findings) = PrivilegedRoleAuditAnalyzer.Analyze(new[] { Make("u1") });
        summary.GlobalAdmins.Should().Be(1);
        findings.Should().ContainSingle(f => f.Category == "Single Global Administrator" && f.Severity == "WARN");
    }

    [Fact]
    public void TwoGAs_NoTenantLevelWarn()
    {
        var (summary, findings) = PrivilegedRoleAuditAnalyzer.Analyze(new[] { Make("u1"), Make("u2") });
        summary.GlobalAdmins.Should().Be(2);
        findings.Should().BeEmpty();
    }

    [Fact]
    public void TooManyGAs_FlaggedSprawl()
    {
        var rs = Enumerable.Range(1, 7).Select(i => Make($"u{i}")).ToArray();
        var (summary, findings) = PrivilegedRoleAuditAnalyzer.Analyze(rs);
        summary.GlobalAdmins.Should().Be(7);
        findings.Should().ContainSingle(f => f.Category == "Too many Global Administrators" && f.Severity == "WARN");
    }

    [Fact]
    public void GuestWithAdminRole_FlaggedError()
    {
        var rs = new[]
        {
            Make("u1"),
            Make("u2"),
            Make("g1", userType: "Guest"),
        };
        var (summary, findings) = PrivilegedRoleAuditAnalyzer.Analyze(rs);
        summary.GuestsWithAdminRole.Should().Be(1);
        findings.Should().Contain(f => f.Category == "Guest with admin role" && f.Severity == "ERROR");
    }

    [Fact]
    public void DisabledAccountWithAdmin_FlaggedError()
    {
        var rs = new[]
        {
            Make("u1"),
            Make("u2"),
            Make("u3", enabled: false),
        };
        var (summary, findings) = PrivilegedRoleAuditAnalyzer.Analyze(rs);
        summary.DisabledWithAdminRole.Should().Be(1);
        findings.Should().Contain(f => f.Category == "Disabled account with admin role" && f.Severity == "ERROR");
    }

    [Fact]
    public void ServicePrincipalWithAdmin_FlaggedInfo()
    {
        var rs = new[]
        {
            Make("u1"),
            Make("u2"),
            Make("sp1", memberType: "ServicePrincipal", enabled: true),
        };
        var (summary, findings) = PrivilegedRoleAuditAnalyzer.Analyze(rs);
        summary.ServicePrincipalsWithAdminRole.Should().Be(1);
        findings.Should().Contain(f => f.Category == "Service principal with admin role" && f.Severity == "INFO");
    }

    [Fact]
    public void DisabledServicePrincipal_DoesNotEmitDisabledFinding()
    {
        // SPs frequently have AccountEnabled=false in our snapshot but we only flag SP-as-admin, not "disabled".
        var rs = new[]
        {
            Make("u1"),
            Make("u2"),
            Make("sp1", memberType: "ServicePrincipal", enabled: false),
        };
        var (_, findings) = PrivilegedRoleAuditAnalyzer.Analyze(rs);
        findings.Should().NotContain(f => f.Category == "Disabled account with admin role");
        findings.Should().Contain(f => f.Category == "Service principal with admin role");
    }

    [Fact]
    public void SameUserMultipleRoles_CountedOncePerRole()
    {
        var rs = new[]
        {
            Make("u1", roleName: "Global Administrator", template: GaTemplate),
            Make("u1", roleName: "Security Administrator", template: OtherTemplate),
            Make("u2"),
        };
        var (summary, _) = PrivilegedRoleAuditAnalyzer.Analyze(rs);
        summary.TotalAssignments.Should().Be(3);
        summary.UniqueAdmins.Should().Be(2);
        summary.GlobalAdmins.Should().Be(2);
    }

    [Fact]
    public void GuestInTwoRoles_FlaggedTwice()
    {
        var rs = new[]
        {
            Make("u1"),
            Make("u2"),
            Make("g1", userType: "Guest", roleName: "Global Administrator", template: GaTemplate),
            Make("g1", userType: "Guest", roleName: "Security Administrator", template: OtherTemplate),
        };
        var (summary, findings) = PrivilegedRoleAuditAnalyzer.Analyze(rs);
        summary.GuestsWithAdminRole.Should().Be(2);
        findings.Where(f => f.Category == "Guest with admin role").Should().HaveCount(2);
    }

    [Fact]
    public void EmptyMemberId_Skipped()
    {
        var rs = new[]
        {
            Make("u1"),
            Make("u2"),
            new PrivilegedRoleAssignment("Global Administrator", GaTemplate,
                MemberId: "",
                MemberDisplayName: "x",
                MemberUpn: "x@a.com",
                MemberType: "User",
                MemberUserType: "Member",
                MemberAccountEnabled: true),
        };
        var (summary, _) = PrivilegedRoleAuditAnalyzer.Analyze(rs);
        summary.TotalAssignments.Should().Be(2);
    }

    [Fact]
    public void NonGaRoleOnly_StillFlagsZeroGAs()
    {
        var rs = new[]
        {
            Make("u1", roleName: "Security Administrator", template: OtherTemplate),
        };
        var (summary, findings) = PrivilegedRoleAuditAnalyzer.Analyze(rs);
        summary.GlobalAdmins.Should().Be(0);
        summary.UniqueAdmins.Should().Be(1);
        findings.Should().Contain(f => f.Category == "No Global Administrators");
    }

    [Fact]
    public void RoleTemplateMatching_IsCaseInsensitive()
    {
        var rs = new[]
        {
            Make("u1", template: GaTemplate.ToUpperInvariant()),
            Make("u2"),
        };
        var (summary, _) = PrivilegedRoleAuditAnalyzer.Analyze(rs);
        summary.GlobalAdmins.Should().Be(2);
    }

    [Fact]
    public void IdentityDerivedFromDisplayName_WhenUpnNull()
    {
        var rs = new[]
        {
            Make("u1"),
            Make("u2"),
            Make("g1", userType: "Guest", upn: null, display: "Guest McGuest"),
        };
        var (_, findings) = PrivilegedRoleAuditAnalyzer.Analyze(rs);
        // Since we still default upn to "g1@a.com" via Make's null-coalesce, override:
        // Use explicit construction:
        rs = rs.Take(2).Concat(new[]
        {
            new PrivilegedRoleAssignment(
                RoleName: "Global Administrator",
                RoleTemplateId: GaTemplate,
                MemberId: "g1",
                MemberDisplayName: "Guest McGuest",
                MemberUpn: null,
                MemberType: "User",
                MemberUserType: "Guest",
                MemberAccountEnabled: true),
        }).ToArray();
        (_, findings) = PrivilegedRoleAuditAnalyzer.Analyze(rs);
        findings.Should().Contain(f => f.Identity == "Guest McGuest" && f.Category == "Guest with admin role");
    }
}
