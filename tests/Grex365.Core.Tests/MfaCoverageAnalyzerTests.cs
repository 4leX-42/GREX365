using FluentAssertions;
using Grex365.Core.Audit;

namespace Grex365.Core.Tests;

public class MfaCoverageAnalyzerTests
{
    [Fact]
    public void EmptyInput_ZeroEverything()
    {
        var (summary, findings) = MfaCoverageAnalyzer.Analyze(Array.Empty<MfaRegistrationRow>());
        summary.Total.Should().Be(0);
        summary.AdminsTotal.Should().Be(0);
        findings.Should().BeEmpty();
    }

    [Fact]
    public void AdminWithoutMfa_FlaggedError()
    {
        var rows = new[]
        {
            new MfaRegistrationRow("admin@a.com", "Admin", "Member",
                IsAdmin: true, IsMfaRegistered: false, IsMfaCapable: true),
        };
        var (summary, findings) = MfaCoverageAnalyzer.Analyze(rows);
        summary.AdminsTotal.Should().Be(1);
        summary.AdminsWithoutMfa.Should().Be(1);
        findings.Should().ContainSingle();
        findings[0].Category.Should().Be("MFA missing (admin)");
        findings[0].Severity.Should().Be("ERROR");
        findings[0].Detail.Should().Contain("capable=true");
    }

    [Fact]
    public void AdminWithMfa_NotFlagged()
    {
        var rows = new[]
        {
            new MfaRegistrationRow("a@a.com", "A", "Member",
                IsAdmin: true, IsMfaRegistered: true, IsMfaCapable: true),
        };
        var (summary, findings) = MfaCoverageAnalyzer.Analyze(rows);
        summary.AdminsWithoutMfa.Should().Be(0);
        findings.Should().BeEmpty();
    }

    [Fact]
    public void MemberWithoutMfa_FlaggedWarn()
    {
        var rows = new[]
        {
            new MfaRegistrationRow("u@a.com", "U", "Member",
                IsAdmin: false, IsMfaRegistered: false, IsMfaCapable: true),
        };
        var (summary, findings) = MfaCoverageAnalyzer.Analyze(rows);
        summary.MembersTotal.Should().Be(1);
        summary.MembersWithoutMfa.Should().Be(1);
        findings.Should().ContainSingle();
        findings[0].Category.Should().Be("MFA missing (member)");
        findings[0].Severity.Should().Be("WARN");
    }

    [Fact]
    public void GuestWithoutMfa_FlaggedInfo()
    {
        var rows = new[]
        {
            new MfaRegistrationRow("g@partner.com", "G", "Guest",
                IsAdmin: false, IsMfaRegistered: false, IsMfaCapable: true),
        };
        var (summary, findings) = MfaCoverageAnalyzer.Analyze(rows);
        summary.GuestsTotal.Should().Be(1);
        summary.GuestsWithoutMfa.Should().Be(1);
        summary.MembersTotal.Should().Be(0);
        findings.Should().ContainSingle();
        findings[0].Category.Should().Be("MFA missing (guest)");
        findings[0].Severity.Should().Be("INFO");
    }

    [Fact]
    public void EmptyUpn_Skipped()
    {
        var rows = new[]
        {
            new MfaRegistrationRow(string.Empty, "X", "Member",
                IsAdmin: true, IsMfaRegistered: false, IsMfaCapable: false),
        };
        var (summary, findings) = MfaCoverageAnalyzer.Analyze(rows);
        summary.Total.Should().Be(0);
        findings.Should().BeEmpty();
    }

    [Fact]
    public void AdminClassification_TakesPrecedenceOverGuestType()
    {
        // Si IsAdmin==true se cuenta como admin independientemente de UserType
        var rows = new[]
        {
            new MfaRegistrationRow("ext-admin@partner.com", "EA", "Guest",
                IsAdmin: true, IsMfaRegistered: false, IsMfaCapable: true),
        };
        var (summary, findings) = MfaCoverageAnalyzer.Analyze(rows);
        summary.AdminsTotal.Should().Be(1);
        summary.GuestsTotal.Should().Be(0);
        findings.Should().ContainSingle(f => f.Category == "MFA missing (admin)");
    }

    [Fact]
    public void MixedPopulation_CountsCorrectly()
    {
        var rows = new[]
        {
            new MfaRegistrationRow("a1@a.com", "A1", "Member", true, false, true),  // admin sin MFA
            new MfaRegistrationRow("a2@a.com", "A2", "Member", true, true,  true),  // admin OK
            new MfaRegistrationRow("u1@a.com", "U1", "Member", false, false, true), // member sin MFA
            new MfaRegistrationRow("u2@a.com", "U2", "Member", false, true,  true), // member OK
            new MfaRegistrationRow("g1@a.com", "G1", "Guest",  false, false, true), // guest sin MFA
        };
        var (summary, findings) = MfaCoverageAnalyzer.Analyze(rows);
        summary.Total.Should().Be(5);
        summary.AdminsTotal.Should().Be(2);
        summary.AdminsWithoutMfa.Should().Be(1);
        summary.MembersTotal.Should().Be(2);
        summary.MembersWithoutMfa.Should().Be(1);
        summary.GuestsTotal.Should().Be(1);
        summary.GuestsWithoutMfa.Should().Be(1);
        findings.Should().HaveCount(3);
    }

    [Fact]
    public void AdminWithoutCapable_DetailReportsCapableFalse()
    {
        var rows = new[]
        {
            new MfaRegistrationRow("a@a.com", "A", "Member",
                IsAdmin: true, IsMfaRegistered: false, IsMfaCapable: false),
        };
        var (_, findings) = MfaCoverageAnalyzer.Analyze(rows);
        findings[0].Detail.Should().Contain("capable=false");
    }
}
