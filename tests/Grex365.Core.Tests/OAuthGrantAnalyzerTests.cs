using FluentAssertions;
using Grex365.Core.Audit;

namespace Grex365.Core.Tests;

public class OAuthGrantAnalyzerTests
{
    private static OAuthGrantSnapshot Make(
        string clientId = "client-1",
        string clientName = "Suspicious App",
        string consentType = "AllPrincipals",
        string? principalId = null,
        params string[] scopes) =>
        new(
            GrantId: Guid.NewGuid().ToString(),
            ClientId: clientId,
            ClientDisplayName: clientName,
            ResourceId: "graph-id",
            ResourceDisplayName: "Microsoft Graph",
            ConsentType: consentType,
            PrincipalId: principalId,
            Scopes: scopes.Length == 0 ? new List<string> { "User.Read" } : scopes.ToList());

    [Fact]
    public void EmptyInput_NoFindings()
    {
        var (summary, findings) = OAuthGrantAnalyzer.Analyze(Array.Empty<OAuthGrantSnapshot>());
        summary.TotalGrants.Should().Be(0);
        findings.Should().BeEmpty();
    }

    [Fact]
    public void LowRiskScopeOnly_NotFlagged()
    {
        var g = Make(scopes: new[] { "User.Read", "openid", "profile" });
        var (_, findings) = OAuthGrantAnalyzer.Analyze(new[] { g });
        findings.Should().BeEmpty();
    }

    [Fact]
    public void TenantWideHighRisk_FlaggedError()
    {
        var g = Make(consentType: "AllPrincipals", scopes: new[] { "Mail.ReadWrite" });
        var (summary, findings) = OAuthGrantAnalyzer.Analyze(new[] { g });
        summary.TenantWideHighRisk.Should().Be(1);
        findings.Should().ContainSingle();
        findings[0].Category.Should().Be("OAuth grant tenant-wide high-risk");
        findings[0].Severity.Should().Be("ERROR");
        findings[0].Detail.Should().Contain("Mail.ReadWrite");
    }

    [Fact]
    public void UserConsentedHighRisk_FlaggedWarn()
    {
        var g = Make(
            consentType: "Principal",
            principalId: "user-42",
            scopes: new[] { "Files.ReadWrite.All" });
        var (summary, findings) = OAuthGrantAnalyzer.Analyze(new[] { g });
        summary.UserConsentedHighRisk.Should().Be(1);
        findings.Should().ContainSingle();
        findings[0].Category.Should().Be("OAuth grant user-consented high-risk");
        findings[0].Severity.Should().Be("WARN");
        findings[0].Detail.Should().Contain("user-42");
        findings[0].Detail.Should().Contain("Files.ReadWrite.All");
    }

    [Fact]
    public void EmptyClientId_Skipped()
    {
        var g = Make(clientId: "", scopes: new[] { "Mail.ReadWrite" });
        var (summary, findings) = OAuthGrantAnalyzer.Analyze(new[] { g });
        summary.TotalGrants.Should().Be(0);
        findings.Should().BeEmpty();
    }

    [Fact]
    public void ScopesAreTrimmed()
    {
        var g = Make(scopes: new[] { "  Mail.ReadWrite  " });
        var (_, findings) = OAuthGrantAnalyzer.Analyze(new[] { g });
        findings.Should().ContainSingle();
    }

    [Fact]
    public void ScopeMatchIsCaseInsensitive()
    {
        var g = Make(scopes: new[] { "mail.readwrite" });
        var (_, findings) = OAuthGrantAnalyzer.Analyze(new[] { g });
        findings.Should().ContainSingle();
    }

    [Fact]
    public void MixedScopes_MentionedInDetail()
    {
        var g = Make(scopes: new[] { "User.Read", "Mail.ReadWrite", "Files.Read.All" });
        var (_, findings) = OAuthGrantAnalyzer.Analyze(new[] { g });
        findings[0].Detail.Should().Contain("Mail.ReadWrite");
        findings[0].Detail.Should().Contain("Files.Read.All");
        findings[0].Detail.Should().NotContain("User.Read,");
    }

    [Fact]
    public void IdentityFallbackToClientId_WhenNameEmpty()
    {
        var g = Make(clientId: "abc-123", clientName: "", scopes: new[] { "Mail.Read" });
        var (_, findings) = OAuthGrantAnalyzer.Analyze(new[] { g });
        findings[0].Identity.Should().Be("abc-123");
    }

    [Fact]
    public void UniqueClients_CountedOnce()
    {
        var grants = new[]
        {
            Make(clientId: "c1", scopes: new[] { "Mail.Read" }),
            Make(clientId: "c1", scopes: new[] { "Files.Read.All" }),
            Make(clientId: "c2", scopes: new[] { "Mail.Read" }),
        };
        var (summary, _) = OAuthGrantAnalyzer.Analyze(grants);
        summary.UniqueClients.Should().Be(2);
        summary.TotalGrants.Should().Be(3);
    }

    [Fact]
    public void FullAccessAsUser_FlaggedAsHighRisk()
    {
        var g = Make(scopes: new[] { "full_access_as_user" });
        var (_, findings) = OAuthGrantAnalyzer.Analyze(new[] { g });
        findings.Should().ContainSingle();
    }

    [Fact]
    public void IsHighRiskScope_PublicHelper_ReturnsExpected()
    {
        OAuthGrantAnalyzer.IsHighRiskScope("Mail.ReadWrite").Should().BeTrue();
        OAuthGrantAnalyzer.IsHighRiskScope("openid").Should().BeFalse();
        OAuthGrantAnalyzer.IsHighRiskScope("").Should().BeFalse();
        OAuthGrantAnalyzer.IsHighRiskScope("  Mail.Read  ").Should().BeTrue();
    }

    [Fact]
    public void MixedGrants_CountsCorrectly()
    {
        var grants = new[]
        {
            Make(clientId: "c1", consentType: "AllPrincipals", scopes: new[] { "Mail.ReadWrite" }),
            Make(clientId: "c2", consentType: "Principal", principalId: "u1", scopes: new[] { "Files.Read.All" }),
            Make(clientId: "c3", consentType: "AllPrincipals", scopes: new[] { "User.Read" }),
        };
        var (summary, findings) = OAuthGrantAnalyzer.Analyze(grants);
        summary.TotalGrants.Should().Be(3);
        summary.TenantWideHighRisk.Should().Be(1);
        summary.UserConsentedHighRisk.Should().Be(1);
        findings.Should().HaveCount(2);
    }
}
