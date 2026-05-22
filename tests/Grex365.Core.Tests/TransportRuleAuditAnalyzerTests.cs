using FluentAssertions;
using Grex365.Core.Audit;

namespace Grex365.Core.Tests;

public class TransportRuleAuditAnalyzerTests
{
    private static readonly string[] Domains = { "tenant.com", "tenant.onmicrosoft.com" };

    private static TransportRuleSnapshot Make(
        string name = "Rule1",
        string state = "Enabled",
        string mode = "Enforce",
        IReadOnlyList<string>? forwardTo = null,
        IReadOnlyList<string>? bcc = null,
        IReadOnlyList<string>? redirect = null,
        string? routeConnector = null,
        bool deleteMessage = false,
        string? sentToScope = "InOrganization",
        string? fromScope = null) =>
        new(
            Name: name,
            State: state,
            Priority: 0,
            Mode: mode,
            Description: null,
            ForwardTo: forwardTo ?? Array.Empty<string>(),
            BlindCopyTo: bcc ?? Array.Empty<string>(),
            RedirectMessageTo: redirect ?? Array.Empty<string>(),
            RouteMessageOutboundConnector: routeConnector,
            DeleteMessage: deleteMessage,
            SentToScope: sentToScope,
            FromScope: fromScope);

    [Fact]
    public void EmptyInput_NoFindings()
    {
        var (summary, findings) = TransportRuleAuditAnalyzer.Analyze(
            Array.Empty<TransportRuleSnapshot>(), Domains);
        summary.Total.Should().Be(0);
        findings.Should().BeEmpty();
    }

    [Fact]
    public void EnabledRuleNoActions_NotFlagged()
    {
        var r = Make();
        var (summary, findings) = TransportRuleAuditAnalyzer.Analyze(new[] { r }, Domains);
        summary.Enabled.Should().Be(1);
        findings.Should().BeEmpty();
    }

    [Fact]
    public void ForwardToExternal_FlaggedError()
    {
        var r = Make(forwardTo: new[] { "bad@evil.com" });
        var (summary, findings) = TransportRuleAuditAnalyzer.Analyze(new[] { r }, Domains);
        summary.WithExternalForward.Should().Be(1);
        findings.Should().ContainSingle();
        findings[0].Category.Should().Be("Transport rule forwards externally");
        findings[0].Severity.Should().Be("ERROR");
        findings[0].Detail.Should().Contain("bad@evil.com");
    }

    [Fact]
    public void ForwardToInternal_NotFlagged()
    {
        var r = Make(forwardTo: new[] { "user@tenant.com" });
        var (summary, findings) = TransportRuleAuditAnalyzer.Analyze(new[] { r }, Domains);
        summary.WithExternalForward.Should().Be(0);
        findings.Should().BeEmpty();
    }

    [Fact]
    public void BlindCopyToExternal_FlaggedError()
    {
        var r = Make(bcc: new[] { "spy@external.org" });
        var (summary, findings) = TransportRuleAuditAnalyzer.Analyze(new[] { r }, Domains);
        summary.WithExternalBcc.Should().Be(1);
        findings.Should().Contain(f => f.Category == "Transport rule BCC externally" && f.Severity == "ERROR");
    }

    [Fact]
    public void RedirectExternal_FlaggedError()
    {
        var r = Make(redirect: new[] { "x@partner.example" });
        var (summary, findings) = TransportRuleAuditAnalyzer.Analyze(new[] { r }, Domains);
        summary.WithExternalRedirect.Should().Be(1);
        findings.Should().Contain(f => f.Category == "Transport rule redirects externally" && f.Severity == "ERROR");
    }

    [Fact]
    public void OutboundConnectorCustom_FlaggedInfo()
    {
        var r = Make(routeConnector: "PartnerConnector");
        var (_, findings) = TransportRuleAuditAnalyzer.Analyze(new[] { r }, Domains);
        findings.Should().ContainSingle();
        findings[0].Category.Should().Be("Transport rule routes via custom connector");
        findings[0].Severity.Should().Be("INFO");
        findings[0].Detail.Should().Contain("PartnerConnector");
    }

    [Fact]
    public void DeleteWithBroadScope_FlaggedWarn()
    {
        var r = Make(deleteMessage: true, sentToScope: null, fromScope: null);
        var (_, findings) = TransportRuleAuditAnalyzer.Analyze(new[] { r }, Domains);
        findings.Should().Contain(f => f.Category == "Transport rule deletes with broad scope" && f.Severity == "WARN");
    }

    [Fact]
    public void DeleteWithSpecificScope_NotFlagged()
    {
        var r = Make(deleteMessage: true, sentToScope: "InOrganization");
        var (_, findings) = TransportRuleAuditAnalyzer.Analyze(new[] { r }, Domains);
        findings.Should().NotContain(f => f.Category == "Transport rule deletes with broad scope");
    }

    [Fact]
    public void AuditMode_FlaggedInfo()
    {
        var r = Make(mode: "Audit");
        var (_, findings) = TransportRuleAuditAnalyzer.Analyze(new[] { r }, Domains);
        findings.Should().Contain(f => f.Category == "Transport rule in audit mode" && f.Severity == "INFO");
    }

    [Fact]
    public void DisabledRule_FlaggedInfo()
    {
        var r = Make(name: "Some old rule", state: "Disabled");
        var (summary, findings) = TransportRuleAuditAnalyzer.Analyze(new[] { r }, Domains);
        summary.Disabled.Should().Be(1);
        findings.Should().ContainSingle();
        findings[0].Category.Should().Be("Transport rule disabled");
        findings[0].Severity.Should().Be("INFO");
    }

    [Fact]
    public void DisabledSecurityRule_FlaggedWarn()
    {
        var r = Make(name: "Anti-phishing block", state: "Disabled");
        var (_, findings) = TransportRuleAuditAnalyzer.Analyze(new[] { r }, Domains);
        findings.Should().ContainSingle();
        findings[0].Category.Should().Be("Transport rule disabled (security keyword)");
        findings[0].Severity.Should().Be("WARN");
    }

    [Fact]
    public void SmtpPrefix_IsStripped()
    {
        var r = Make(forwardTo: new[] { "smtp:exfil@evil.com" });
        var (_, findings) = TransportRuleAuditAnalyzer.Analyze(new[] { r }, Domains);
        findings.Should().ContainSingle();
        findings[0].Detail.Should().Contain("exfil@evil.com");
    }

    [Fact]
    public void DomainMatchingIsCaseInsensitive()
    {
        var r = Make(forwardTo: new[] { "u@TENANT.COM" });
        var (summary, _) = TransportRuleAuditAnalyzer.Analyze(new[] { r }, Domains);
        summary.WithExternalForward.Should().Be(0);
    }

    [Fact]
    public void RecipientWithoutDomain_Skipped()
    {
        var r = Make(forwardTo: new[] { "internal-alias" });
        var (summary, findings) = TransportRuleAuditAnalyzer.Analyze(new[] { r }, Domains);
        summary.WithExternalForward.Should().Be(0);
        findings.Should().BeEmpty();
    }

    [Fact]
    public void EmptyName_Skipped()
    {
        var r = Make(name: "");
        var (summary, _) = TransportRuleAuditAnalyzer.Analyze(new[] { r }, Domains);
        summary.Total.Should().Be(0);
    }

    [Fact]
    public void MultipleFindingsSameRule_AllEmitted()
    {
        var r = Make(
            forwardTo: new[] { "fwd@evil.com" },
            bcc: new[] { "bcc@evil.com" },
            redirect: new[] { "red@evil.com" },
            routeConnector: "X");
        var (_, findings) = TransportRuleAuditAnalyzer.Analyze(new[] { r }, Domains);
        findings.Should().HaveCount(4);
    }
}
