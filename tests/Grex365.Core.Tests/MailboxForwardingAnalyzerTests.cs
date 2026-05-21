using FluentAssertions;
using Grex365.Core.Audit;

namespace Grex365.Core.Tests;

public class MailboxForwardingAnalyzerTests
{
    private static readonly string[] Accepted =
        new[] { "contoso.com", "contoso.onmicrosoft.com" };

    [Fact]
    public void NoForwarding_NoFindings()
    {
        var rows = new[]
        {
            new MailboxForwardingRow("a@contoso.com", null, null),
        };
        MailboxForwardingAnalyzer.Analyze(rows, Accepted).Should().BeEmpty();
    }

    [Fact]
    public void ForwardingToAcceptedDomain_NoFinding()
    {
        var rows = new[]
        {
            new MailboxForwardingRow("a@contoso.com", "shared@contoso.com", null),
        };
        MailboxForwardingAnalyzer.Analyze(rows, Accepted).Should().BeEmpty();
    }

    [Fact]
    public void ForwardingToExternal_Flagged()
    {
        var rows = new[]
        {
            new MailboxForwardingRow("vict@contoso.com", "attacker@evil.com", null),
        };
        var findings = MailboxForwardingAnalyzer.Analyze(rows, Accepted);
        findings.Should().ContainSingle();
        findings[0].Category.Should().Be("External forwarding (SMTP)");
        findings[0].Identity.Should().Be("vict@contoso.com");
        findings[0].Severity.Should().Be("WARN");
        findings[0].Detail.Should().Contain("evil.com");
    }

    [Fact]
    public void SmtpPrefix_Stripped()
    {
        var rows = new[]
        {
            new MailboxForwardingRow("v@contoso.com", "SMTP:attacker@evil.com", null),
        };
        var findings = MailboxForwardingAnalyzer.Analyze(rows, Accepted);
        findings.Should().ContainSingle();
        findings[0].Detail.Should().Contain("attacker@evil.com");
    }

    [Fact]
    public void AcceptedDomainsCaseInsensitive()
    {
        var rows = new[]
        {
            new MailboxForwardingRow("a@contoso.com", "x@CONTOSO.com", null),
        };
        MailboxForwardingAnalyzer.Analyze(rows, Accepted).Should().BeEmpty();
    }

    [Fact]
    public void EmptyUpn_Skipped()
    {
        var rows = new[]
        {
            new MailboxForwardingRow(string.Empty, "x@evil.com", null),
        };
        MailboxForwardingAnalyzer.Analyze(rows, Accepted).Should().BeEmpty();
    }

    [Fact]
    public void MalformedSmtp_NoDomain_NotFlagged()
    {
        var rows = new[]
        {
            new MailboxForwardingRow("a@contoso.com", "notanemail", null),
        };
        MailboxForwardingAnalyzer.Analyze(rows, Accepted).Should().BeEmpty();
    }

    [Fact]
    public void ForwardingAddress_NotConsidered_OnlySmtpFlagged()
    {
        // ForwardingAddress (recipient-based) suele ser interno; spec actual sólo flagea SMTP.
        var rows = new[]
        {
            new MailboxForwardingRow("a@contoso.com", null, "external@evil.com"),
        };
        MailboxForwardingAnalyzer.Analyze(rows, Accepted).Should().BeEmpty();
    }

    [Fact]
    public void TrailingDotInDomain_Normalized()
    {
        var rows = new[]
        {
            new MailboxForwardingRow("a@contoso.com", "x@evil.com.", null),
        };
        var findings = MailboxForwardingAnalyzer.Analyze(rows, Accepted);
        findings.Should().ContainSingle();
        findings[0].Detail.Should().Contain("evil.com");
    }

    [Fact]
    public void MultipleRows_OnlyExternalsFlagged()
    {
        var rows = new[]
        {
            new MailboxForwardingRow("u1@contoso.com", "alias@contoso.com", null),
            new MailboxForwardingRow("u2@contoso.com", "boss@partner.com", null),
            new MailboxForwardingRow("u3@contoso.com", null, null),
            new MailboxForwardingRow("u4@contoso.com", "another@evil.com", null),
        };
        var findings = MailboxForwardingAnalyzer.Analyze(rows, Accepted);
        findings.Should().HaveCount(2);
        findings.Select(f => f.Identity).Should().BeEquivalentTo(new[] { "u2@contoso.com", "u4@contoso.com" });
    }
}
