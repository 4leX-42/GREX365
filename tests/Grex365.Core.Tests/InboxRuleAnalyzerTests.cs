using FluentAssertions;
using Grex365.Core.Audit;

namespace Grex365.Core.Tests;

public class InboxRuleAnalyzerTests
{
    private static readonly string[] Accepted = new[] { "contoso.com" };

    private static InboxRuleRow Make(
        string upn = "u@contoso.com",
        string name = "rule",
        bool enabled = true,
        bool delete = false,
        string? moveTo = null,
        string[]? fwd = null,
        string[]? fwdAtt = null,
        string[]? redir = null,
        string[]? subjectWords = null,
        string[]? bodyWords = null)
        => new(
            MailboxUpn: upn,
            RuleName: name,
            Enabled: enabled,
            DeleteMessage: delete,
            MoveToFolder: moveTo,
            ForwardTo: fwd ?? Array.Empty<string>(),
            ForwardAsAttachmentTo: fwdAtt ?? Array.Empty<string>(),
            RedirectTo: redir ?? Array.Empty<string>(),
            SubjectContainsWords: subjectWords ?? Array.Empty<string>(),
            BodyContainsWords: bodyWords ?? Array.Empty<string>());

    [Fact]
    public void DisabledRule_Skipped()
    {
        var findings = InboxRuleAnalyzer.Analyze(new[] { Make(enabled: false, delete: true) }, Accepted);
        findings.Should().BeEmpty();
    }

    [Fact]
    public void EmptyUpn_Skipped()
    {
        var findings = InboxRuleAnalyzer.Analyze(new[] { Make(upn: string.Empty, delete: true) }, Accepted);
        findings.Should().BeEmpty();
    }

    [Fact]
    public void DeleteMessage_FlaggedInfo()
    {
        var findings = InboxRuleAnalyzer.Analyze(new[] { Make(delete: true) }, Accepted);
        findings.Should().ContainSingle();
        findings[0].Category.Should().Be("Inbox rule: delete");
        findings[0].Severity.Should().Be("INFO");
    }

    [Fact]
    public void DeleteMessage_WithBecKeyword_FlaggedWarn()
    {
        var findings = InboxRuleAnalyzer.Analyze(
            new[] { Make(delete: true, subjectWords: new[] { "invoice" }) },
            Accepted);
        findings.Should().ContainSingle();
        findings[0].Severity.Should().Be("WARN");
        findings[0].Detail.Should().Contain("invoice").And.Contain("BEC");
    }

    [Fact]
    public void MoveToDeletedItems_Flagged()
    {
        var findings = InboxRuleAnalyzer.Analyze(new[] { Make(moveTo: "Deleted Items") }, Accepted);
        findings.Should().ContainSingle();
        findings[0].Category.Should().Be("Inbox rule: hide");
    }

    [Fact]
    public void MoveToRegularFolder_NotFlagged()
    {
        var findings = InboxRuleAnalyzer.Analyze(new[] { Make(moveTo: "Projects") }, Accepted);
        findings.Should().BeEmpty();
    }

    [Fact]
    public void MoveToRss_WithKeyword_Warn()
    {
        var findings = InboxRuleAnalyzer.Analyze(
            new[] { Make(moveTo: "RSS Feeds", bodyWords: new[] { "payment" }) },
            Accepted);
        findings.Should().ContainSingle();
        findings[0].Severity.Should().Be("WARN");
    }

    [Fact]
    public void ForwardToExternal_Flagged()
    {
        var findings = InboxRuleAnalyzer.Analyze(
            new[] { Make(fwd: new[] { "attacker@evil.com" }) },
            Accepted);
        findings.Should().ContainSingle();
        findings[0].Category.Should().Be("Inbox rule: external forward");
        findings[0].Severity.Should().Be("WARN");
        findings[0].Detail.Should().Contain("attacker@evil.com");
    }

    [Fact]
    public void ForwardToInternal_NotFlagged()
    {
        var findings = InboxRuleAnalyzer.Analyze(
            new[] { Make(fwd: new[] { "boss@contoso.com" }) },
            Accepted);
        findings.Should().BeEmpty();
    }

    [Fact]
    public void RedirectToExternal_Flagged()
    {
        var findings = InboxRuleAnalyzer.Analyze(
            new[] { Make(redir: new[] { "exfil@evil.com" }) },
            Accepted);
        findings.Should().ContainSingle();
        findings[0].Category.Should().Be("Inbox rule: external forward");
    }

    [Fact]
    public void SmtpInBrackets_Extracted()
    {
        var findings = InboxRuleAnalyzer.Analyze(
            new[] { Make(fwd: new[] { "John Doe [SMTP:john@evil.com]" }) },
            Accepted);
        findings.Should().ContainSingle();
        findings[0].Detail.Should().Contain("john@evil.com");
    }

    [Fact]
    public void MultipleFindings_PerRule()
    {
        var findings = InboxRuleAnalyzer.Analyze(
            new[]
            {
                Make(delete: true,
                     moveTo: "Deleted Items",
                     fwd: new[] { "attacker@evil.com" },
                     subjectWords: new[] { "wire" })
            },
            Accepted);
        // delete + hide + external forward = 3 findings
        findings.Should().HaveCount(3);
        findings.Select(f => f.Category).Should().BeEquivalentTo(new[]
        {
            "Inbox rule: delete",
            "Inbox rule: hide",
            "Inbox rule: external forward",
        });
    }

    [Fact]
    public void SpanishKeyword_DetectedToo()
    {
        var findings = InboxRuleAnalyzer.Analyze(
            new[] { Make(delete: true, subjectWords: new[] { "factura" }) },
            Accepted);
        findings.Should().ContainSingle();
        findings[0].Severity.Should().Be("WARN");
        findings[0].Detail.Should().Contain("factura");
    }
}
