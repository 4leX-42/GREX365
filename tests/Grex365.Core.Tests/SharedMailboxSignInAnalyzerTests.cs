using FluentAssertions;
using Grex365.Core.Audit;

namespace Grex365.Core.Tests;

public class SharedMailboxSignInAnalyzerTests
{
    [Fact]
    public void EmptyInput_NoFindings()
    {
        var (summary, findings) = SharedMailboxSignInAnalyzer.Analyze(Array.Empty<SharedMailboxSignInRow>());
        summary.Total.Should().Be(0);
        findings.Should().BeEmpty();
    }

    [Fact]
    public void DisabledShared_NotFlagged()
    {
        var row = new SharedMailboxSignInRow("shared@a.com", "Shared", AccountDisabled: true);
        var (summary, findings) = SharedMailboxSignInAnalyzer.Analyze(new[] { row });
        summary.SignInDisabled.Should().Be(1);
        summary.SignInEnabled.Should().Be(0);
        findings.Should().BeEmpty();
    }

    [Fact]
    public void EnabledShared_FlaggedWarn()
    {
        var row = new SharedMailboxSignInRow("vuln@a.com", "Vulnerable", AccountDisabled: false);
        var (summary, findings) = SharedMailboxSignInAnalyzer.Analyze(new[] { row });
        summary.SignInEnabled.Should().Be(1);
        findings.Should().ContainSingle();
        findings[0].Category.Should().Be("Shared mailbox sign-in enabled");
        findings[0].Severity.Should().Be("WARN");
        findings[0].Identity.Should().Be("vuln@a.com");
    }

    [Fact]
    public void UnknownAccountDisabled_FlaggedInfo()
    {
        var row = new SharedMailboxSignInRow("unknown@a.com", "Unknown", AccountDisabled: null);
        var (summary, findings) = SharedMailboxSignInAnalyzer.Analyze(new[] { row });
        summary.Unknown.Should().Be(1);
        findings.Should().ContainSingle();
        findings[0].Category.Should().Be("Shared mailbox sign-in unknown");
        findings[0].Severity.Should().Be("INFO");
    }

    [Fact]
    public void EmptyUpn_Skipped()
    {
        var row = new SharedMailboxSignInRow("", "X", false);
        var (summary, findings) = SharedMailboxSignInAnalyzer.Analyze(new[] { row });
        summary.Total.Should().Be(0);
        findings.Should().BeEmpty();
    }

    [Fact]
    public void MixedPopulation_CountsCorrectly()
    {
        var rows = new[]
        {
            new SharedMailboxSignInRow("ok@a.com", "OK", true),
            new SharedMailboxSignInRow("vuln1@a.com", "V1", false),
            new SharedMailboxSignInRow("vuln2@a.com", "V2", false),
            new SharedMailboxSignInRow("unknown@a.com", "U", null),
        };
        var (summary, findings) = SharedMailboxSignInAnalyzer.Analyze(rows);
        summary.Total.Should().Be(4);
        summary.SignInDisabled.Should().Be(1);
        summary.SignInEnabled.Should().Be(2);
        summary.Unknown.Should().Be(1);
        findings.Should().HaveCount(3);
        findings.Where(f => f.Severity == "WARN").Should().HaveCount(2);
        findings.Where(f => f.Severity == "INFO").Should().HaveCount(1);
    }
}
