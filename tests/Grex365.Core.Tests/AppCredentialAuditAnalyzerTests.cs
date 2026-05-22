using FluentAssertions;
using Grex365.Core.Audit;

namespace Grex365.Core.Tests;

public class AppCredentialAuditAnalyzerTests
{
    private static readonly DateTimeOffset Now = new(2026, 5, 22, 12, 0, 0, TimeSpan.Zero);

    private static AppCredentialSnapshot Make(
        string appId = "app1",
        string display = "App 1",
        string type = "Password",
        string? keyId = "k1",
        string? credName = null,
        DateTimeOffset? end = null) =>
        new(appId, display, type, keyId, credName, end);

    [Fact]
    public void EmptyInput_NoFindings()
    {
        var (summary, findings) = AppCredentialAuditAnalyzer.Analyze(
            Array.Empty<AppCredentialSnapshot>(), Now);
        summary.Total.Should().Be(0);
        findings.Should().BeEmpty();
    }

    [Fact]
    public void ExpiredCredential_FlaggedError()
    {
        var c = Make(end: Now.AddDays(-10));
        var (summary, findings) = AppCredentialAuditAnalyzer.Analyze(new[] { c }, Now);
        summary.Expired.Should().Be(1);
        findings.Should().ContainSingle();
        findings[0].Category.Should().Be("App credential expired");
        findings[0].Severity.Should().Be("ERROR");
        findings[0].Detail.Should().Contain("10");
    }

    [Fact]
    public void ExpiringSoon_FlaggedWarn()
    {
        var c = Make(end: Now.AddDays(15));
        var (summary, findings) = AppCredentialAuditAnalyzer.Analyze(new[] { c }, Now);
        summary.ExpiringSoon.Should().Be(1);
        findings.Should().ContainSingle();
        findings[0].Category.Should().Be("App credential expiring");
        findings[0].Severity.Should().Be("WARN");
    }

    [Fact]
    public void NotExpiringSoon_NotFlagged()
    {
        var c = Make(end: Now.AddDays(180));
        var (summary, findings) = AppCredentialAuditAnalyzer.Analyze(new[] { c }, Now);
        summary.ExpiringSoon.Should().Be(0);
        findings.Should().BeEmpty();
    }

    [Fact]
    public void LongLived_FlaggedInfo()
    {
        var c = Make(end: Now.AddDays(800));
        var (summary, findings) = AppCredentialAuditAnalyzer.Analyze(new[] { c }, Now);
        summary.LongLived.Should().Be(1);
        findings.Should().ContainSingle();
        findings[0].Category.Should().Be("App credential long-lived");
        findings[0].Severity.Should().Be("INFO");
    }

    [Fact]
    public void NullEndDate_Skipped()
    {
        var c = Make(end: null);
        var (summary, findings) = AppCredentialAuditAnalyzer.Analyze(new[] { c }, Now);
        summary.Total.Should().Be(1);
        findings.Should().BeEmpty();
    }

    [Fact]
    public void EmptyAppId_Skipped()
    {
        var c = Make(appId: "", end: Now.AddDays(-1));
        var (summary, _) = AppCredentialAuditAnalyzer.Analyze(new[] { c }, Now);
        summary.Total.Should().Be(0);
    }

    [Fact]
    public void DetailIncludesCredentialName_WhenProvided()
    {
        var c = Make(credName: "prod-secret", end: Now.AddDays(5));
        var (_, findings) = AppCredentialAuditAnalyzer.Analyze(new[] { c }, Now);
        findings[0].Detail.Should().Contain("prod-secret");
    }

    [Fact]
    public void MixedCredentials_CountsCorrectly()
    {
        var creds = new[]
        {
            Make(end: Now.AddDays(-5)),  // expired
            Make(end: Now.AddDays(10)),  // expiring
            Make(end: Now.AddDays(60)),  // normal
            Make(end: Now.AddDays(800)), // long-lived
            Make(end: null),             // unknown
        };
        var (summary, findings) = AppCredentialAuditAnalyzer.Analyze(creds, Now);
        summary.Total.Should().Be(5);
        summary.Expired.Should().Be(1);
        summary.ExpiringSoon.Should().Be(1);
        summary.LongLived.Should().Be(1);
        findings.Should().HaveCount(3);
    }

    [Fact]
    public void KeyTypeAlsoFlagged()
    {
        var c = Make(type: "Key", end: Now.AddDays(-1));
        var (_, findings) = AppCredentialAuditAnalyzer.Analyze(new[] { c }, Now);
        findings[0].Detail.Should().Contain("Key");
    }

    [Fact]
    public void ExactlyAtThreshold_30Days_StillFlaggedWarn()
    {
        var c = Make(end: Now.AddDays(AppCredentialAuditAnalyzer.ExpiringSoonDays));
        var (summary, _) = AppCredentialAuditAnalyzer.Analyze(new[] { c }, Now);
        summary.ExpiringSoon.Should().Be(1);
    }
}
