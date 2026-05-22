using FluentAssertions;
using Grex365.Core.Audit;

namespace Grex365.Core.Tests;

public class CaPolicyAnalyzerTests
{
    private static readonly DateTimeOffset Now = new(2026, 5, 22, 12, 0, 0, TimeSpan.Zero);

    private static CaPolicySnapshot Make(
        string name = "P1",
        string state = "enabled",
        DateTimeOffset? modified = null,
        IReadOnlyList<string>? includeUsers = null,
        IReadOnlyList<string>? excludeUsers = null,
        IReadOnlyList<string>? includeGroups = null,
        IReadOnlyList<string>? includeRoles = null,
        IReadOnlyList<string>? builtInControls = null) =>
        new(
            Id: Guid.NewGuid().ToString(),
            DisplayName: name,
            State: state,
            CreatedDateTime: modified ?? Now.AddDays(-1),
            ModifiedDateTime: modified ?? Now.AddDays(-1),
            IncludeUsers: includeUsers ?? new List<string> { "All" },
            ExcludeUsers: excludeUsers ?? new List<string>(),
            IncludeGroups: includeGroups ?? new List<string>(),
            ExcludeGroups: new List<string>(),
            IncludeRoles: includeRoles ?? new List<string>(),
            ExcludeRoles: new List<string>(),
            IncludeApplications: new List<string> { "All" },
            BuiltInControls: builtInControls ?? new List<string> { "mfa" },
            GrantOperator: "OR");

    [Fact]
    public void EmptyInput_FlaggedAsCriticalError()
    {
        var (summary, findings) = CaPolicyAnalyzer.Analyze(Array.Empty<CaPolicySnapshot>(), Now);
        summary.Total.Should().Be(0);
        findings.Should().ContainSingle();
        findings[0].Category.Should().Be("No CA policies");
        findings[0].Severity.Should().Be("ERROR");
    }

    [Fact]
    public void EnabledWithMfaAndExclusions_NotFlagged()
    {
        var policy = Make(
            state: "enabled",
            includeUsers: new List<string> { "All" },
            excludeUsers: new List<string> { "break-glass-id" },
            builtInControls: new List<string> { "mfa" });
        var (summary, findings) = CaPolicyAnalyzer.Analyze(new[] { policy }, Now);
        summary.Enabled.Should().Be(1);
        findings.Should().BeEmpty();
    }

    [Fact]
    public void DisabledPolicy_FlaggedInfo()
    {
        var policy = Make(state: "disabled");
        var (summary, findings) = CaPolicyAnalyzer.Analyze(new[] { policy }, Now);
        summary.Disabled.Should().Be(1);
        summary.Enabled.Should().Be(0);
        findings.Should().ContainSingle();
        findings[0].Category.Should().Be("CA policy disabled");
        findings[0].Severity.Should().Be("INFO");
    }

    [Fact]
    public void ReportOnlyStale_FlaggedWarn()
    {
        var policy = Make(
            state: "enabledForReportingButNotEnforced",
            modified: Now.AddDays(-45));
        var (summary, findings) = CaPolicyAnalyzer.Analyze(new[] { policy }, Now);
        summary.ReportOnly.Should().Be(1);
        findings.Should().ContainSingle();
        findings[0].Category.Should().Be("CA policy stale report-only");
        findings[0].Severity.Should().Be("WARN");
        findings[0].Detail.Should().Contain("45");
    }

    [Fact]
    public void ReportOnlyFresh_FlaggedInfo()
    {
        var policy = Make(
            state: "enabledForReportingButNotEnforced",
            modified: Now.AddDays(-5));
        var (_, findings) = CaPolicyAnalyzer.Analyze(new[] { policy }, Now);
        findings.Should().ContainSingle();
        findings[0].Category.Should().Be("CA policy report-only");
        findings[0].Severity.Should().Be("INFO");
    }

    [Fact]
    public void EnabledWithoutControls_FlaggedError()
    {
        var policy = Make(builtInControls: new List<string>());
        var (summary, findings) = CaPolicyAnalyzer.Analyze(new[] { policy }, Now);
        summary.WithoutEffectiveControls.Should().Be(1);
        findings.Should().Contain(f => f.Category == "CA policy without controls" && f.Severity == "ERROR");
    }

    [Fact]
    public void EnabledWithOnlyWeakControls_FlaggedWarn()
    {
        var policy = Make(builtInControls: new List<string> { "approvedApplication" });
        var (_, findings) = CaPolicyAnalyzer.Analyze(new[] { policy }, Now);
        findings.Should().Contain(f => f.Category == "CA policy weak controls" && f.Severity == "WARN");
    }

    [Fact]
    public void EnabledAllUsersNoExclusions_FlaggedWarn()
    {
        var policy = Make(
            includeUsers: new List<string> { "All" },
            excludeUsers: new List<string>(),
            builtInControls: new List<string> { "mfa" });
        var (_, findings) = CaPolicyAnalyzer.Analyze(new[] { policy }, Now);
        findings.Should().Contain(f =>
            f.Category == "CA policy targets All without exclusions" && f.Severity == "WARN");
    }

    [Fact]
    public void EnabledAllUsersWithExclusionGroup_NotFlagged()
    {
        var policy = Make(
            includeUsers: new List<string> { "All" },
            excludeUsers: new List<string>(),
            builtInControls: new List<string> { "mfa" }) with
            { ExcludeGroups = new List<string> { "break-glass-group" } };
        var (_, findings) = CaPolicyAnalyzer.Analyze(new[] { policy }, Now);
        findings.Should().BeEmpty();
    }

    [Fact]
    public void EnabledNoUserScope_FlaggedError()
    {
        var policy = Make(
            includeUsers: new List<string>(),
            includeGroups: new List<string>(),
            includeRoles: new List<string>(),
            builtInControls: new List<string> { "mfa" });
        var (_, findings) = CaPolicyAnalyzer.Analyze(new[] { policy }, Now);
        findings.Should().Contain(f =>
            f.Category == "CA policy without user scope" && f.Severity == "ERROR");
    }

    [Fact]
    public void EmptyDisplayName_Skipped()
    {
        var policy = Make() with { DisplayName = "" };
        var (summary, findings) = CaPolicyAnalyzer.Analyze(new[] { policy }, Now);
        summary.Total.Should().Be(0);
        findings.Should().ContainSingle(f => f.Category == "No CA policies");
    }

    [Fact]
    public void UnknownState_FlaggedWarn()
    {
        var policy = Make(state: "futureState");
        var (_, findings) = CaPolicyAnalyzer.Analyze(new[] { policy }, Now);
        findings.Should().ContainSingle();
        findings[0].Category.Should().Be("CA policy unknown state");
        findings[0].Severity.Should().Be("WARN");
    }

    [Fact]
    public void CompliantDeviceCountsAsStrongControl()
    {
        var policy = Make(builtInControls: new List<string> { "compliantDevice" });
        var (_, findings) = CaPolicyAnalyzer.Analyze(new[] { policy }, Now);
        findings.Should().NotContain(f => f.Category == "CA policy weak controls");
    }

    [Fact]
    public void BlockControlCountsAsStrong()
    {
        var policy = Make(builtInControls: new List<string> { "block" });
        var (_, findings) = CaPolicyAnalyzer.Analyze(new[] { policy }, Now);
        findings.Should().NotContain(f => f.Category == "CA policy weak controls");
    }

    [Fact]
    public void BuiltInControlsCaseInsensitive()
    {
        var policy = Make(
            excludeUsers: new List<string> { "bg" },
            builtInControls: new List<string> { "MFA" });
        var (_, findings) = CaPolicyAnalyzer.Analyze(new[] { policy }, Now);
        findings.Should().BeEmpty();
    }

    [Fact]
    public void MixedPolicies_CountsCorrectly()
    {
        var policies = new[]
        {
            Make("Strong", state: "enabled", excludeUsers: new List<string>{"bg"}),
            Make("Disabled1", state: "disabled"),
            Make("ReportOld", state: "enabledForReportingButNotEnforced", modified: Now.AddDays(-60)),
            Make("ReportNew", state: "enabledForReportingButNotEnforced", modified: Now.AddDays(-1)),
            Make("NoControls", builtInControls: new List<string>(), excludeUsers: new List<string>{"bg"}),
        };
        var (summary, findings) = CaPolicyAnalyzer.Analyze(policies, Now);
        summary.Total.Should().Be(5);
        summary.Enabled.Should().Be(2); // Strong + NoControls (still enabled state)
        summary.Disabled.Should().Be(1);
        summary.ReportOnly.Should().Be(2);
        summary.WithoutEffectiveControls.Should().Be(1);
        findings.Should().Contain(f => f.Category == "CA policy disabled");
        findings.Should().Contain(f => f.Category == "CA policy stale report-only");
        findings.Should().Contain(f => f.Category == "CA policy report-only");
        findings.Should().Contain(f => f.Category == "CA policy without controls");
    }
}
