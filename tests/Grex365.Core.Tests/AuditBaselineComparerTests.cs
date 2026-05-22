using FluentAssertions;
using Grex365.Core.Audit;
using Grex365.Core.Models;

namespace Grex365.Core.Tests;

public class AuditBaselineComparerTests
{
    private static AuditFinding F(string cat, string id, string detail, string sev) =>
        new(cat, id, detail, sev);

    [Fact]
    public void Empty_baseline_and_current_returns_empty_diff()
    {
        var diff = AuditBaselineComparer.Compare(Array.Empty<AuditFinding>(), Array.Empty<AuditFinding>());
        diff.New.Should().BeEmpty();
        diff.Resolved.Should().BeEmpty();
        diff.Persistent.Should().BeEmpty();
        diff.HasChanges.Should().BeFalse();
    }

    [Fact]
    public void Null_inputs_do_not_throw()
    {
        var act = () => AuditBaselineComparer.Compare(null!, null!);
        act.Should().NotThrow();
    }

    [Fact]
    public void Finding_in_both_marked_persistent()
    {
        var f = F("c", "u@a", "d", "ERROR");
        var diff = AuditBaselineComparer.Compare(new[] { f }, new[] { f });
        diff.Persistent.Should().ContainSingle().And.Contain(f);
        diff.New.Should().BeEmpty();
        diff.Resolved.Should().BeEmpty();
        diff.HasChanges.Should().BeFalse();
    }

    [Fact]
    public void Finding_only_in_current_is_new()
    {
        var existing = F("c", "u1", "d", "ERROR");
        var added = F("c", "u2", "d", "WARN");
        var diff = AuditBaselineComparer.Compare(new[] { existing }, new[] { existing, added });
        diff.New.Should().ContainSingle().And.Contain(added);
        diff.Persistent.Should().ContainSingle().And.Contain(existing);
        diff.Resolved.Should().BeEmpty();
        diff.HasChanges.Should().BeTrue();
    }

    [Fact]
    public void Finding_only_in_baseline_is_resolved()
    {
        var existing = F("c", "u1", "d", "ERROR");
        var removed = F("c", "u2", "d", "WARN");
        var diff = AuditBaselineComparer.Compare(new[] { existing, removed }, new[] { existing });
        diff.Resolved.Should().ContainSingle().And.Contain(removed);
        diff.Persistent.Should().ContainSingle().And.Contain(existing);
        diff.New.Should().BeEmpty();
        diff.HasChanges.Should().BeTrue();
    }

    [Fact]
    public void Severity_match_is_case_insensitive()
    {
        var diff = AuditBaselineComparer.Compare(
            new[] { F("c", "u", "d", "error") },
            new[] { F("c", "u", "d", "ERROR") });
        diff.Persistent.Should().ContainSingle();
        diff.New.Should().BeEmpty();
    }

    [Fact]
    public void Category_match_is_case_sensitive()
    {
        var diff = AuditBaselineComparer.Compare(
            new[] { F("Cat", "u", "d", "ERROR") },
            new[] { F("cat", "u", "d", "ERROR") });
        diff.New.Should().HaveCount(1);
        diff.Resolved.Should().HaveCount(1);
    }

    [Fact]
    public void Duplicate_current_findings_deduplicated()
    {
        var f = F("c", "u", "d", "ERROR");
        var diff = AuditBaselineComparer.Compare(Array.Empty<AuditFinding>(), new[] { f, f, f });
        diff.New.Should().ContainSingle();
    }

    [Fact]
    public void Mixed_state_correct_partition()
    {
        var common = F("c", "common", "d", "WARN");
        var added = F("c", "added", "d", "ERROR");
        var resolved = F("c", "resolved", "d", "INFO");
        var diff = AuditBaselineComparer.Compare(
            new[] { common, resolved },
            new[] { common, added });
        diff.Persistent.Should().ContainSingle().And.Contain(common);
        diff.New.Should().ContainSingle().And.Contain(added);
        diff.Resolved.Should().ContainSingle().And.Contain(resolved);
        diff.HasChanges.Should().BeTrue();
    }

    [Fact]
    public void Counts_helpers_reflect_collections()
    {
        var diff = AuditBaselineComparer.Compare(
            new[] { F("c", "r", "d", "ERROR") },
            new[] { F("c", "n", "d", "ERROR") });
        diff.NewCount.Should().Be(1);
        diff.ResolvedCount.Should().Be(1);
        diff.PersistentCount.Should().Be(0);
    }

    [Fact]
    public void Null_severity_treated_as_empty()
    {
        var a = new AuditFinding("c", "u", "d", null!);
        var b = new AuditFinding("c", "u", "d", null!);
        var diff = AuditBaselineComparer.Compare(new[] { a }, new[] { b });
        diff.Persistent.Should().ContainSingle();
    }

    [Fact]
    public void Different_detail_treated_as_different_finding()
    {
        var diff = AuditBaselineComparer.Compare(
            new[] { F("c", "u", "old", "ERROR") },
            new[] { F("c", "u", "new", "ERROR") });
        diff.Resolved.Should().HaveCount(1);
        diff.New.Should().HaveCount(1);
        diff.Persistent.Should().BeEmpty();
    }
}
