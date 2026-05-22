using System.Text.Json;
using FluentAssertions;
using Grex365.Core.Audit;
using Grex365.Core.Models;

namespace Grex365.Core.Tests;

public class AuditReportJsonBuilderTests
{
    private static readonly DateTime FixedTime = new(2026, 5, 23, 14, 30, 0, DateTimeKind.Utc);
    private static AuditReportContext Ctx(string? tenant = null, string? actor = null) =>
        new("Informe", FixedTime, tenant, actor);

    [Fact]
    public void Empty_findings_emits_zero_counts()
    {
        var json = AuditReportJsonBuilder.Build(Array.Empty<AuditFinding>(), Ctx());
        var parsed = AuditReportJsonBuilder.Parse(json);
        parsed.Should().NotBeNull();
        parsed!.Findings.Should().BeEmpty();
        parsed.Counts.Should().Be(new AuditReportCounts(0, 0, 0, 0));
    }

    [Fact]
    public void Null_findings_does_not_throw()
    {
        var act = () => AuditReportJsonBuilder.Build(null!, Ctx());
        act.Should().NotThrow();
    }

    [Fact]
    public void Counts_breakdown_by_severity()
    {
        var findings = new[]
        {
            new AuditFinding("c", "i1", "d", "ERROR"),
            new AuditFinding("c", "i2", "d", "ERROR"),
            new AuditFinding("c", "i3", "d", "WARN"),
            new AuditFinding("c", "i4", "d", "INFO"),
            new AuditFinding("c", "i5", "d", "INFO"),
        };
        var json = AuditReportJsonBuilder.Build(findings, Ctx());
        var parsed = AuditReportJsonBuilder.Parse(json);
        parsed!.Counts.Should().Be(new AuditReportCounts(2, 1, 2, 5));
    }

    [Fact]
    public void Severity_match_case_insensitive()
    {
        var findings = new[]
        {
            new AuditFinding("c", "i", "d", "error"),
            new AuditFinding("c", "i", "d", "Warn"),
            new AuditFinding("c", "i", "d", "INFO"),
        };
        var json = AuditReportJsonBuilder.Build(findings, Ctx());
        var parsed = AuditReportJsonBuilder.Parse(json);
        parsed!.Counts.Should().Be(new AuditReportCounts(1, 1, 1, 3));
    }

    [Fact]
    public void Schema_version_embedded()
    {
        var json = AuditReportJsonBuilder.Build(Array.Empty<AuditFinding>(), Ctx());
        json.Should().Contain("\"Schema\":");
        json.Should().Contain(AuditReportJsonBuilder.SchemaVersion);
    }

    [Fact]
    public void Roundtrip_preserves_findings()
    {
        var findings = new[]
        {
            new AuditFinding("Cat A", "u@a", "Detail < & >", "ERROR"),
            new AuditFinding("Cat B", "u@b", "Detail \"quoted\"", "WARN"),
        };
        var json = AuditReportJsonBuilder.Build(findings, Ctx("contoso.com", "admin"));
        var parsed = AuditReportJsonBuilder.Parse(json);
        parsed!.Findings.Should().BeEquivalentTo(findings);
        parsed.TenantDomain.Should().Be("contoso.com");
        parsed.GeneratedBy.Should().Be("admin");
    }

    [Fact]
    public void GeneratedAt_serialized_iso8601()
    {
        var json = AuditReportJsonBuilder.Build(Array.Empty<AuditFinding>(), Ctx());
        json.Should().Contain("2026-05-23T14:30:00");
    }

    [Fact]
    public void Output_pretty_printed_by_default()
    {
        var json = AuditReportJsonBuilder.Build(Array.Empty<AuditFinding>(), Ctx());
        json.Should().Contain("\n");
        json.Should().Contain("  ");
    }

    [Fact]
    public void Custom_options_respected()
    {
        var compact = new JsonSerializerOptions { WriteIndented = false };
        var json = AuditReportJsonBuilder.Build(Array.Empty<AuditFinding>(), Ctx(), compact);
        json.Should().NotContain("\n  ");
    }

    [Fact]
    public void Parse_returns_null_for_empty_input()
    {
        AuditReportJsonBuilder.Parse(string.Empty).Should().BeNull();
        AuditReportJsonBuilder.Parse("   ").Should().BeNull();
    }

    [Fact]
    public void Parse_throws_on_invalid_json()
    {
        var act = () => AuditReportJsonBuilder.Parse("{not json}");
        act.Should().Throw<JsonException>();
    }

    [Fact]
    public void Tenant_and_actor_serialized_when_provided()
    {
        var json = AuditReportJsonBuilder.Build(Array.Empty<AuditFinding>(), Ctx("contoso.onmicrosoft.com", "admin@x"));
        json.Should().Contain("contoso.onmicrosoft.com");
        json.Should().Contain("admin@x");
    }

    [Fact]
    public void Unknown_severity_excluded_from_counts_but_included_in_findings()
    {
        var findings = new[]
        {
            new AuditFinding("c", "i", "d", "DEBUG"),
            new AuditFinding("c", "i", "d", "ERROR"),
        };
        var json = AuditReportJsonBuilder.Build(findings, Ctx());
        var parsed = AuditReportJsonBuilder.Parse(json);
        parsed!.Counts.Total.Should().Be(2);
        parsed.Counts.Errors.Should().Be(1);
        parsed.Counts.Warnings.Should().Be(0);
        parsed.Counts.Info.Should().Be(0);
        parsed.Findings.Should().HaveCount(2);
    }
}
