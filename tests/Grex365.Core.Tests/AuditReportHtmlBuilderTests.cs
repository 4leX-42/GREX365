using FluentAssertions;
using Grex365.Core.Audit;
using Grex365.Core.Models;

namespace Grex365.Core.Tests;

public class AuditReportHtmlBuilderTests
{
    private static readonly DateTime FixedTime = new(2026, 5, 23, 14, 30, 0, DateTimeKind.Utc);
    private static AuditReportContext Ctx(string? tenant = null, string? actor = null) =>
        new("Informe", FixedTime, tenant, actor);

    [Fact]
    public void Empty_findings_emits_doctype_and_empty_marker()
    {
        var html = AuditReportHtmlBuilder.Build(Array.Empty<AuditFinding>(), Ctx());
        html.Should().StartWith("<!DOCTYPE html>");
        html.Should().Contain("Sin hallazgos");
        html.Should().Contain("<title>Informe</title>");
    }

    [Fact]
    public void Null_findings_does_not_throw()
    {
        var act = () => AuditReportHtmlBuilder.Build(null!, Ctx());
        act.Should().NotThrow();
    }

    [Fact]
    public void Groups_findings_by_severity()
    {
        var findings = new[]
        {
            new AuditFinding("Cat1", "u1@a", "Crítico", "ERROR"),
            new AuditFinding("Cat2", "u2@a", "Aviso",   "WARN"),
            new AuditFinding("Cat3", "u3@a", "Nota",    "INFO"),
            new AuditFinding("Cat4", "u4@a", "Otro",    "ERROR"),
        };
        var html = AuditReportHtmlBuilder.Build(findings, Ctx());
        html.Should().Contain("Errores").And.Contain("(2)");
        html.Should().Contain("Advertencias").And.Contain("(1)");
        html.Should().Contain("Informativos").And.Contain("(1)");
    }

    [Fact]
    public void Hides_severity_section_when_empty()
    {
        var html = AuditReportHtmlBuilder.Build(new[]
        {
            new AuditFinding("c", "i", "d", "ERROR"),
        }, Ctx());
        html.Should().Contain("Errores");
        html.Should().NotContain("<h2>Advertencias");
        html.Should().NotContain("<h2>Informativos");
    }

    [Theory]
    [InlineData("<script>alert(1)</script>", "&lt;script&gt;alert(1)&lt;/script&gt;")]
    [InlineData("A & B", "A &amp; B")]
    [InlineData("\"quoted\"", "&quot;quoted&quot;")]
    [InlineData("it's", "it&#39;s")]
    public void Escapes_html_entities_in_content(string input, string expected)
    {
        var html = AuditReportHtmlBuilder.Build(new[]
        {
            new AuditFinding(input, "i", "d", "ERROR"),
        }, Ctx());
        html.Should().Contain(expected);
        html.Should().NotContain("<script>alert");
    }

    [Fact]
    public void Pills_show_all_four_counts_including_total()
    {
        var findings = new[]
        {
            new AuditFinding("c", "i", "d", "ERROR"),
            new AuditFinding("c", "i", "d", "ERROR"),
            new AuditFinding("c", "i", "d", "WARN"),
        };
        var html = AuditReportHtmlBuilder.Build(findings, Ctx());
        html.Should().Contain("pill error").And.Contain(">2<");
        html.Should().Contain("pill warn").And.Contain(">1<");
        html.Should().Contain("pill info").And.Contain(">0<");
        html.Should().Contain("pill total").And.Contain(">3<");
    }

    [Fact]
    public void Includes_tenant_and_actor_when_provided()
    {
        var html = AuditReportHtmlBuilder.Build(
            Array.Empty<AuditFinding>(),
            Ctx(tenant: "contoso.onmicrosoft.com", actor: "admin@contoso"));
        html.Should().Contain("contoso.onmicrosoft.com");
        html.Should().Contain("admin@contoso");
        html.Should().Contain("Tenant:");
        html.Should().Contain("Operador:");
    }

    [Fact]
    public void Omits_tenant_when_null_or_whitespace()
    {
        var html = AuditReportHtmlBuilder.Build(
            Array.Empty<AuditFinding>(),
            Ctx(tenant: "  ", actor: null));
        html.Should().NotContain("Tenant:");
        html.Should().NotContain("Operador:");
    }

    [Fact]
    public void Severity_match_is_case_insensitive()
    {
        var findings = new[]
        {
            new AuditFinding("c", "i", "d", "error"),
            new AuditFinding("c", "i", "d", "Warn"),
            new AuditFinding("c", "i", "d", "INFO"),
        };
        var html = AuditReportHtmlBuilder.Build(findings, Ctx());
        html.Should().Contain("Errores").And.Contain("(1)");
        html.Should().Contain("Advertencias");
        html.Should().Contain("Informativos");
    }

    [Fact]
    public void Generated_timestamp_uses_invariant_format()
    {
        var html = AuditReportHtmlBuilder.Build(Array.Empty<AuditFinding>(), Ctx());
        html.Should().Contain("2026-05-23 14:30:00");
    }

    [Fact]
    public void Renders_table_with_three_columns_per_row()
    {
        var html = AuditReportHtmlBuilder.Build(new[]
        {
            new AuditFinding("CAT", "UPN", "MSG", "ERROR"),
        }, Ctx());
        html.Should().Contain("<th>Categoría</th>");
        html.Should().Contain("<th>Identidad</th>");
        html.Should().Contain("<th>Detalle</th>");
        html.Should().Contain("<td>CAT</td>");
        html.Should().Contain("<td>UPN</td>");
        html.Should().Contain("<td>MSG</td>");
    }

    [Fact]
    public void Esc_helper_returns_empty_for_null()
    {
        AuditReportHtmlBuilder.Esc(null).Should().BeEmpty();
        AuditReportHtmlBuilder.Esc("").Should().BeEmpty();
    }

    [Fact]
    public void Esc_helper_passes_through_safe_chars()
    {
        AuditReportHtmlBuilder.Esc("hello world 123").Should().Be("hello world 123");
        AuditReportHtmlBuilder.Esc("acentos: áéíóú ñÑ").Should().Be("acentos: áéíóú ñÑ");
    }

    [Fact]
    public void Unknown_severity_excluded_from_grouped_tables()
    {
        var findings = new[]
        {
            new AuditFinding("c", "i", "d", "DEBUG"),
            new AuditFinding("c", "i", "d", "ERROR"),
        };
        var html = AuditReportHtmlBuilder.Build(findings, Ctx());
        html.Should().Contain("Errores").And.Contain("(1)");
        html.Should().NotContain("DEBUG");
    }

    [Fact]
    public void Embeds_inline_css_block()
    {
        var html = AuditReportHtmlBuilder.Build(Array.Empty<AuditFinding>(), Ctx());
        html.Should().Contain("<style>");
        html.Should().Contain("prefers-color-scheme: dark");
    }
}
