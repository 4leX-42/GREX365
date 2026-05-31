using FluentAssertions;
using Grex365.Core.Models;
using Grex365.Core.Offboarding;

namespace Grex365.Core.Tests;

public class OffboardingReportTests
{
    private static OffboardingResult Sample(params OffboardingStep[] steps) =>
        new("jane@a", true, steps, DryRun: false,
            StartedAt: new DateTimeOffset(2026, 5, 31, 10, 0, 0, TimeSpan.Zero),
            EndedAt: new DateTimeOffset(2026, 5, 31, 10, 1, 0, TimeSpan.Zero));

    [Fact]
    public void ToCsv_HasHeader_AndOneRowPerStep()
    {
        var csv = OffboardingReport.ToCsv(new[]
        {
            Sample(
                new OffboardingStep("Buscar usuario", "OK", "found"),
                new OffboardingStep("Quitar licencias", "OK", "2 retiradas")),
        });

        var lines = csv.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        lines[0].Should().StartWith("upn,success,dryRun,startedAt,endedAt,step,status,detail,at");
        lines.Should().HaveCount(3); // header + 2 steps
        lines[1].Should().Contain("jane@a").And.Contain("Buscar usuario");
        lines[2].Should().Contain("Quitar licencias");
    }

    [Fact]
    public void ToCsv_EscapesCommasAndQuotes()
    {
        var csv = OffboardingReport.ToCsv(new[]
        {
            Sample(new OffboardingStep("Verificaciones previas", "AVISO", "buzón SharedMailbox; 12,5 GB; hold \"activo\"")),
        });

        // The detail with a comma and quotes must be wrapped and quotes doubled.
        csv.Should().Contain("\"buzón SharedMailbox; 12,5 GB; hold \"\"activo\"\"\"");
    }

    [Fact]
    public void ToCsv_RunWithNoSteps_StillEmitsSummaryRow()
    {
        var csv = OffboardingReport.ToCsv(new[] { Sample() });
        var lines = csv.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        lines.Should().HaveCount(2); // header + 1 summary row
        lines[1].Should().StartWith("jane@a,true,false");
    }

    [Fact]
    public void ToJson_IncludesKeyFields()
    {
        var json = OffboardingReport.ToJson(new[]
        {
            new OffboardingResult("bob@a", false, new[] { new OffboardingStep("Quitar licencias", "OMITIDO", "bloqueado") }, DryRun: true),
        });

        json.Should().Contain("\"upn\": \"bob@a\"");
        json.Should().Contain("\"success\": false");
        json.Should().Contain("\"dryRun\": true");
        json.Should().Contain("\"status\": \"OMITIDO\"");
    }
}
