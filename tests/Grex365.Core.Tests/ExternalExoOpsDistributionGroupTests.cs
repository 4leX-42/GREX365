using FluentAssertions;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Grex365.PowerShell;
using Moq;

namespace Grex365.Core.Tests;

// RemoveFromDistributionGroupsAsync: body generation (injection-safe literals, single
// invocation over N groups) + defensive JSON parsing (array and 1-element collapse).
public class ExternalExoOpsDistributionGroupTests
{
    private static (ExternalExoOps Ops, List<string> Bodies, Mock<IExternalExoRunner> Runner) Build(string? json)
    {
        var runner = new Mock<IExternalExoRunner>();
        var bodies = new List<string>();
        runner.Setup(r => r.RunAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
            .Returns<string, IProgress<LogEntry>?, CancellationToken>((body, _, _) =>
            {
                bodies.Add(body);
                return Task.FromResult(json);
            });
        return (new ExternalExoOps(runner.Object), bodies, runner);
    }

    [Fact]
    public async Task EmptyGroups_ShortCircuits_NoPwsh()
    {
        var (ops, _, runner) = Build(null);

        var results = await ops.RemoveFromDistributionGroupsAsync("u@a", Array.Empty<string>());

        results.Should().BeEmpty();
        runner.Verify(r => r.RunAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task Body_SingleInvocation_AllGroups_EscapedLiterals()
    {
        var (ops, bodies, _) = Build("""[{"Group":"a","Success":true,"Detail":"quitado"}]""");

        await ops.RemoveFromDistributionGroupsAsync("o'brien@a", new[] { "ventas@a", "dl o'hara" });

        bodies.Should().HaveCount(1); // connect once, iterate inside
        var body = bodies[0];
        body.Should().Contain("Remove-DistributionGroupMember");
        body.Should().Contain("-BypassSecurityGroupManagerCheck");
        body.Should().Contain("'o''brien@a'");           // member quote doubled
        body.Should().Contain("'ventas@a', 'dl o''hara'"); // group quotes doubled
    }

    [Fact]
    public async Task Parse_Array_MapsSuccessAndFailure()
    {
        var (ops, _, _) = Build("""[{"Group":"a@x","Success":true,"Detail":"quitado"},{"Group":"b@x","Success":false,"Detail":"propietario"}]""");

        var results = await ops.RemoveFromDistributionGroupsAsync("u@a", new[] { "a@x", "b@x" });

        results.Should().HaveCount(2);
        results[0].Should().Be(new DistributionGroupRemovalResult("a@x", true, "quitado"));
        results[1].Should().Be(new DistributionGroupRemovalResult("b@x", false, "propietario"));
    }

    [Fact]
    public async Task Parse_SingleElementCollapse_Tolerated()
    {
        // ConvertTo-Json can emit a bare object for 1-element lists.
        var (ops, _, _) = Build("""{"Group":"a@x","Success":true,"Detail":"quitado"}""");

        var results = await ops.RemoveFromDistributionGroupsAsync("u@a", new[] { "a@x" });

        results.Should().ContainSingle().Which.Should().Be(new DistributionGroupRemovalResult("a@x", true, "quitado"));
    }

    [Fact]
    public async Task Parse_EmptyOrNullJson_ReturnsEmpty()
    {
        var (ops, _, _) = Build(null);

        var results = await ops.RemoveFromDistributionGroupsAsync("u@a", new[] { "a@x" });

        results.Should().BeEmpty(); // caller treats unreported groups as failed
    }
}
