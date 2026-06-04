using FluentAssertions;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Grex365.PowerShell;
using Moq;

namespace Grex365.Core.Tests;

// DistributionListsService runs its EXO cmdlets through the external pwsh host now. The mock
// routes by script content (exists / create / list-members / add-member) and returns canned JSON,
// so these exercise the C# orchestration + JSON parsing without a tenant.
public class DistributionListsServiceTests
{
    private sealed class Harness
    {
        public Mock<IExternalExoRunner> Runner { get; } = new();
        public List<string> Bodies { get; } = new();
        public bool GroupExists { get; set; }
        public HashSet<string> Members { get; } = new(StringComparer.OrdinalIgnoreCase);
        public bool ThrowOnCreate { get; set; }

        public DistributionListsService Build()
        {
            Runner.Setup(r => r.RunAsync(It.IsAny<string>(), It.IsAny<IProgress<LogEntry>>(), It.IsAny<CancellationToken>()))
                .Returns<string, IProgress<LogEntry>?, CancellationToken>((body, _, __) =>
                {
                    Bodies.Add(body);
                    if (body.Contains("New-DistributionGroup"))
                    {
                        if (ThrowOnCreate) throw new InvalidOperationException("Operación Exchange Online falló: ya existe");
                        return Task.FromResult<string?>("""{"Note":"DL creada"}""");
                    }
                    if (body.Contains("Add-DistributionGroupMember"))
                    {
                        return Task.FromResult<string?>("""{"Note":"ok"}""");
                    }
                    if (body.Contains("Get-DistributionGroupMember"))
                    {
                        var items = string.Join(",", Members.Select(m => $$"""{"Smtp":"{{m}}"}"""));
                        return Task.FromResult<string?>("[" + items + "]");
                    }
                    // Get-DistributionGroup existence probe
                    return Task.FromResult<string?>(GroupExists ? """{"Found":true}""" : """{"Found":false}""");
                });
            return new DistributionListsService(Runner.Object);
        }
    }

    private static IReadOnlyList<BulkGroupRow> Rows(params (string g, string e)[] items) =>
        items.Select(i => new BulkGroupRow(i.g, i.e, "DL")).ToList();

    [Fact]
    public async Task EmptyDomain_Throws()
    {
        var sut = new Harness().Build();
        var act = () => sut.CreateFromRowsAsync(Rows(("Sales", "a@x.com")), "  ");
        await act.Should().ThrowAsync<ArgumentException>();
    }

    [Fact]
    public async Task NewGroup_CreatesThenAddsMembers()
    {
        var h = new Harness { GroupExists = false };
        var sut = h.Build();

        var results = await sut.CreateFromRowsAsync(Rows(("Sales", "a@contoso.com"), ("Sales", "b@contoso.com")), "contoso.com");

        results.Should().ContainSingle(r => r.Action == "Created");
        results.Count(r => r.Action == "MemberAdded").Should().Be(2);
    }

    [Fact]
    public async Task ExistingGroup_SkipsCreate_AndSkipsExistingMember()
    {
        var h = new Harness { GroupExists = true };
        h.Members.Add("a@contoso.com");
        var sut = h.Build();

        var results = await sut.CreateFromRowsAsync(Rows(("Sales", "a@contoso.com"), ("Sales", "b@contoso.com")), "contoso.com");

        results.Should().ContainSingle(r => r.Action == "Skipped");
        results.Should().ContainSingle(r => r.Action == "MemberSkipped" && r.UserEmail == "a@contoso.com");
        results.Should().ContainSingle(r => r.Action == "MemberAdded" && r.UserEmail == "b@contoso.com");
    }

    [Fact]
    public async Task InvalidEmail_ReportsError_NotAdded()
    {
        var h = new Harness { GroupExists = true };
        var sut = h.Build();

        var results = await sut.CreateFromRowsAsync(Rows(("Sales", "not-an-email")), "contoso.com");

        results.Should().ContainSingle(r => r.Action == "Error" && r.Detail.Contains("inválido"));
        h.Bodies.Should().NotContain(b => b.Contains("Add-DistributionGroupMember"));
    }

    [Fact]
    public async Task CreateFails_ReportsError_AndSkipsMembers()
    {
        var h = new Harness { GroupExists = false, ThrowOnCreate = true };
        var sut = h.Build();

        var results = await sut.CreateFromRowsAsync(Rows(("Sales", "a@contoso.com")), "contoso.com");

        results.Should().ContainSingle(r => r.Action == "Error" && r.Detail.Contains("Creación fallida"));
        h.Bodies.Should().NotContain(b => b.Contains("Add-DistributionGroupMember"));
    }

    [Fact]
    public async Task GroupNameWithQuote_IsEscaped_NotInjected()
    {
        var h = new Harness { GroupExists = true };
        var sut = h.Build();

        await sut.CreateFromRowsAsync(Rows(("O'Brien", "a@contoso.com")), "contoso.com");

        // Single quote doubled by ExternalExoRunner.Lit — never a bare quote that would break the
        // single-quoted PowerShell literal (injection-safe).
        h.Bodies.Should().Contain(b => b.Contains("O''Brien"));
        h.Bodies.Should().NotContain(b => b.Contains("'O'Brien'"));
    }
}
