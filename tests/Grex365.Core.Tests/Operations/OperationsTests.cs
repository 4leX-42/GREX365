using System.Text.Json.Nodes;
using FluentAssertions;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Grex365.Core.Operations;
using Grex365.Core.Operations.Builtin;
using Grex365.Core.Operations.Mcp;
using Moq;

namespace Grex365.Core.Tests.Operations;

internal sealed class FakeTime : TimeProvider
{
    public DateTimeOffset Now { get; set; } = new(2026, 1, 1, 12, 0, 0, TimeSpan.Zero);

    public override DateTimeOffset GetUtcNow() => Now;
}

internal sealed class RecordingAuditLog : IAuditLog
{
    public List<AuditRecord> Records { get; } = [];

    public Task WriteAsync(AuditRecord record, CancellationToken cancellationToken = default)
    {
        Records.Add(record);
        return Task.CompletedTask;
    }

    public Task<IReadOnlyList<AuditRecord>> ReadMonthAsync(int year, int month, CancellationToken cancellationToken = default) =>
        Task.FromResult<IReadOnlyList<AuditRecord>>(Records);

    public string GetMonthFilePath(int year, int month) => string.Empty;
}

internal sealed class FakeSession : IOperationSession
{
    public Exception? Failure { get; set; }

    public Task EnsureAsync(OperationDescriptor descriptor, CancellationToken cancellationToken) =>
        Failure is null ? Task.CompletedTask : Task.FromException(Failure);
}

internal sealed class FakeRbac : IRbacGuard
{
    public RbacDecision Decision { get; set; } = new(true, "ok");

    public Task<RbacDecision> EvaluateAsync(CancellationToken cancellationToken = default) => Task.FromResult(Decision);

    public void Invalidate()
    {
    }
}

internal sealed class FakeOperation : IOperation
{
    public FakeOperation(string id, OperationImpact impact)
    {
        Descriptor = new OperationDescriptor(
            id, "test", id, "desc", impact,
            [new("target", OperationParameterType.String, "t"), new("count", OperationParameterType.Integer, "c", Required: false)],
            ["Some.Permission"]);
    }

    public OperationDescriptor Descriptor { get; }

    public OperationPreflight Preflight { get; set; } = OperationPreflight.Allow("plan");

    public int Executions { get; private set; }

    public Exception? Failure { get; set; }

    public Task<OperationPreflight> PreflightAsync(OperationArguments arguments, CancellationToken cancellationToken) =>
        Task.FromResult(Preflight);

    public Task<object?> ExecuteAsync(OperationArguments arguments, IProgress<LogEntry>? progress, CancellationToken cancellationToken)
    {
        Executions++;
        return Failure is null ? Task.FromResult<object?>("done") : Task.FromException<object?>(Failure);
    }
}

internal sealed class Harness
{
    public FakeTime Time { get; } = new();

    public RecordingAuditLog Audit { get; } = new();

    public FakeSession Session { get; } = new();

    public FakeRbac Rbac { get; } = new();

    public FakeOperation Read { get; } = new("thing.read", OperationImpact.Read);

    public FakeOperation Write { get; } = new("thing.write", OperationImpact.Write);

    public FakeOperation Critical { get; } = new("thing.destroy", OperationImpact.Critical);

    public OperationRegistry Registry => new([Read, Write, Critical]);

    public OperationExecutor Executor(OperationPolicy? policy = null) => new(
        Registry, Session, Rbac, Audit,
        new PlanTokenService(new byte[32], Time),
        policy ?? new OperationPolicy(),
        Time);

    public static OperationRequest Request(
        string id,
        OperationChannel channel = OperationChannel.Cli,
        bool dryRun = false,
        string? planId = null,
        string target = "a") =>
        new(id, new Dictionary<string, string?> { ["target"] = target }, channel, "tester", dryRun, planId);
}

public class OperationExecutorTests
{
    [Fact]
    public async Task Read_runs_directly_and_is_not_audited()
    {
        var h = new Harness();
        var result = await h.Executor().RunAsync(Harness.Request("thing.read"));

        result.Status.Should().Be(OperationStatus.Ok);
        result.Data.Should().Be("done");
        h.Audit.Records.Should().BeEmpty();
    }

    [Fact]
    public async Task Unknown_operation_is_not_found()
    {
        var result = await new Harness().Executor().RunAsync(Harness.Request("nope"));
        result.Status.Should().Be(OperationStatus.NotFound);
    }

    [Fact]
    public async Task Missing_unknown_and_mistyped_arguments_are_rejected_before_anything_runs()
    {
        var h = new Harness();
        var request = new OperationRequest(
            "thing.read",
            new Dictionary<string, string?> { ["bogus"] = "1", ["count"] = "many" },
            OperationChannel.Cli,
            "tester");

        var result = await h.Executor().RunAsync(request);

        result.Status.Should().Be(OperationStatus.InvalidArguments);
        result.Reasons.Should().HaveCount(3);
        h.Read.Executions.Should().Be(0);
    }

    [Fact]
    public async Task Connection_failure_stops_the_operation()
    {
        var h = new Harness();
        h.Session.Failure = new InvalidOperationException("sin certificado");

        var result = await h.Executor().RunAsync(Harness.Request("thing.read"));

        result.Status.Should().Be(OperationStatus.NotConnected);
        h.Read.Executions.Should().Be(0);
    }

    [Fact]
    public async Task Write_without_plan_asks_for_confirmation_and_changes_nothing()
    {
        var h = new Harness();
        var result = await h.Executor().RunAsync(Harness.Request("thing.write"));

        result.Status.Should().Be(OperationStatus.ConfirmationRequired);
        h.Write.Executions.Should().Be(0);
    }

    [Fact]
    public async Task Write_runs_with_the_plan_id_of_its_dry_run_and_is_audited()
    {
        var h = new Harness();
        var executor = h.Executor();

        var plan = await executor.RunAsync(Harness.Request("thing.write", dryRun: true));
        plan.Status.Should().Be(OperationStatus.DryRun);
        h.Write.Executions.Should().Be(0);

        var result = await executor.RunAsync(Harness.Request("thing.write", planId: plan.PlanId));

        result.Status.Should().Be(OperationStatus.Ok);
        h.Write.Executions.Should().Be(1);
        h.Audit.Records.Should().ContainSingle();
        h.Audit.Records[0].Actor.Should().Be("tester");
        h.Audit.Records[0].Source.Should().Be("Cli:thing.write");
        h.Audit.Records[0].Outcome.Should().Be("Ok");
        h.Audit.Records[0].Detail.Should().Contain("target=a");
    }

    [Fact]
    public async Task Plan_id_does_not_authorize_different_arguments()
    {
        var h = new Harness();
        var executor = h.Executor();
        var plan = await executor.RunAsync(Harness.Request("thing.write", dryRun: true, target: "a"));

        var result = await executor.RunAsync(Harness.Request("thing.write", planId: plan.PlanId, target: "b"));

        result.Status.Should().Be(OperationStatus.ConfirmationRequired);
        h.Write.Executions.Should().Be(0);
    }

    [Fact]
    public async Task Plan_id_expires()
    {
        var h = new Harness();
        var executor = h.Executor();
        var plan = await executor.RunAsync(Harness.Request("thing.write", dryRun: true));

        h.Time.Now += PlanTokenService.Lifetime + TimeSpan.FromSeconds(1);
        var result = await executor.RunAsync(Harness.Request("thing.write", planId: plan.PlanId));

        result.Status.Should().Be(OperationStatus.ConfirmationRequired);
        h.Write.Executions.Should().Be(0);
    }

    [Fact]
    public async Task Blocking_preflight_stops_even_with_a_valid_plan_and_is_audited()
    {
        var h = new Harness();
        var executor = h.Executor();
        var plan = await executor.RunAsync(Harness.Request("thing.write", dryRun: true));

        h.Write.Preflight = OperationPreflight.Block("plan", "el estado cambió");
        var result = await executor.RunAsync(Harness.Request("thing.write", planId: plan.PlanId));

        result.Status.Should().Be(OperationStatus.Blocked);
        result.Reasons.Should().Contain("el estado cambió");
        h.Write.Executions.Should().Be(0);
        h.Audit.Records.Should().ContainSingle(r => r.Outcome == "Blocked");
    }

    [Fact]
    public async Task Rbac_denial_stops_writes_but_not_reads()
    {
        var h = new Harness();
        h.Rbac.Decision = new RbacDecision(false, "no autorizado");
        var executor = h.Executor();

        (await executor.RunAsync(Harness.Request("thing.read"))).Status.Should().Be(OperationStatus.Ok);
        var write = await executor.RunAsync(Harness.Request("thing.write", dryRun: true));

        write.Status.Should().Be(OperationStatus.Denied);
        h.Audit.Records.Should().ContainSingle(r => r.Outcome == "Denied");
    }

    [Fact]
    public async Task Agents_are_read_only_by_default()
    {
        var h = new Harness();
        var executor = h.Executor();

        (await executor.RunAsync(Harness.Request("thing.read", OperationChannel.Mcp))).Status.Should().Be(OperationStatus.Ok);
        var write = await executor.RunAsync(Harness.Request("thing.write", OperationChannel.Mcp, dryRun: true));

        write.Status.Should().Be(OperationStatus.Denied);
        h.Audit.Records.Should().ContainSingle(r => r.Source == "Mcp:thing.write");
    }

    [Fact]
    public async Task Agents_can_write_when_the_operator_allows_it_but_never_run_critical_operations()
    {
        var h = new Harness();
        var executor = h.Executor(new OperationPolicy(AllowAgentWrites: true));

        var plan = await executor.RunAsync(Harness.Request("thing.write", OperationChannel.Mcp, dryRun: true));
        var write = await executor.RunAsync(Harness.Request("thing.write", OperationChannel.Mcp, planId: plan.PlanId));
        var critical = await executor.RunAsync(Harness.Request("thing.destroy", OperationChannel.Mcp, dryRun: true));

        write.Status.Should().Be(OperationStatus.Ok);
        critical.Status.Should().Be(OperationStatus.Denied);
        h.Critical.Executions.Should().Be(0);
    }

    [Fact]
    public async Task Scheduler_only_reads()
    {
        var h = new Harness();
        var result = await h.Executor().RunAsync(Harness.Request("thing.write", OperationChannel.Scheduler, dryRun: true));
        result.Status.Should().Be(OperationStatus.Denied);
    }

    [Fact]
    public async Task Execution_failure_is_reported_and_audited()
    {
        var h = new Harness();
        h.Write.Failure = new InvalidOperationException("Graph dijo que no");
        var executor = h.Executor();
        var plan = await executor.RunAsync(Harness.Request("thing.write", dryRun: true));

        var result = await executor.RunAsync(Harness.Request("thing.write", planId: plan.PlanId));

        result.Status.Should().Be(OperationStatus.Failed);
        result.Summary.Should().Be("Graph dijo que no");
        h.Audit.Records.Should().ContainSingle(r => r.Outcome == "Failed");
    }

    [Fact]
    public void Registry_rejects_duplicate_ids()
    {
        var act = () => new OperationRegistry([new FakeOperation("x", OperationImpact.Read), new FakeOperation("X", OperationImpact.Read)]);
        act.Should().Throw<InvalidOperationException>();
    }
}

public class LicenseRemovalGuardTests
{
    private static MailboxInfo Mailbox(
        long? bytes = 1024,
        bool litigation = false,
        int holds = 0,
        bool archive = false) =>
        new("id", "Name", "a@contoso.com", "UserMailbox", litigation, holds, archive, bytes);

    [Fact]
    public void No_mailbox_allows_removal() =>
        LicenseRemovalGuard.Evaluate(null).Should().BeEmpty();

    [Fact]
    public void Small_mailbox_without_holds_or_archive_allows_removal() =>
        LicenseRemovalGuard.Evaluate(Mailbox()).Should().BeEmpty();

    [Fact]
    public void Mailbox_over_50_gb_blocks() =>
        LicenseRemovalGuard.Evaluate(Mailbox(bytes: 51L * 1024 * 1024 * 1024)).Should().ContainSingle();

    [Fact]
    public void Unknown_size_blocks() =>
        LicenseRemovalGuard.Evaluate(Mailbox(bytes: null)).Should().ContainSingle();

    [Fact]
    public void Litigation_hold_blocks() =>
        LicenseRemovalGuard.Evaluate(Mailbox(litigation: true)).Should().ContainSingle();

    [Fact]
    public void In_place_hold_blocks() =>
        LicenseRemovalGuard.Evaluate(Mailbox(holds: 2)).Should().ContainSingle();

    [Fact]
    public void Archive_blocks() =>
        LicenseRemovalGuard.Evaluate(Mailbox(archive: true)).Should().ContainSingle();

    [Fact]
    public void Every_condition_is_reported() =>
        LicenseRemovalGuard.Evaluate(Mailbox(bytes: 60L * 1024 * 1024 * 1024, litigation: true, holds: 1, archive: true))
            .Should().HaveCount(4);
}

public class UserLicenseRemoveOperationTests
{
    private static readonly Guid Sku = Guid.Parse("11111111-1111-1111-1111-111111111111");
    private static readonly UserSummary User = new("uid", "Ana", "ana@contoso.com", "ana@contoso.com", true, false, 1, null);

    private static (UserLicenseRemoveOperation Operation, Mock<IUsersService> Users, Mock<IExternalExoOps> Exo) Build(MailboxInfo? mailbox)
    {
        var users = new Mock<IUsersService>();
        users.Setup(u => u.GetByIdAsync("ana@contoso.com", It.IsAny<CancellationToken>())).ReturnsAsync(User);
        users.Setup(u => u.ListSkusAsync(It.IsAny<CancellationToken>())).ReturnsAsync([new SkuInfo(Sku, "SPE_E3", 10, 5)]);
        users.Setup(u => u.GetAssignedLicensesAsync("uid", It.IsAny<CancellationToken>())).ReturnsAsync([Sku]);
        var exo = new Mock<IExternalExoOps>();
        exo.Setup(e => e.GetMailboxFactsAsync("ana@contoso.com", null, It.IsAny<CancellationToken>())).ReturnsAsync(mailbox);
        return (new UserLicenseRemoveOperation(users.Object, exo.Object), users, exo);
    }

    private static OperationArguments Args(UserLicenseRemoveOperation operation, string sku = "SPE_E3") =>
        OperationArguments.Validate(
            operation.Descriptor,
            new Dictionary<string, string?> { ["user"] = "ana@contoso.com", ["sku"] = sku },
            out _);

    [Fact]
    public void Is_classified_as_critical()
    {
        var (operation, _, _) = Build(null);
        operation.Descriptor.Impact.Should().Be(OperationImpact.Critical);
    }

    [Fact]
    public async Task Preflight_allows_a_safe_mailbox()
    {
        var (operation, _, _) = Build(new MailboxInfo("id", "Ana", "ana@contoso.com", "UserMailbox", TotalItemBytes: 1024));
        var preflight = await operation.PreflightAsync(Args(operation), CancellationToken.None);
        preflight.IsBlocked.Should().BeFalse();
    }

    [Fact]
    public async Task Preflight_blocks_a_mailbox_on_hold()
    {
        var (operation, _, _) = Build(new MailboxInfo("id", "Ana", "ana@contoso.com", "UserMailbox", LitigationHoldEnabled: true, TotalItemBytes: 1024));
        var preflight = await operation.PreflightAsync(Args(operation), CancellationToken.None);
        preflight.IsBlocked.Should().BeTrue();
    }

    [Fact]
    public async Task Preflight_fails_closed_when_the_mailbox_cannot_be_read()
    {
        var (operation, _, exo) = Build(null);
        exo.Setup(e => e.GetMailboxFactsAsync("ana@contoso.com", null, It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("EXO caído"));

        var act = () => operation.PreflightAsync(Args(operation), CancellationToken.None);

        await act.Should().ThrowAsync<InvalidOperationException>();
    }

    [Fact]
    public async Task Preflight_blocks_when_the_license_is_not_assigned()
    {
        var (operation, users, _) = Build(null);
        users.Setup(u => u.GetAssignedLicensesAsync("uid", It.IsAny<CancellationToken>())).ReturnsAsync([]);
        var preflight = await operation.PreflightAsync(Args(operation), CancellationToken.None);
        preflight.IsBlocked.Should().BeTrue();
    }

    [Fact]
    public async Task Execute_fails_when_the_tenant_still_shows_the_license()
    {
        var (operation, users, _) = Build(null);

        var act = () => operation.ExecuteAsync(Args(operation), null, CancellationToken.None);

        await act.Should().ThrowAsync<InvalidOperationException>().WithMessage("*sigue asignada*");
        users.Verify(u => u.RemoveLicenseAsync("uid", Sku, null, It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task Execute_succeeds_when_the_tenant_confirms_the_removal()
    {
        var (operation, users, _) = Build(null);
        users.Setup(u => u.RemoveLicenseAsync("uid", Sku, null, It.IsAny<CancellationToken>()))
            .Callback(() => users.Setup(u => u.GetAssignedLicensesAsync("uid", It.IsAny<CancellationToken>())).ReturnsAsync([]))
            .Returns(Task.CompletedTask);

        var outcome = await operation.ExecuteAsync(Args(operation), null, CancellationToken.None);

        outcome.Should().BeOfType<LicenseRemovalOutcome>().Which.StillAssigned.Should().BeFalse();
    }
}

public class McpServerTests
{
    private static (McpServer Server, Harness Harness) Build(OperationPolicy? policy = null)
    {
        var h = new Harness();
        return (new McpServer(h.Registry, h.Executor(policy), "agent", "1.2.3"), h);
    }

    private static async Task<JsonObject> CallAsync(McpServer server, string json) =>
        (await server.HandleAsync(json))!;

    [Fact]
    public async Task Initialize_echoes_a_supported_protocol_and_offers_tools()
    {
        var (server, _) = Build();
        var response = await CallAsync(server, """{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-03-26"}}""");

        response["result"]!["protocolVersion"]!.GetValue<string>().Should().Be("2025-03-26");
        response["result"]!["capabilities"]!["tools"].Should().NotBeNull();
        response["result"]!["serverInfo"]!["version"]!.GetValue<string>().Should().Be("1.2.3");
    }

    [Fact]
    public async Task Notifications_get_no_response()
    {
        var (server, _) = Build();
        (await server.HandleAsync("""{"jsonrpc":"2.0","method":"notifications/initialized"}""")).Should().BeNull();
    }

    [Fact]
    public async Task Tools_list_exposes_every_operation_with_its_schema_and_write_controls()
    {
        var (server, _) = Build();
        var response = await CallAsync(server, """{"jsonrpc":"2.0","id":2,"method":"tools/list"}""");
        var tools = response["result"]!["tools"]!.AsArray();

        tools.Select(t => t!["name"]!.GetValue<string>()).Should().BeEquivalentTo("thing_read", "thing_write", "thing_destroy");

        var read = tools.Single(t => t!["name"]!.GetValue<string>() == "thing_read")!;
        read["annotations"]!["readOnlyHint"]!.GetValue<bool>().Should().BeTrue();
        read["inputSchema"]!["required"]!.AsArray().Select(r => r!.GetValue<string>()).Should().Equal("target");
        read["inputSchema"]!["properties"]!["dry_run"].Should().BeNull();

        var write = tools.Single(t => t!["name"]!.GetValue<string>() == "thing_write")!;
        write["inputSchema"]!["properties"]!["dry_run"].Should().NotBeNull();
        write["inputSchema"]!["properties"]!["plan_id"].Should().NotBeNull();
    }

    [Fact]
    public async Task Tools_call_runs_a_read_through_the_executor()
    {
        var (server, h) = Build();
        var response = await CallAsync(server,
            """{"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"thing_read","arguments":{"target":"a","count":3}}}""");

        response["result"]!["isError"]!.GetValue<bool>().Should().BeFalse();
        response["result"]!["content"]![0]!["text"]!.GetValue<string>().Should().Contain("\"status\":\"Ok\"");
        h.Read.Executions.Should().Be(1);
    }

    [Fact]
    public async Task Tools_call_cannot_write_by_default()
    {
        var (server, h) = Build();
        var response = await CallAsync(server,
            """{"jsonrpc":"2.0","id":4,"method":"tools/call","params":{"name":"thing_write","arguments":{"target":"a","dry_run":true}}}""");

        response["result"]!["isError"]!.GetValue<bool>().Should().BeTrue();
        response["result"]!["content"]![0]!["text"]!.GetValue<string>().Should().Contain("\"status\":\"Denied\"");
        h.Write.Executions.Should().Be(0);
    }

    [Fact]
    public async Task Tools_call_writes_in_two_steps_when_allowed()
    {
        var (server, h) = Build(new OperationPolicy(AllowAgentWrites: true));
        var dry = await CallAsync(server,
            """{"jsonrpc":"2.0","id":5,"method":"tools/call","params":{"name":"thing_write","arguments":{"target":"a","dry_run":true}}}""");
        var planId = JsonNode.Parse(dry["result"]!["content"]![0]!["text"]!.GetValue<string>())!["planId"]!.GetValue<string>();

        var run = await CallAsync(server,
            """{"jsonrpc":"2.0","id":6,"method":"tools/call","params":{"name":"thing_write","arguments":{"target":"a","plan_id":"PLAN"}}}"""
                .Replace("PLAN", planId));

        run["result"]!["isError"]!.GetValue<bool>().Should().BeFalse();
        h.Write.Executions.Should().Be(1);
    }

    [Fact]
    public async Task Unknown_method_and_bad_json_return_protocol_errors()
    {
        var (server, _) = Build();
        (await CallAsync(server, """{"jsonrpc":"2.0","id":7,"method":"resources/list"}"""))["error"]!["code"]!.GetValue<int>().Should().Be(-32601);
        (await CallAsync(server, "{nope"))["error"]!["code"]!.GetValue<int>().Should().Be(-32700);
    }

    [Fact]
    public async Task Run_speaks_line_delimited_json_until_the_input_ends()
    {
        var (server, _) = Build();
        var input = new StringReader(
            """{"jsonrpc":"2.0","id":1,"method":"ping"}""" + "\n"
            + """{"jsonrpc":"2.0","method":"notifications/initialized"}""" + "\n"
            + """{"jsonrpc":"2.0","id":2,"method":"tools/list"}""" + "\n");
        var output = new StringWriter();

        await server.RunAsync(input, output);

        output.ToString().Split('\n', StringSplitOptions.RemoveEmptyEntries).Should().HaveCount(2);
    }
}
