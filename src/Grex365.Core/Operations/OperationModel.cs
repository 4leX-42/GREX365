using Grex365.Core.Models;

namespace Grex365.Core.Operations;

// How much damage an operation can do. Drives every gate in OperationExecutor.
public enum OperationImpact
{
    // Queries and reports. Never changes the tenant.
    Read,

    // Reversible change. Needs authorization and a confirmed plan.
    Write,

    // Destructive or hard to undo. Needs a confirmed plan and is never run on an agent's own say-so.
    Critical,
}

// Who is asking. The same operation runs from every channel; the channel only tightens the policy.
public enum OperationChannel
{
    Gui,
    Cli,
    Mcp,
    Scheduler,
}

public enum OperationStatus
{
    Ok,
    DryRun,
    NotFound,
    InvalidArguments,
    NotConnected,
    Denied,
    Blocked,
    ConfirmationRequired,
    Failed,
}

public enum OperationParameterType
{
    String,
    Integer,
    Boolean,
}

public sealed record OperationParameter(
    string Name,
    OperationParameterType Type,
    string Description,
    bool Required = true);

// Everything a human, a CLI or an agent needs to discover and call an operation without
// reading its source: what it does, what it takes, what it needs and what it refuses to do.
public sealed record OperationDescriptor(
    string Id,
    string Module,
    string Title,
    string Description,
    OperationImpact Impact,
    IReadOnlyList<OperationParameter> Parameters,
    IReadOnlyList<string> RequiredPermissions,
    bool RequiresGraph = true,
    bool RequiresExchange = false,
    IReadOnlyList<string>? Safeguards = null);

public sealed record OperationRequest(
    string OperationId,
    IReadOnlyDictionary<string, string?> Arguments,
    OperationChannel Channel,
    string Actor,
    bool DryRun = false,
    string? PlanId = null);

public sealed record OperationResult(
    OperationStatus Status,
    string Summary,
    object? Data = null,
    IReadOnlyList<string>? Reasons = null,
    string? PlanId = null)
{
    public bool Succeeded => Status is OperationStatus.Ok or OperationStatus.DryRun;
}

// What a write operation found when it looked at the real state before touching anything.
public sealed record OperationPreflight(
    string Plan,
    IReadOnlyList<string> BlockingReasons)
{
    public bool IsBlocked => BlockingReasons.Count > 0;

    public static OperationPreflight Allow(string plan) => new(plan, Array.Empty<string>());

    public static OperationPreflight Block(string plan, params string[] reasons) => new(plan, reasons);
}

public interface IOperation
{
    OperationDescriptor Descriptor { get; }

    // Reads the real state and says what would happen and why it must not. Read operations
    // return OperationPreflight.Allow with an empty plan.
    Task<OperationPreflight> PreflightAsync(OperationArguments arguments, CancellationToken cancellationToken);

    Task<object?> ExecuteAsync(
        OperationArguments arguments,
        IProgress<LogEntry>? progress,
        CancellationToken cancellationToken);
}

// Brings up whatever connection the operation declares it needs. Each host supplies its own:
// the desktop app reports the session the operator already opened, headless hosts connect by
// certificate. Throws with a readable message when the connection cannot be established.
public interface IOperationSession
{
    Task EnsureAsync(OperationDescriptor descriptor, CancellationToken cancellationToken);
}
