using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.Core.Operations;

// What each channel is allowed to do beyond reading. Agents are read-only unless the operator
// opens writes for that session, and never run critical operations by themselves.
public sealed record OperationPolicy(bool AllowAgentWrites = false)
{
    public string? Deny(OperationDescriptor descriptor, OperationChannel channel)
    {
        if (descriptor.Impact == OperationImpact.Read)
        {
            return null;
        }

        if (channel == OperationChannel.Scheduler)
        {
            return "Las tareas programadas solo ejecutan operaciones de lectura.";
        }

        if (channel != OperationChannel.Mcp)
        {
            return null;
        }

        if (descriptor.Impact == OperationImpact.Critical)
        {
            return "Una operación crítica no se ejecuta a petición de un agente: debe lanzarla un operador desde la aplicación o la línea de comandos.";
        }

        return AllowAgentWrites
            ? null
            : "Los agentes están en modo solo lectura. El operador puede abrir las escrituras al arrancar el servidor MCP.";
    }
}

// The single path every channel uses to run an operation. The order of the gates is the
// security model: validate, connect, authorize, look at the real state, require a confirmed
// plan, execute, and leave a record of who asked for what and how it ended.
public sealed class OperationExecutor
{
    private readonly OperationRegistry _registry;
    private readonly IOperationSession _session;
    private readonly IRbacGuard _rbac;
    private readonly IAuditLog _audit;
    private readonly PlanTokenService _plans;
    private readonly OperationPolicy _policy;
    private readonly TimeProvider _time;

    public OperationExecutor(
        OperationRegistry registry,
        IOperationSession session,
        IRbacGuard rbac,
        IAuditLog audit,
        PlanTokenService plans,
        OperationPolicy policy,
        TimeProvider? time = null)
    {
        _registry = registry;
        _session = session;
        _rbac = rbac;
        _audit = audit;
        _plans = plans;
        _policy = policy;
        _time = time ?? TimeProvider.System;
    }

    public async Task<OperationResult> RunAsync(
        OperationRequest request,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var operation = _registry.Find(request.OperationId);
        if (operation is null)
        {
            return new OperationResult(OperationStatus.NotFound, $"No existe la operación '{request.OperationId}'.");
        }

        var descriptor = operation.Descriptor;
        var arguments = OperationArguments.Validate(descriptor, request.Arguments, out var errors);
        if (errors.Count > 0)
        {
            return new OperationResult(OperationStatus.InvalidArguments, "Argumentos no válidos.", Reasons: errors);
        }

        var isWrite = descriptor.Impact != OperationImpact.Read;

        if (_policy.Deny(descriptor, request.Channel) is { } policyReason)
        {
            return await CloseAsync(request, descriptor, arguments,
                new OperationResult(OperationStatus.Denied, policyReason), cancellationToken).ConfigureAwait(false);
        }

        try
        {
            await _session.EnsureAsync(descriptor, cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            return new OperationResult(OperationStatus.NotConnected, $"Sin conexión: {ex.Message}");
        }

        if (isWrite)
        {
            var decision = await _rbac.EvaluateAsync(cancellationToken).ConfigureAwait(false);
            if (!decision.Allowed)
            {
                return await CloseAsync(request, descriptor, arguments,
                    new OperationResult(OperationStatus.Denied, decision.Reason), cancellationToken).ConfigureAwait(false);
            }
        }

        try
        {
            if (isWrite)
            {
                var preflight = await operation.PreflightAsync(arguments, cancellationToken).ConfigureAwait(false);
                if (preflight.IsBlocked)
                {
                    return await CloseAsync(request, descriptor, arguments,
                        new OperationResult(OperationStatus.Blocked, "Operación detenida por una comprobación previa.",
                            Data: preflight.Plan, Reasons: preflight.BlockingReasons),
                        cancellationToken).ConfigureAwait(false);
                }

                if (request.DryRun)
                {
                    return new OperationResult(OperationStatus.DryRun, preflight.Plan,
                        PlanId: _plans.Issue(descriptor.Id, arguments));
                }

                if (!_plans.Validate(request.PlanId, descriptor.Id, arguments))
                {
                    return new OperationResult(OperationStatus.ConfirmationRequired,
                        "Esta operación modifica el tenant. Ejecútala primero en simulación, revisa el plan y repítela con el plan_id que devuelve (caduca a los 15 minutos y solo vale para esos mismos argumentos).",
                        Data: preflight.Plan);
                }
            }

            var data = await operation.ExecuteAsync(arguments, progress, cancellationToken).ConfigureAwait(false);
            return await CloseAsync(request, descriptor, arguments,
                new OperationResult(OperationStatus.Ok, $"{descriptor.Title}: correcto.", data),
                cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            return await CloseAsync(request, descriptor, arguments,
                new OperationResult(OperationStatus.Failed, ex.Message), cancellationToken).ConfigureAwait(false);
        }
    }

    // Every attempted write is recorded, including the ones that were refused; reads are not.
    private async Task<OperationResult> CloseAsync(
        OperationRequest request,
        OperationDescriptor descriptor,
        OperationArguments arguments,
        OperationResult result,
        CancellationToken cancellationToken)
    {
        if (descriptor.Impact == OperationImpact.Read)
        {
            return result;
        }

        var detail = arguments.Canonical().Replace("\n", "; ", StringComparison.Ordinal);
        if (result.Reasons is { Count: > 0 })
        {
            detail += " | " + string.Join(" | ", result.Reasons);
        }

        var record = new AuditRecord(
            _time.GetUtcNow(),
            request.Actor,
            $"{request.Channel}:{descriptor.Id}",
            result.Status.ToString(),
            result.Summary,
            detail);
        await _audit.WriteAsync(record, cancellationToken).ConfigureAwait(false);
        return result;
    }
}
