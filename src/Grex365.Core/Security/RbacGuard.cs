using Grex365.Core.Abstractions;

namespace Grex365.Core.Security;

public sealed class RbacGuard : IRbacGuard
{
    private readonly IMembershipChecker _checker;
    private readonly Func<string?> _groupIdProvider;
    private RbacDecision? _cached;
    private readonly object _gate = new();

    public RbacGuard(IMembershipChecker checker, Func<string?> groupIdProvider)
    {
        _checker = checker;
        _groupIdProvider = groupIdProvider;
    }

    public async Task<RbacDecision> EvaluateAsync(CancellationToken cancellationToken = default)
    {
        lock (_gate)
        {
            if (_cached is not null)
            {
                return _cached;
            }
        }

        var groupId = _groupIdProvider()?.Trim();
        RbacDecision decision;
        if (string.IsNullOrEmpty(groupId))
        {
            decision = new RbacDecision(true, "RBAC no configurado (sin restricción)");
        }
        else
        {
            try
            {
                var member = await _checker.IsMemberOfAsync(groupId, cancellationToken).ConfigureAwait(false);
                decision = member
                    ? new RbacDecision(true, $"Miembro del grupo {groupId}")
                    : new RbacDecision(false, $"No autorizado: tu cuenta no pertenece al grupo {groupId}");
            }
            catch (Exception ex)
            {
                decision = new RbacDecision(false, $"No autorizado: error consultando membership ({ex.Message})");
            }
        }

        lock (_gate)
        {
            _cached = decision;
        }
        return decision;
    }

    public void Invalidate()
    {
        lock (_gate)
        {
            _cached = null;
        }
    }
}
