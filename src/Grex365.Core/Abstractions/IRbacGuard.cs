namespace Grex365.Core.Abstractions;

public interface IMembershipChecker
{
    Task<bool> IsMemberOfAsync(string groupId, CancellationToken cancellationToken = default);
}

public interface IRbacGuard
{
    Task<RbacDecision> EvaluateAsync(CancellationToken cancellationToken = default);

    void Invalidate();
}

public sealed record RbacDecision(bool Allowed, string Reason);
