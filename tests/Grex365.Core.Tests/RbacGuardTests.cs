using FluentAssertions;
using Grex365.Core.Abstractions;
using Grex365.Core.Security;

namespace Grex365.Core.Tests;

public class RbacGuardTests
{
    private sealed class FakeChecker : IMembershipChecker
    {
        public bool Result { get; set; }
        public int Calls { get; private set; }
        public Exception? Throws { get; set; }
        public Task<bool> IsMemberOfAsync(string groupId, CancellationToken cancellationToken = default)
        {
            Calls++;
            if (Throws is not null) throw Throws;
            return Task.FromResult(Result);
        }
    }

    [Fact]
    public async Task NoGroupConfigured_AllowsAndShortCircuits()
    {
        var checker = new FakeChecker { Result = false };
        var guard = new RbacGuard(checker, () => null);
        var dec = await guard.EvaluateAsync();
        dec.Allowed.Should().BeTrue();
        dec.Reason.Should().Contain("RBAC no configurado");
        checker.Calls.Should().Be(0);
    }

    [Fact]
    public async Task EmptyOrWhitespaceGroup_AllowsAndShortCircuits()
    {
        var checker = new FakeChecker { Result = false };
        var guard = new RbacGuard(checker, () => "   ");
        (await guard.EvaluateAsync()).Allowed.Should().BeTrue();
        checker.Calls.Should().Be(0);
    }

    [Fact]
    public async Task Member_Allowed()
    {
        var checker = new FakeChecker { Result = true };
        var guard = new RbacGuard(checker, () => "00000000-0000-0000-0000-000000000001");
        var dec = await guard.EvaluateAsync();
        dec.Allowed.Should().BeTrue();
        dec.Reason.Should().Contain("Miembro del grupo");
    }

    [Fact]
    public async Task NotMember_Denied()
    {
        var checker = new FakeChecker { Result = false };
        var guard = new RbacGuard(checker, () => "g1");
        var dec = await guard.EvaluateAsync();
        dec.Allowed.Should().BeFalse();
        dec.Reason.Should().Contain("no pertenece");
    }

    [Fact]
    public async Task CheckerThrows_DeniedWithReason()
    {
        var checker = new FakeChecker { Throws = new InvalidOperationException("graph 403") };
        var guard = new RbacGuard(checker, () => "g1");
        var dec = await guard.EvaluateAsync();
        dec.Allowed.Should().BeFalse();
        dec.Reason.Should().Contain("graph 403");
    }

    [Fact]
    public async Task DecisionCached_DoesNotReCallChecker()
    {
        var checker = new FakeChecker { Result = true };
        var guard = new RbacGuard(checker, () => "g1");
        await guard.EvaluateAsync();
        await guard.EvaluateAsync();
        checker.Calls.Should().Be(1);
    }

    [Fact]
    public async Task Invalidate_ForcesReevaluation()
    {
        var checker = new FakeChecker { Result = true };
        var guard = new RbacGuard(checker, () => "g1");
        await guard.EvaluateAsync();
        guard.Invalidate();
        await guard.EvaluateAsync();
        checker.Calls.Should().Be(2);
    }

    [Fact]
    public async Task GroupIdTrimmedBeforeShortCircuit()
    {
        var checker = new FakeChecker { Result = true };
        var guard = new RbacGuard(checker, () => "  g1  ");
        var dec = await guard.EvaluateAsync();
        dec.Allowed.Should().BeTrue();
        checker.Calls.Should().Be(1);
    }
}
