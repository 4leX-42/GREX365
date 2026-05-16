using FluentAssertions;
using Grex365.Core.Abstractions;
using Grex365.Core.Connections;
using Grex365.Core.Models;
using Moq;

namespace Grex365.Core.Tests;

public class ConnectionStateMonitorTests
{
    private static Mock<IGraphConnection> Graph(bool connected, bool live, string? tenant = "t", string? account = "acc")
    {
        var m = new Mock<IGraphConnection>();
        m.SetupGet(g => g.IsConnected).Returns(connected);
        m.Setup(g => g.CheckLiveAsync(It.IsAny<CancellationToken>())).ReturnsAsync(live);
        m.SetupGet(g => g.TenantId).Returns(tenant);
        m.SetupGet(g => g.Account).Returns(account);
        return m;
    }

    private static Mock<IExchangeConnection> Exo(bool connected, bool live, string? org = "contoso.onmicrosoft.com")
    {
        var m = new Mock<IExchangeConnection>();
        m.SetupGet(e => e.IsConnected).Returns(connected);
        m.Setup(e => e.CheckLiveAsync(It.IsAny<CancellationToken>())).ReturnsAsync(live);
        m.SetupGet(e => e.Organization).Returns(org);
        m.SetupGet(e => e.TenantId).Returns((string?)null);
        return m;
    }

    [Fact]
    public void DefaultState_IsDisconnected()
    {
        var sut = new ConnectionStateMonitor(Graph(false, false).Object, Exo(false, false).Object);
        sut.Current.GraphConnected.Should().BeFalse();
        sut.Current.ExchangeConnected.Should().BeFalse();
    }

    [Fact]
    public async Task PollsAndReportsConnectedState()
    {
        var graph = Graph(true, true);
        var exo = Exo(true, true);
        var sut = new ConnectionStateMonitor(graph.Object, exo.Object, TimeSpan.FromMilliseconds(20));

        ConnectionState? captured = null;
        sut.PropertyChanged += (_, _) => captured = sut.Current;

        sut.Start();

        // wait briefly for poll cycle
        var deadline = DateTime.UtcNow.AddSeconds(2);
        while (captured is null && DateTime.UtcNow < deadline)
        {
            await Task.Delay(20);
        }
        await sut.DisposeAsync();

        captured.Should().NotBeNull();
        captured!.GraphConnected.Should().BeTrue();
        captured.ExchangeConnected.Should().BeTrue();
        captured.TenantId.Should().Be("t");
        captured.TenantDomain.Should().Be("contoso.onmicrosoft.com");
        captured.Account.Should().Be("acc");
    }

    [Fact]
    public async Task ProbeException_DoesNotKillLoop()
    {
        var graph = new Mock<IGraphConnection>();
        graph.SetupGet(g => g.IsConnected).Returns(true);
        var calls = 0;
        graph.Setup(g => g.CheckLiveAsync(It.IsAny<CancellationToken>())).ReturnsAsync(() =>
        {
            calls++;
            if (calls == 1) throw new InvalidOperationException("boom");
            return true;
        });
        graph.SetupGet(g => g.TenantId).Returns("t");

        var exo = Exo(true, true);
        var sut = new ConnectionStateMonitor(graph.Object, exo.Object, TimeSpan.FromMilliseconds(15));

        sut.Start();
        var deadline = DateTime.UtcNow.AddSeconds(2);
        while (calls < 2 && DateTime.UtcNow < deadline)
        {
            await Task.Delay(15);
        }
        await sut.DisposeAsync();

        calls.Should().BeGreaterThanOrEqualTo(2);
    }

    [Fact]
    public async Task Dispose_CompletesEvenIfPollingFails()
    {
        var graph = new Mock<IGraphConnection>();
        graph.SetupGet(g => g.IsConnected).Returns(true);
        graph.Setup(g => g.CheckLiveAsync(It.IsAny<CancellationToken>())).ThrowsAsync(new InvalidOperationException("x"));
        var sut = new ConnectionStateMonitor(graph.Object, Exo(false, false).Object, TimeSpan.FromMilliseconds(10));

        sut.Start();
        await Task.Delay(50);
        await sut.DisposeAsync();
        // no assertion: must not hang or throw
    }
}
