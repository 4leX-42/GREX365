using System.ComponentModel;
using FluentAssertions;
using Grex365.App.ViewModels;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Moq;

namespace Grex365.App.Tests;

public class UsersViewModelAutoLoadSkusTests
{
    private sealed class FakeMonitor : IConnectionStateMonitor
    {
        public ConnectionState Current { get; private set; } = ConnectionState.Disconnected;
        public event PropertyChangedEventHandler? PropertyChanged;
        public void Start() { }
        public void Stop() { }
        public ValueTask DisposeAsync() => ValueTask.CompletedTask;

        public void Set(ConnectionState state)
        {
            Current = state;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Current)));
        }
    }

    private static SkuInfo Sku(string sku) => new(Guid.NewGuid(), sku, 10, 5);

    private static (UsersViewModel vm, Mock<IUsersService> users, FakeMonitor monitor) Build(
        ConnectionState initialState,
        IReadOnlyList<SkuInfo>? skus = null)
    {
        var users = new Mock<IUsersService>();
        users.Setup(u => u.ListSkusAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(skus ?? new[] { Sku("ENTERPRISEPACK") });
        var rbac = new Mock<IRbacGuard>();
        rbac.Setup(r => r.EvaluateAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new RbacDecision(true, "OK"));
        var monitor = new FakeMonitor();
        monitor.Set(initialState);
        var vm = new UsersViewModel(users.Object, new TestUiLogSink(), rbac.Object, new TestDialogService(), monitor);
        return (vm, users, monitor);
    }

    [Fact]
    public async Task Ctor_GraphAlreadyConnected_AutoLoadsSkus()
    {
        // Singleton VM materialized after Graph already came up: must back-fill.
        var (vm, users, _) = Build(new ConnectionState(GraphConnected: true, ExchangeConnected: false, null, null, null));

        // Auto-load fires fire-and-forget; await any pending continuations.
        for (var i = 0; i < 20 && vm.AvailableSkus.Count == 0; i++)
        {
            await Task.Delay(20);
        }

        vm.AvailableSkus.Should().HaveCount(1);
        users.Verify(u => u.ListSkusAsync(It.IsAny<CancellationToken>()), Times.AtLeastOnce);
    }

    [Fact]
    public async Task GraphConnects_AfterCtor_TriggersAutoLoad()
    {
        var (vm, users, monitor) = Build(ConnectionState.Disconnected);
        vm.AvailableSkus.Should().BeEmpty();
        users.Verify(u => u.ListSkusAsync(It.IsAny<CancellationToken>()), Times.Never);

        monitor.Set(new ConnectionState(GraphConnected: true, ExchangeConnected: false, null, null, null));

        for (var i = 0; i < 20 && vm.AvailableSkus.Count == 0; i++)
        {
            await Task.Delay(20);
        }

        vm.AvailableSkus.Should().HaveCount(1);
    }

    [Fact]
    public async Task GraphDisconnects_DoesNotClearOrReload()
    {
        var (vm, users, monitor) = Build(new ConnectionState(GraphConnected: true, ExchangeConnected: false, null, null, null));
        for (var i = 0; i < 20 && vm.AvailableSkus.Count == 0; i++) await Task.Delay(20);
        var firstCount = vm.AvailableSkus.Count;

        monitor.Set(ConnectionState.Disconnected);
        await Task.Delay(50);

        vm.AvailableSkus.Count.Should().Be(firstCount, "disconnect must not re-trigger load nor empty cached SKUs");
        users.Verify(u => u.ListSkusAsync(It.IsAny<CancellationToken>()), Times.AtMost(2));
    }

    [Fact]
    public async Task GraphReconnects_WithSkusAlreadyLoaded_DoesNotRefetch()
    {
        var (vm, users, monitor) = Build(new ConnectionState(GraphConnected: true, ExchangeConnected: false, null, null, null));
        for (var i = 0; i < 20 && vm.AvailableSkus.Count == 0; i++) await Task.Delay(20);

        // Simulate disconnect+reconnect cycle. SKUs already cached → no refetch.
        monitor.Set(ConnectionState.Disconnected);
        monitor.Set(new ConnectionState(GraphConnected: true, ExchangeConnected: false, null, null, null));
        await Task.Delay(50);

        users.Verify(u => u.ListSkusAsync(It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public void Ctor_WithoutMonitor_DoesNotThrow_NoAutoLoad()
    {
        // Backwards compat: existing tests construct VM without the monitor argument.
        var users = new Mock<IUsersService>();
        var rbac = new Mock<IRbacGuard>();
        rbac.Setup(r => r.EvaluateAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new RbacDecision(true, "OK"));

        var vm = new UsersViewModel(users.Object, new TestUiLogSink(), rbac.Object, new TestDialogService());

        vm.AvailableSkus.Should().BeEmpty();
        users.Verify(u => u.ListSkusAsync(It.IsAny<CancellationToken>()), Times.Never);
    }
}
