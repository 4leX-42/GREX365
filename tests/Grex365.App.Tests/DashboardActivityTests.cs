using System.IO;
using FluentAssertions;
using Grex365.App.ViewModels;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Microsoft.Extensions.DependencyInjection;
using Moq;

namespace Grex365.App.Tests;

// DashboardViewModel.RefreshActivity: reads the local audit JSONL (current month, topping up
// with the previous month when sparse) and surfaces today's counters + the 5 newest ops.
public class DashboardActivityTests
{
    private static readonly DateTimeOffset Now = DateTimeOffset.Now;

    private static AuditRecord R(DateTimeOffset at, string outcome = "OK") =>
        new(at, "admin", "Users", outcome, "op");

    private static DashboardViewModel Build(IAuditLog? log)
    {
        var monitor = new Mock<IConnectionStateMonitor>();
        monitor.Setup(m => m.Current).Returns(new ConnectionState(false, false, null, null, null));
        var store = new Mock<IAuditFindingsStore>();
        var services = new ServiceCollection().BuildServiceProvider();
        return new DashboardViewModel(monitor.Object, store.Object, services, log);
    }

    [Fact]
    public async Task NoAuditLog_NoActivity_NoThrow()
    {
        var vm = Build(null);
        await vm.RefreshActivityCommand.ExecuteAsync(null);
        vm.HasActivity.Should().BeFalse();
        vm.RecentOps.Should().BeEmpty();
    }

    [Fact]
    public async Task PopulatesTodayCounters_AndRecent()
    {
        var log = new Mock<IAuditLog>();
        log.Setup(l => l.ReadMonthAsync(Now.Year, Now.Month, It.IsAny<CancellationToken>()))
            .ReturnsAsync(new[]
            {
                R(Now.AddMinutes(-5)),
                R(Now.AddMinutes(-10), "ERROR"),
                R(Now.AddMinutes(-15)),
                R(Now.AddMinutes(-20)),
                R(Now.AddMinutes(-25)),
                R(Now.AddMinutes(-30)),
            });
        var vm = Build(log.Object);

        await vm.RefreshActivityCommand.ExecuteAsync(null);

        vm.HasActivity.Should().BeTrue();
        vm.TodayOpsCount.Should().Be(6);
        vm.TodayErrorCount.Should().Be(1);
        vm.RecentOps.Should().HaveCount(5); // cap
        // Mes anterior NO leído: el mes en curso ya trae >= 5 registros.
        var prev = Now.AddMonths(-1);
        log.Verify(l => l.ReadMonthAsync(prev.Year, prev.Month, It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task SparseCurrentMonth_TopsUpWithPreviousMonth()
    {
        var prev = Now.AddMonths(-1);
        var log = new Mock<IAuditLog>();
        log.Setup(l => l.ReadMonthAsync(Now.Year, Now.Month, It.IsAny<CancellationToken>()))
            .ReturnsAsync(new[] { R(Now.AddMinutes(-1)) });
        log.Setup(l => l.ReadMonthAsync(prev.Year, prev.Month, It.IsAny<CancellationToken>()))
            .ReturnsAsync(new[] { R(prev), R(prev.AddMinutes(-1)) });
        var vm = Build(log.Object);

        await vm.RefreshActivityCommand.ExecuteAsync(null);

        vm.RecentOps.Should().HaveCount(3);
        vm.TodayOpsCount.Should().Be(1); // los del mes anterior no cuentan como "hoy"
    }

    [Fact]
    public async Task ReadFailure_BestEffort_NoActivity()
    {
        var log = new Mock<IAuditLog>();
        log.Setup(l => l.ReadMonthAsync(It.IsAny<int>(), It.IsAny<int>(), It.IsAny<CancellationToken>()))
            .ThrowsAsync(new IOException("jsonl bloqueado"));
        var vm = Build(log.Object);

        await vm.RefreshActivityCommand.ExecuteAsync(null);

        vm.HasActivity.Should().BeFalse();
    }
}
