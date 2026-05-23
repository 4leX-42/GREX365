using System.ComponentModel;
using FluentAssertions;
using Grex365.Core.Audit;

namespace Grex365.Core.Tests;

public class InMemoryAuditFindingsStoreTests
{
    [Fact]
    public void Initial_Defaults_AreEmpty()
    {
        var store = new InMemoryAuditFindingsStore();

        store.LastRunAt.Should().BeNull();
        store.LastAuditName.Should().BeNull();
        store.ErrorCount.Should().Be(0);
        store.WarnCount.Should().Be(0);
        store.InfoCount.Should().Be(0);
    }

    [Fact]
    public void Update_SetsAllFields()
    {
        var store = new InMemoryAuditFindingsStore();

        store.Update("Identity", errorCount: 3, warnCount: 5, infoCount: 7);

        store.LastAuditName.Should().Be("Identity");
        store.ErrorCount.Should().Be(3);
        store.WarnCount.Should().Be(5);
        store.InfoCount.Should().Be(7);
        store.LastRunAt.Should().NotBeNull();
        store.LastRunAt!.Value.Should().BeCloseTo(DateTimeOffset.UtcNow, TimeSpan.FromSeconds(5));
    }

    [Fact]
    public void Update_RaisesPropertyChanged_ForAllFiveFields()
    {
        var store = new InMemoryAuditFindingsStore();
        var raised = new List<string>();
        store.PropertyChanged += (_, e) => raised.Add(e.PropertyName ?? string.Empty);

        store.Update("Audit", 1, 2, 3);

        raised.Should().BeEquivalentTo([
            nameof(InMemoryAuditFindingsStore.LastAuditName),
            nameof(InMemoryAuditFindingsStore.ErrorCount),
            nameof(InMemoryAuditFindingsStore.WarnCount),
            nameof(InMemoryAuditFindingsStore.InfoCount),
            nameof(InMemoryAuditFindingsStore.LastRunAt),
        ]);
    }

    [Fact]
    public void Update_Twice_OverwritesPreviousValues()
    {
        var store = new InMemoryAuditFindingsStore();
        store.Update("First", 10, 20, 30);
        var firstRun = store.LastRunAt;

        Thread.Sleep(10);  // ensure UtcNow advances
        store.Update("Second", 1, 2, 3);

        store.LastAuditName.Should().Be("Second");
        store.ErrorCount.Should().Be(1);
        store.WarnCount.Should().Be(2);
        store.InfoCount.Should().Be(3);
        store.LastRunAt.Should().BeAfter(firstRun!.Value);
    }

    [Fact]
    public void Update_AcceptsZeroCounts()
    {
        var store = new InMemoryAuditFindingsStore();

        store.Update("Clean", 0, 0, 0);

        store.ErrorCount.Should().Be(0);
        store.WarnCount.Should().Be(0);
        store.InfoCount.Should().Be(0);
        store.LastAuditName.Should().Be("Clean");
        store.LastRunAt.Should().NotBeNull();
    }
}
