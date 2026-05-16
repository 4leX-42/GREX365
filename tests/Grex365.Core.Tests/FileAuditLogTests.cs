using FluentAssertions;
using Grex365.Core.Abstractions;
using Grex365.Core.Audit;

namespace Grex365.Core.Tests;

public class FileAuditLogTests : IDisposable
{
    private readonly string _dir;

    public FileAuditLogTests()
    {
        _dir = Path.Combine(Path.GetTempPath(), "grex365-audit-" + Guid.NewGuid().ToString("N"));
    }

    public void Dispose()
    {
        if (Directory.Exists(_dir))
        {
            Directory.Delete(_dir, recursive: true);
        }
    }

    [Fact]
    public async Task Writes_And_Reads_Back_RoundTrip()
    {
        var log = new FileAuditLog(_dir);
        var record = new AuditRecord(
            Timestamp: new DateTimeOffset(2026, 5, 16, 10, 0, 0, TimeSpan.Zero),
            Actor: "alex",
            Source: "Users",
            Outcome: "OK",
            Message: "Asignada licencia E3 a user@x.com");

        await log.WriteAsync(record);
        var read = await log.ReadMonthAsync(2026, 5);
        read.Should().ContainSingle();
        read[0].Actor.Should().Be("alex");
        read[0].Source.Should().Be("Users");
        read[0].Outcome.Should().Be("OK");
        read[0].Message.Should().Contain("E3");
    }

    [Fact]
    public async Task Appends_Multiple_Records_AsJsonl()
    {
        var log = new FileAuditLog(_dir);
        for (var i = 0; i < 5; i++)
        {
            await log.WriteAsync(new AuditRecord(
                Timestamp: new DateTimeOffset(2026, 6, 1, 8, i, 0, TimeSpan.Zero),
                Actor: "alex",
                Source: "BulkGroups",
                Outcome: "OK",
                Message: $"create #{i}"));
        }
        var read = await log.ReadMonthAsync(2026, 6);
        read.Should().HaveCount(5);
        var rawLines = await File.ReadAllLinesAsync(log.GetMonthFilePath(2026, 6));
        rawLines.Length.Should().Be(5);
    }

    [Fact]
    public async Task ReadMonth_Returns_Empty_If_File_Missing()
    {
        var log = new FileAuditLog(_dir);
        var read = await log.ReadMonthAsync(2030, 1);
        read.Should().BeEmpty();
    }

    [Fact]
    public async Task Concurrent_Writes_DoNot_Corrupt()
    {
        var log = new FileAuditLog(_dir);
        var ts = new DateTimeOffset(2026, 7, 1, 0, 0, 0, TimeSpan.Zero);
        var tasks = Enumerable.Range(0, 50).Select(i => log.WriteAsync(new AuditRecord(
            Timestamp: ts.AddSeconds(i),
            Actor: "alex",
            Source: "Concurrent",
            Outcome: "OK",
            Message: $"row {i}"))).ToArray();
        await Task.WhenAll(tasks);
        var read = await log.ReadMonthAsync(2026, 7);
        read.Should().HaveCount(50);
    }
}
