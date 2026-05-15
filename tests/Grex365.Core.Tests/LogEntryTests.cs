using FluentAssertions;
using Grex365.Core.Models;

namespace Grex365.Core.Tests;

public class LogEntryTests
{
    [Fact]
    public void Info_Should_Set_Severity_Info()
    {
        var entry = LogEntry.Info("Test", "hola");
        entry.Severity.Should().Be(LogSeverity.Info);
        entry.Source.Should().Be("Test");
        entry.Message.Should().Be("hola");
        entry.Exception.Should().BeNull();
    }

    [Fact]
    public void Error_Should_Carry_Exception()
    {
        var ex = new InvalidOperationException("boom");
        var entry = LogEntry.Error("Test", ex.Message, ex);
        entry.Severity.Should().Be(LogSeverity.Error);
        entry.Exception.Should().BeSameAs(ex);
    }
}
