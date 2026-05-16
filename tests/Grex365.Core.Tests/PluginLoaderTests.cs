using FluentAssertions;
using Grex365.Core.Plugins;

namespace Grex365.Core.Tests;

public class PluginLoaderTests
{
    [Fact]
    public void NonExistent_Directory_Returns_Empty()
    {
        var report = PluginLoader.LoadFrom(Path.Combine(Path.GetTempPath(), "grex365-plugins-doesnt-exist-" + Guid.NewGuid().ToString("N")));
        report.Plugins.Should().BeEmpty();
        report.Failures.Should().BeEmpty();
        report.AllModules.Should().BeEmpty();
    }

    [Fact]
    public void Empty_Directory_Returns_Empty()
    {
        var dir = Path.Combine(Path.GetTempPath(), "grex365-plugins-empty-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var report = PluginLoader.LoadFrom(dir);
            report.Plugins.Should().BeEmpty();
            report.Failures.Should().BeEmpty();
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void Corrupt_Dll_Is_Reported_AsFailure_NotThrown()
    {
        var dir = Path.Combine(Path.GetTempPath(), "grex365-plugins-bad-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "garbage.dll");
        File.WriteAllBytes(path, new byte[] { 0x47, 0x52, 0x45, 0x58 });
        try
        {
            var report = PluginLoader.LoadFrom(dir);
            report.Plugins.Should().BeEmpty();
            report.Failures.Should().HaveCount(1);
            report.Failures[0].AssemblyPath.Should().Be(path);
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void Whitespace_Directory_Returns_Empty_NoThrow()
    {
        var report = PluginLoader.LoadFrom("   ");
        report.Plugins.Should().BeEmpty();
        report.Failures.Should().BeEmpty();
    }
}
