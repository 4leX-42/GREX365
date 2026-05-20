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
        report.Disabled.Should().BeEmpty();
    }

    [Fact]
    public void Disabled_Assembly_Is_Skipped_Without_Attempting_Load()
    {
        var dir = Path.Combine(Path.GetTempPath(), "grex365-plugins-disabled-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        // Even garbage bytes — if the loader were invoked, it would land in Failures.
        var garbage = Path.Combine(dir, "muted.dll");
        File.WriteAllBytes(garbage, new byte[] { 0x47, 0x52, 0x45, 0x58 });
        try
        {
            var report = PluginLoader.LoadFrom(dir, disabledAssemblies: new[] { "muted.dll" });
            report.Plugins.Should().BeEmpty();
            report.Failures.Should().BeEmpty();
            report.Disabled.Should().HaveCount(1);
            report.Disabled[0].AssemblyFileName.Should().Be("muted.dll");
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void Disabled_Match_Is_Case_Insensitive()
    {
        var dir = Path.Combine(Path.GetTempPath(), "grex365-plugins-case-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var garbage = Path.Combine(dir, "MyPlugin.dll");
        File.WriteAllBytes(garbage, new byte[] { 0x47, 0x52, 0x45, 0x58 });
        try
        {
            var report = PluginLoader.LoadFrom(dir, disabledAssemblies: new[] { "myplugin.dll" });
            report.Plugins.Should().BeEmpty();
            report.Failures.Should().BeEmpty();
            report.Disabled.Should().ContainSingle().Which.AssemblyFileName.Should().Be("MyPlugin.dll");
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }
}
