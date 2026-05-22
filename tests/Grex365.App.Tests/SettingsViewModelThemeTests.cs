using FluentAssertions;
using Grex365.App.ViewModels;
using Grex365.Core.Abstractions;

namespace Grex365.App.Tests;

public class SettingsViewModelThemeTests : IDisposable
{
    private readonly ISystemThemeProvider? _originalProvider;

    public SettingsViewModelThemeTests()
    {
        _originalProvider = SettingsViewModel.SystemThemeProvider;
    }

    public void Dispose()
    {
        SettingsViewModel.SystemThemeProvider = _originalProvider;
    }

    private sealed class StubProvider : ISystemThemeProvider
    {
        public bool Dark { get; set; } = true;
        public bool IsDarkTheme() => Dark;
    }

    [Fact]
    public void ResolveActualTheme_Dark_ReturnsDark()
    {
        SettingsViewModel.ResolveActualTheme("Dark").Should().Be("Dark");
    }

    [Fact]
    public void ResolveActualTheme_Light_ReturnsLight()
    {
        SettingsViewModel.ResolveActualTheme("Light").Should().Be("Light");
    }

    [Fact]
    public void ResolveActualTheme_Auto_WhenSystemDark_ReturnsDark()
    {
        SettingsViewModel.SystemThemeProvider = new StubProvider { Dark = true };
        SettingsViewModel.ResolveActualTheme("Auto").Should().Be("Dark");
    }

    [Fact]
    public void ResolveActualTheme_Auto_WhenSystemLight_ReturnsLight()
    {
        SettingsViewModel.SystemThemeProvider = new StubProvider { Dark = false };
        SettingsViewModel.ResolveActualTheme("Auto").Should().Be("Light");
    }

    [Fact]
    public void ResolveActualTheme_Auto_WithNullProvider_DefaultsToDark()
    {
        SettingsViewModel.SystemThemeProvider = null;
        SettingsViewModel.ResolveActualTheme("Auto").Should().Be("Dark");
    }

    [Fact]
    public void ResolveActualTheme_Null_DefaultsToDark()
    {
        SettingsViewModel.ResolveActualTheme(null).Should().Be("Dark");
    }

    [Fact]
    public void ResolveActualTheme_Unknown_DefaultsToDark()
    {
        SettingsViewModel.ResolveActualTheme("Sepia").Should().Be("Dark");
    }

    [Fact]
    public void ResolveActualTheme_CaseInsensitive()
    {
        SettingsViewModel.SystemThemeProvider = new StubProvider { Dark = false };
        SettingsViewModel.ResolveActualTheme("auto").Should().Be("Light");
        SettingsViewModel.ResolveActualTheme("LIGHT").Should().Be("Light");
        SettingsViewModel.ResolveActualTheme("dark").Should().Be("Dark");
    }
}
