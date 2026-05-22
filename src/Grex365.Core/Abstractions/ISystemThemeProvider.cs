namespace Grex365.Core.Abstractions;

public interface ISystemThemeProvider
{
    /// <summary>Returns true if the OS is currently using a dark theme for apps.</summary>
    bool IsDarkTheme();
}
