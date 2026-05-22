using Grex365.Core.Abstractions;
using Microsoft.Win32;

namespace Grex365.App.Services;

public sealed class WindowsRegistryThemeProvider : ISystemThemeProvider
{
    // HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize
    //   AppsUseLightTheme  (DWORD)  0 = dark, 1 = light
    private const string KeyPath = @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";
    private const string ValueName = "AppsUseLightTheme";

    public bool IsDarkTheme()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(KeyPath);
            var v = key?.GetValue(ValueName);
            if (v is int i) return i == 0;
        }
        catch
        {
            // Registry access can fail in restricted shells; default to dark.
        }
        return true;
    }
}
