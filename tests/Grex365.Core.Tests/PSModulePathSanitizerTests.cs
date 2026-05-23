using FluentAssertions;
using Grex365.PowerShell;

namespace Grex365.Core.Tests;

public class PSModulePathSanitizerTests
{
    private const string UserPath = @"C:\Users\test\Documents\PowerShell\Modules";
    private static readonly char Sep = System.IO.Path.PathSeparator;

    [Fact]
    public void Sanitize_EmptyInput_ReturnsSafeDefault_WithUserPath()
    {
        var result = PSModulePathSanitizer.Sanitize(string.Empty, UserPath);
        result.Should().StartWith(UserPath);
    }

    [Fact]
    public void Sanitize_WindowsAppsPath_Removed()
    {
        var input = $@"C:\Program Files\WindowsApps\Microsoft.PowerShell_7.5.1.0_x64{Sep}{UserPath}";
        var result = PSModulePathSanitizer.Sanitize(input, UserPath);
        result.Should().NotContain("WindowsApps");
    }

    [Fact]
    public void Sanitize_MixedPaths_OnlyWindowsAppsDropped()
    {
        var prog = @"C:\Program Files\PowerShell\Modules";
        var winApps = @"C:\Program Files\WindowsApps\Bad";
        var sys = @"C:\Windows\System32\WindowsPowerShell\v1.0\Modules";
        var input = string.Join(Sep, [UserPath, prog, winApps, sys]);

        var result = PSModulePathSanitizer.Sanitize(input, UserPath);

        result.Should().Contain(UserPath);
        result.Should().Contain(prog);
        result.Should().Contain(sys);
        result.Should().NotContain("WindowsApps");
    }

    [Fact]
    public void Sanitize_UserPath_Missing_IsPrependedFirst()
    {
        var prog = @"C:\Program Files\PowerShell\Modules";
        var input = prog;

        var result = PSModulePathSanitizer.Sanitize(input, UserPath);

        result.Should().StartWith(UserPath);
        result.Should().Contain(prog);
    }

    [Fact]
    public void Sanitize_UserPath_AlreadyPresent_NotDuplicated()
    {
        var prog = @"C:\Program Files\PowerShell\Modules";
        var input = $"{UserPath}{Sep}{prog}";

        var result = PSModulePathSanitizer.Sanitize(input, UserPath);

        // Count occurrences of UserPath: should be 1.
        var parts = result.Split(Sep, StringSplitOptions.RemoveEmptyEntries);
        parts.Count(p => string.Equals(p, UserPath, StringComparison.OrdinalIgnoreCase)).Should().Be(1);
    }

    [Fact]
    public void Sanitize_UserPath_CaseInsensitiveMatch_NoDuplicate()
    {
        var prog = @"C:\Program Files\PowerShell\Modules";
        var upperUser = UserPath.ToUpperInvariant();
        var input = $"{upperUser}{Sep}{prog}";

        var result = PSModulePathSanitizer.Sanitize(input, UserPath);

        var parts = result.Split(Sep, StringSplitOptions.RemoveEmptyEntries);
        parts.Count(p => string.Equals(p, UserPath, StringComparison.OrdinalIgnoreCase)).Should().Be(1,
            because: "case-insensitive comparison should treat upper/mixed as already present");
    }

    [Fact]
    public void Sanitize_EmptyParts_FilteredOut()
    {
        var input = $"{Sep}{Sep}{UserPath}{Sep}{Sep}";
        var result = PSModulePathSanitizer.Sanitize(input, UserPath);
        var parts = result.Split(Sep, StringSplitOptions.RemoveEmptyEntries);
        parts.Should().NotContain(string.Empty);
        parts.Should().Contain(UserPath);
    }

    [Fact]
    public void BuildSafeDefault_HasUserPathFirst()
    {
        var result = PSModulePathSanitizer.BuildSafeDefault(UserPath);
        result.Should().StartWith(UserPath);
    }

    [Fact]
    public void BuildSafeDefault_IncludesProgramFilesAndSystem32()
    {
        var result = PSModulePathSanitizer.BuildSafeDefault(UserPath);
        var parts = result.Split(Sep, StringSplitOptions.RemoveEmptyEntries);
        parts.Should().HaveCount(3);
        parts.Should().Contain(UserPath);
        parts.Should().Contain(p => p.Contains("PowerShell") && p.Contains("Modules"));
        parts.Should().Contain(p => p.Contains("WindowsPowerShell"));
    }

    [Fact]
    public void ResolveUserModulesPath_ContainsExpectedSegments()
    {
        var path = PSModulePathSanitizer.ResolveUserModulesPath();
        path.Should().EndWith(System.IO.Path.Combine("PowerShell", "Modules"));
    }
}
