using FluentAssertions;
using Grex365.PowerShell;

namespace Grex365.Core.Tests;

public class ExeResolverTests
{
    private static readonly char Sep = System.IO.Path.PathSeparator;

    [Fact]
    public void Resolve_FirstCandidateExists_ReturnsIt()
    {
        var path = $@"C:\Windows{Sep}C:\Tools";
        var exists = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { @"C:\Tools\pwsh.exe" };

        var result = ExeResolver.Resolve(
            candidateNames: ["pwsh.exe", "powershell.exe"],
            pathEnv: path,
            fileExists: exists.Contains,
            fallback: "fb.exe");

        result.Should().Be(@"C:\Tools\pwsh.exe");
    }

    [Fact]
    public void Resolve_SecondCandidate_WhenFirstMissing()
    {
        var path = $@"C:\Windows{Sep}C:\Legacy";
        var exists = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { @"C:\Legacy\powershell.exe" };

        var result = ExeResolver.Resolve(
            candidateNames: ["pwsh.exe", "powershell.exe"],
            pathEnv: path,
            fileExists: exists.Contains,
            fallback: "fb.exe");

        result.Should().Be(@"C:\Legacy\powershell.exe");
    }

    [Fact]
    public void Resolve_NoCandidateExists_ReturnsFallback()
    {
        var path = $@"C:\Windows{Sep}C:\Tools";
        var exists = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var result = ExeResolver.Resolve(
            candidateNames: ["pwsh.exe"],
            pathEnv: path,
            fileExists: exists.Contains,
            fallback: "fallback.exe");

        result.Should().Be("fallback.exe");
    }

    [Fact]
    public void Resolve_NullPath_ReturnsFallback()
    {
        var result = ExeResolver.Resolve(
            candidateNames: ["pwsh.exe"],
            pathEnv: null,
            fileExists: _ => true,  // would match if it got that far
            fallback: "fb.exe");

        result.Should().Be("fb.exe");
    }

    [Fact]
    public void Resolve_EmptyPath_ReturnsFallback()
    {
        var result = ExeResolver.Resolve(
            candidateNames: ["pwsh.exe"],
            pathEnv: string.Empty,
            fileExists: _ => true,
            fallback: "fb.exe");

        result.Should().Be("fb.exe");
    }

    [Fact]
    public void Resolve_EmptySegmentsInPath_Skipped()
    {
        var path = $@"{Sep}{Sep}C:\Tools{Sep}";
        var exists = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { @"C:\Tools\pwsh.exe" };

        var result = ExeResolver.Resolve(
            candidateNames: ["pwsh.exe"],
            pathEnv: path,
            fileExists: exists.Contains,
            fallback: "fb.exe");

        result.Should().Be(@"C:\Tools\pwsh.exe");
    }

    [Fact]
    public void Resolve_CandidateOrderRespected_PwshFirstThenLegacy()
    {
        // Both exist in different paths; first candidate found in PATH wins.
        var path = $@"C:\Modern{Sep}C:\Legacy";
        var exists = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            @"C:\Modern\pwsh.exe",
            @"C:\Legacy\powershell.exe",
        };

        var result = ExeResolver.Resolve(
            candidateNames: ["pwsh.exe", "powershell.exe"],
            pathEnv: path,
            fileExists: exists.Contains,
            fallback: "fb.exe");

        result.Should().Be(@"C:\Modern\pwsh.exe");
    }

    [Fact]
    public void Resolve_SecondCandidateFound_WhenFirstNotInPath()
    {
        // pwsh.exe not in any segment; powershell.exe in second.
        var path = $@"C:\A{Sep}C:\B";
        var exists = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            @"C:\B\powershell.exe",
        };

        var result = ExeResolver.Resolve(
            candidateNames: ["pwsh.exe", "powershell.exe"],
            pathEnv: path,
            fileExists: exists.Contains,
            fallback: "fb.exe");

        result.Should().Be(@"C:\B\powershell.exe");
    }

    [Fact]
    public void Resolve_NullCandidates_Throws()
    {
        Action act = () => ExeResolver.Resolve(null!, "anything", _ => true, "fb.exe");
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void Resolve_NullFileExists_Throws()
    {
        Action act = () => ExeResolver.Resolve(["pwsh.exe"], "anything", null!, "fb.exe");
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void Resolve_NullFallback_Throws()
    {
        Action act = () => ExeResolver.Resolve(["pwsh.exe"], "anything", _ => true, null!);
        act.Should().Throw<ArgumentNullException>();
    }
}
