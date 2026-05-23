using FluentAssertions;
using Grex365.App;

namespace Grex365.App.Tests;

public class WindowPlacementGuardTests
{
    // Standard 1920x1080 primary monitor at origin.
    private static readonly VirtualScreenBounds Single = new(Left: 0, Top: 0, Width: 1920, Height: 1080);

    // Dual monitor setup: primary at (0,0) + secondary at (-1920,0). Virt bounds span both.
    private static readonly VirtualScreenBounds Dual = new(Left: -1920, Top: 0, Width: 3840, Height: 1080);

    [Fact]
    public void Centered_OnSingleMonitor_IsOnScreen()
    {
        WindowPlacementGuard.IsOnScreen(left: 400, top: 200, width: 1280, height: 720, Single)
            .Should().BeTrue();
    }

    [Fact]
    public void TopLeft_OnSingleMonitor_IsOnScreen()
    {
        WindowPlacementGuard.IsOnScreen(left: 0, top: 0, width: 1280, height: 720, Single).Should().BeTrue();
    }

    [Fact]
    public void FullyOffRight_NotOnScreen()
    {
        WindowPlacementGuard.IsOnScreen(left: 2000, top: 100, width: 800, height: 600, Single)
            .Should().BeFalse();
    }

    [Fact]
    public void FullyOffBottom_NotOnScreen()
    {
        WindowPlacementGuard.IsOnScreen(left: 100, top: 1200, width: 800, height: 600, Single)
            .Should().BeFalse();
    }

    [Fact]
    public void FullyOffLeft_NotOnScreen()
    {
        WindowPlacementGuard.IsOnScreen(left: -2000, top: 100, width: 800, height: 600, Single)
            .Should().BeFalse();
    }

    [Fact]
    public void FullyOffTop_NotOnScreen()
    {
        WindowPlacementGuard.IsOnScreen(left: 100, top: -1000, width: 800, height: 600, Single)
            .Should().BeFalse();
    }

    [Fact]
    public void OnSecondaryMonitor_DualSetup_IsOnScreen()
    {
        // Secondary at (-1920, 0); place window there at (-1820, 100).
        WindowPlacementGuard.IsOnScreen(left: -1820, top: 100, width: 800, height: 600, Dual)
            .Should().BeTrue();
    }

    [Fact]
    public void SecondaryDisconnected_PreviouslyValid_NowOffScreen()
    {
        // Secondary at -1920 in Dual now disconnected — bounds collapse to Single.
        // Previously valid window at (-1820, 100) now off-screen.
        WindowPlacementGuard.IsOnScreen(left: -1820, top: 100, width: 800, height: 600, Single)
            .Should().BeFalse();
    }

    [Fact]
    public void PartiallyVisible_AtLeast50px_IsOnScreen()
    {
        // Window slides off right edge but 51px of width still visible.
        WindowPlacementGuard.IsOnScreen(left: 1869, top: 100, width: 800, height: 600, Single)
            .Should().BeTrue(because: "1869+800=2669 > virt.Right=1920? No, but 1869 < 1920-50=1870 ✓ + right edge 2669 > 0+50=50 ✓");
    }

    [Fact]
    public void EdgeCase_ExactlyAtMargin_NotOnScreen()
    {
        // left = virt.Left + virt.Width - 50 (exactly at margin) → false (strict <)
        WindowPlacementGuard.IsOnScreen(left: 1870, top: 100, width: 800, height: 600, Single)
            .Should().BeFalse();
    }

    [Fact]
    public void Negative_Top_AboveScreen_NotOnScreen()
    {
        WindowPlacementGuard.IsOnScreen(left: 100, top: -700, width: 800, height: 600, Single)
            .Should().BeFalse();
    }
}
