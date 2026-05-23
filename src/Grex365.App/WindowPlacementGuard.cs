namespace Grex365.App;

public readonly record struct VirtualScreenBounds(double Left, double Top, double Width, double Height);

public static class WindowPlacementGuard
{
    private const double EdgeMargin = 50;

    // True iff the window rect (left/top/width/height) remains at least EdgeMargin
    // pixels visible inside the virtual screen bounds. Prevents restoring a window
    // off-screen after multi-monitor changes / RDP / disconnected secondary display.
    public static bool IsOnScreen(double left, double top, double width, double height, VirtualScreenBounds virt)
    {
        return left + width > virt.Left + EdgeMargin
            && left < virt.Left + virt.Width - EdgeMargin
            && top + height > virt.Top + EdgeMargin
            && top < virt.Top + virt.Height - EdgeMargin;
    }
}
