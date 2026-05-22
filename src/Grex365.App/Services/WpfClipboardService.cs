using System.Windows;
using Grex365.Core.Abstractions;

namespace Grex365.App.Services;

public sealed class WpfClipboardService : IClipboardService
{
    public void SetText(string text)
    {
        try
        {
            Clipboard.SetText(text ?? string.Empty);
        }
        catch
        {
            // Clipboard occasionally throws under terminal services or restricted shells; swallow.
        }
    }
}
