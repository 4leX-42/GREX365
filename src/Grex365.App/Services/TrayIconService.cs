using D = System.Drawing;
using D2 = System.Drawing.Drawing2D;
using WinForms = System.Windows.Forms;

namespace Grex365.App.Services;

public sealed class TrayIconService : IDisposable
{
    private readonly WinForms.NotifyIcon _icon;
    private readonly WinForms.ToolStripMenuItem _connectionStateItem;
    private bool _disposed;

    public event EventHandler? OpenRequested;
    public event EventHandler? ReconnectRequested;
    public event EventHandler? ExitRequested;

    public TrayIconService()
    {
        _icon = new WinForms.NotifyIcon
        {
            Visible = false,
            Text = "GREX365",
            Icon = BuildIcon(false, false),
        };
        _icon.DoubleClick += (s, e) => OpenRequested?.Invoke(s, e);

        var menu = new WinForms.ContextMenuStrip();
        var openItem = new WinForms.ToolStripMenuItem("Abrir GREX365");
        openItem.Font = new D.Font(openItem.Font, D.FontStyle.Bold);
        openItem.Click += (s, e) => OpenRequested?.Invoke(s, e);

        _connectionStateItem = new WinForms.ToolStripMenuItem("Estado: desconectado") { Enabled = false };

        var reconnectItem = new WinForms.ToolStripMenuItem("Reconectar ahora");
        reconnectItem.Click += (s, e) => ReconnectRequested?.Invoke(s, e);

        var exitItem = new WinForms.ToolStripMenuItem("Salir");
        exitItem.Click += (s, e) => ExitRequested?.Invoke(s, e);

        menu.Items.Add(openItem);
        menu.Items.Add(_connectionStateItem);
        menu.Items.Add(new WinForms.ToolStripSeparator());
        menu.Items.Add(reconnectItem);
        menu.Items.Add(new WinForms.ToolStripSeparator());
        menu.Items.Add(exitItem);

        _icon.ContextMenuStrip = menu;
    }

    public void Show()
    {
        _icon.Visible = true;
    }

    public void Hide()
    {
        _icon.Visible = false;
    }

    public void UpdateConnectionState(bool graphConnected, bool exchangeConnected)
    {
        var label = (graphConnected, exchangeConnected) switch
        {
            (true, true) => "Estado: Graph + Exchange",
            (true, false) => "Estado: Graph (sin EXO)",
            (false, true) => "Estado: Exchange (sin Graph)",
            _ => "Estado: desconectado",
        };
        _connectionStateItem.Text = label;

        var oldIcon = _icon.Icon;
        _icon.Icon = BuildIcon(graphConnected, exchangeConnected);
        oldIcon?.Dispose();
    }

    public void ShowBalloon(string title, string message)
    {
        _icon.BalloonTipTitle = title;
        _icon.BalloonTipText = message;
        _icon.ShowBalloonTip(3500);
    }

    private static D.Icon BuildIcon(bool graphConnected, bool exchangeConnected)
    {
        using var bmp = new D.Bitmap(32, 32);
        using var g = D.Graphics.FromImage(bmp);
        g.SmoothingMode = D2.SmoothingMode.AntiAlias;
        g.InterpolationMode = D2.InterpolationMode.HighQualityBicubic;
        g.PixelOffsetMode = D2.PixelOffsetMode.HighQuality;

        var rect = new D.Rectangle(0, 0, 32, 32);
        using var bg = new D2.LinearGradientBrush(rect,
            D.Color.FromArgb(0x1E, 0x40, 0xAF),
            D.Color.FromArgb(0x3B, 0x82, 0xF6),
            45f);
        g.FillRectangle(bg, rect);

        using var font = new D.Font("Segoe UI", 16f, D.FontStyle.Bold, D.GraphicsUnit.Pixel);
        var fmt = new D.StringFormat
        {
            Alignment = D.StringAlignment.Center,
            LineAlignment = D.StringAlignment.Center,
        };
        g.DrawString("G", font, D.Brushes.White, new D.RectangleF(0, -1, 32, 32), fmt);

        if (!graphConnected || !exchangeConnected)
        {
            var dotColor = (!graphConnected && !exchangeConnected)
                ? D.Color.FromArgb(0xF8, 0x71, 0x71)
                : D.Color.FromArgb(0xFB, 0xBF, 0x24);
            using var dot = new D.SolidBrush(dotColor);
            g.FillEllipse(dot, new D.Rectangle(22, 2, 9, 9));
        }

        var hicon = bmp.GetHicon();
        var icon = D.Icon.FromHandle(hicon);
        return (D.Icon)icon.Clone();
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _icon.Visible = false;
        _icon.ContextMenuStrip?.Dispose();
        _icon.Icon?.Dispose();
        _icon.Dispose();
    }
}
