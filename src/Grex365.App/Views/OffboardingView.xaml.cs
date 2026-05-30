using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Grex365.App.ViewModels;

namespace Grex365.App.Views;

public partial class OffboardingView : UserControl
{
    public OffboardingView()
    {
        InitializeComponent();
    }

    // Ctrl+C copies the selected log rows (or all if none selected) as plain text.
    private void LogList_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.C || (Keyboard.Modifiers & ModifierKeys.Control) == 0)
        {
            return;
        }

        var rows = LogList.SelectedItems.Count > 0
            ? LogList.SelectedItems.Cast<object>()
            : LogList.Items.Cast<object>();

        var sb = new StringBuilder();
        foreach (var row in rows)
        {
            if (row is LogLine l)
            {
                sb.Append(l.Time).Append("  ").AppendLine(l.Text);
            }
            else if (row is not null)
            {
                sb.AppendLine(row.ToString());
            }
        }

        if (sb.Length > 0)
        {
            try { Clipboard.SetText(sb.ToString()); } catch { /* clipboard busy */ }
        }
        e.Handled = true;
    }
}
