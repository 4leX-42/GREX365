using System.Diagnostics;
using System.IO;
using System.Windows;
using Wpf.Ui.Controls;

namespace Grex365.App;

public partial class AboutWindow : FluentWindow
{
    public AboutWindow()
    {
        InitializeComponent();
        Loaded += (_, _) =>
        {
            VersionText.Text = "v" + App.AppVersion;
            VersionDetailText.Text = App.AppVersion;
            RuntimeText.Text = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription;
            DataDirText.Text = App.DataDirectory;
        };
    }

    private void OpenDataDir_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (Directory.Exists(App.DataDirectory))
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = App.DataDirectory,
                    UseShellExecute = true
                });
                return;
            }
            ShowOpenFailed();
        }
        catch
        {
            ShowOpenFailed();
        }
    }

    private void ShowOpenFailed()
    {
        _ = new Wpf.Ui.Controls.MessageBox
        {
            Title = "GREX365",
            Content = L10n.Format("About.DataDir.OpenFailed", App.DataDirectory),
            CloseButtonText = L10n.Get("Dialog.Ok"),
        }.ShowDialogAsync();
    }

    private void Close_Click(object sender, RoutedEventArgs e) => Close();
}
