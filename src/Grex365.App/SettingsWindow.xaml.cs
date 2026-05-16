using System.Windows;
using Grex365.App.ViewModels;
using Wpf.Ui.Controls;

namespace Grex365.App;

public partial class SettingsWindow : FluentWindow
{
    public SettingsWindow(SettingsViewModel viewModel)
    {
        InitializeComponent();
        DataContext = viewModel;
        Loaded += async (_, _) => await viewModel.LoadCommand.ExecuteAsync(null);
    }

    private void BrowseCert_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is not SettingsViewModel vm)
        {
            return;
        }
        var dlg = new CertificatePickerWindow
        {
            Owner = this
        };
        if (dlg.ShowDialog() == true && dlg.Selected is not null)
        {
            vm.CertThumbprint = dlg.Selected.Thumbprint;
            vm.ValidateCertCommand.Execute(null);
        }
    }
}
