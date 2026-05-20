using System.Windows;
using System.Windows.Controls;
using Grex365.App.ViewModels;

namespace Grex365.App.Views;

public partial class CertWizardView : UserControl
{
    public CertWizardView()
    {
        InitializeComponent();
    }

    private async void ExportPfx_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is not CertWizardViewModel vm)
        {
            return;
        }
        var password = PfxPasswordBox.Password;
        try
        {
            await vm.ExportPfxAsync(password);
        }
        finally
        {
            PfxPasswordBox.Clear();
        }
    }
}
