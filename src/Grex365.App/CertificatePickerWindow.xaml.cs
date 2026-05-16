using System.Security.Cryptography.X509Certificates;
using System.Windows;
using Wpf.Ui.Controls;

namespace Grex365.App;

public sealed record CertEntry(string Subject, string Thumbprint, DateTime NotAfter, string Issuer);

public partial class CertificatePickerWindow : FluentWindow
{
    public CertEntry? Selected { get; private set; }

    public CertificatePickerWindow()
    {
        InitializeComponent();
        LoadCerts();
    }

    private void LoadCerts()
    {
        try
        {
            using var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadOnly);
            var entries = store.Certificates
                .OfType<X509Certificate2>()
                .Select(c => new CertEntry(c.Subject, c.Thumbprint, c.NotAfter, c.Issuer))
                .OrderByDescending(e => e.NotAfter)
                .ToList();
            CertList.ItemsSource = entries;
        }
        catch (System.Exception ex)
        {
            System.Windows.MessageBox.Show(ex.Message, "Error al leer Cert:\\CurrentUser\\My",
                System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
        }
    }

    private void Select_Click(object sender, RoutedEventArgs e)
    {
        Selected = CertList.SelectedItem as CertEntry;
        if (Selected is null)
        {
            return;
        }
        DialogResult = true;
        Close();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
