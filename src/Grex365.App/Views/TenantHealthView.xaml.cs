using System.Windows;
using System.Windows.Controls;
using Grex365.App.ViewModels;

namespace Grex365.App.Views;

public partial class TenantHealthView : UserControl
{
    public TenantHealthView()
    {
        InitializeComponent();
        Loaded += OnViewLoaded;
    }

    private void OnViewLoaded(object sender, RoutedEventArgs e)
    {
        if (DataContext is TenantHealthViewModel vm)
        {
            vm.TriggerLoadIfNeeded();
        }
    }
}
