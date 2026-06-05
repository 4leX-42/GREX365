using System.Windows.Controls;
using Grex365.App.ViewModels;

namespace Grex365.App.Views;

public partial class DashboardView : UserControl
{
    public DashboardView()
    {
        InitializeComponent();
        // Refresca la tarjeta de actividad cada vez que el Dashboard se muestra.
        Loaded += async (_, _) =>
        {
            if (DataContext is DashboardViewModel vm)
            {
                await vm.RefreshActivityCommand.ExecuteAsync(null);
            }
        };
    }
}
