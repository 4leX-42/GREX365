using Grex365.App.Services;
using Grex365.App.ViewModels;
using Wpf.Ui.Controls;

namespace Grex365.App;

public partial class MainWindow : FluentWindow
{
    public MainWindow(MainViewModel viewModel, WpfUiNotifier notifier)
    {
        InitializeComponent();
        DataContext = viewModel;
        Loaded += (_, _) => notifier.AttachPresenter(RootSnackbarPresenter);
    }
}
