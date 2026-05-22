using System.Windows;
using Grex365.App.ViewModels;
using Wpf.Ui.Controls;

namespace Grex365.App;

public partial class FirstRunWizardWindow : FluentWindow
{
    private readonly FirstRunWizardViewModel _vm;

    public FirstRunWizardWindow(FirstRunWizardViewModel viewModel)
    {
        _vm = viewModel;
        DataContext = viewModel;
        InitializeComponent();
        viewModel.PropertyChanged += OnVmPropertyChanged;
    }

    private void OnVmPropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(FirstRunWizardViewModel.Completed) && _vm.Completed)
        {
            Close();
        }
    }

    private async void Close_Click(object sender, RoutedEventArgs e)
    {
        await _vm.SkipCommand.ExecuteAsync(null);
        // Window closes via Completed property handler.
    }

    private async void Finish_Click(object sender, RoutedEventArgs e)
    {
        await _vm.FinishCommand.ExecuteAsync(null);
    }
}
