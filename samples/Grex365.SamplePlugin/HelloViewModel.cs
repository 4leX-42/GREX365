using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace Grex365.SamplePlugin;

public sealed partial class HelloViewModel : ObservableObject
{
    private readonly HelloService _service;

    [ObservableProperty] private string? _message;

    public HelloViewModel(HelloService service)
    {
        _service = service;
        Message = _service.GetMessage();
    }

    [RelayCommand]
    private void Refresh() => Message = _service.GetMessage();
}
