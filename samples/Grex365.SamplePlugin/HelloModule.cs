using Grex365.Core.Plugins;
using Microsoft.Extensions.DependencyInjection;

namespace Grex365.SamplePlugin;

public sealed class HelloModule : IModule
{
    public string Title => "Sample Hello";

    public string Glyph => ""; // Segoe Fluent: Page

    public Type ViewModelType => typeof(HelloViewModel);

    public Type ViewType => typeof(HelloView);

    public void RegisterServices(IServiceCollection services)
    {
        services.AddSingleton<HelloService>();
    }
}

public sealed class HelloService
{
    private int _counter;

    public string GetMessage()
    {
        _counter++;
        return $"Hola desde Grex365.SamplePlugin · invocacion #{_counter} · {DateTime.Now:HH:mm:ss}";
    }
}
