using Microsoft.Extensions.DependencyInjection;

namespace Grex365.Core.Plugins;

public interface IModule
{
    string Title { get; }

    string Glyph { get; }

    Type ViewModelType { get; }

    Type ViewType { get; }

    void RegisterServices(IServiceCollection services);
}

public sealed record DiscoveredPlugin(
    string AssemblyPath,
    string AssemblyName,
    IReadOnlyList<IModule> Modules);

public sealed record PluginLoadFailure(string AssemblyPath, string Message);

public sealed record PluginLoadReport(
    IReadOnlyList<DiscoveredPlugin> Plugins,
    IReadOnlyList<PluginLoadFailure> Failures)
{
    public IEnumerable<IModule> AllModules => Plugins.SelectMany(p => p.Modules);
}
