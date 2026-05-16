using System.Reflection;
using System.Runtime.Loader;
using Grex365.Core.Models;

namespace Grex365.Core.Plugins;

public static class PluginLoader
{
    public static PluginLoadReport LoadFrom(string directory, IProgress<LogEntry>? progress = null)
    {
        var plugins = new List<DiscoveredPlugin>();
        var failures = new List<PluginLoadFailure>();

        if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
        {
            return new PluginLoadReport(plugins, failures);
        }

        var dlls = Directory.EnumerateFiles(directory, "*.dll", SearchOption.TopDirectoryOnly);
        foreach (var dll in dlls)
        {
            try
            {
                var ctx = new AssemblyLoadContext(name: Path.GetFileNameWithoutExtension(dll), isCollectible: false);
                Assembly asm;
                using (var stream = File.OpenRead(dll))
                {
                    asm = ctx.LoadFromStream(stream);
                }

                var modules = DiscoverModules(asm);
                plugins.Add(new DiscoveredPlugin(dll, asm.GetName().Name ?? Path.GetFileName(dll), modules));
                progress?.Report(LogEntry.Ok("Plugins", $"Cargado {Path.GetFileName(dll)} ({modules.Count} módulo(s))"));
            }
            catch (Exception ex)
            {
                failures.Add(new PluginLoadFailure(dll, ex.Message));
                progress?.Report(LogEntry.Warn("Plugins", $"Plugin descartado {Path.GetFileName(dll)}: {ex.Message}"));
            }
        }

        return new PluginLoadReport(plugins, failures);
    }

    private static IReadOnlyList<IModule> DiscoverModules(Assembly assembly)
    {
        var moduleType = typeof(IModule);
        var found = new List<IModule>();
        Type[] types;
        try
        {
            types = assembly.GetTypes();
        }
        catch (ReflectionTypeLoadException ex)
        {
            types = ex.Types.Where(t => t is not null).Cast<Type>().ToArray();
        }

        foreach (var t in types)
        {
            if (t.IsAbstract || t.IsInterface)
            {
                continue;
            }
            if (!moduleType.IsAssignableFrom(t))
            {
                continue;
            }
            var ctor = t.GetConstructor(Type.EmptyTypes);
            if (ctor is null)
            {
                continue;
            }
            try
            {
                if (Activator.CreateInstance(t) is IModule module)
                {
                    found.Add(module);
                }
            }
            catch
            {
                // skip modules that fail to instantiate
            }
        }

        return found;
    }
}
