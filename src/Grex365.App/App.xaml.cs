using System.IO;
using System.Threading.Tasks;
using System.Windows;
using Grex365.App.Services;
using Grex365.App.ViewModels;
using Grex365.Core.Abstractions;
using Grex365.Core.Connections;
using Grex365.Core.Preferences;
using Grex365.PowerShell;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Serilog;
using Serilog.Extensions.Logging;

namespace Grex365.App;

public partial class App : Application
{
    private IHost? _host;

    public static IServiceProvider Services => ((App)Current)._host!.Services;

    public static string DataDirectory { get; } =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Grex365");

    protected override void OnStartup(StartupEventArgs e)
    {
        Directory.CreateDirectory(DataDirectory);
        var logsDir = Path.Combine(DataDirectory, "logs");
        var configDir = Path.Combine(DataDirectory, "config");
        Directory.CreateDirectory(logsDir);
        Directory.CreateDirectory(configDir);

        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .WriteTo.File(
                path: Path.Combine(logsDir, "grex365-.log"),
                rollingInterval: RollingInterval.Day,
                retainedFileCountLimit: 30,
                outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
            .CreateLogger();

        WireGlobalExceptionHandlers();
        TryImportLegacyConfig(configDir);

        _host = Host.CreateDefaultBuilder()
            .ConfigureServices(services =>
            {
                services.AddSingleton<ILoggerFactory>(new SerilogLoggerFactory(Log.Logger, dispose: true));
                services.AddSingleton(typeof(ILogger<>), typeof(Logger<>));

                services.AddSingleton(_ => new RunspacePoolHost(minRunspaces: 1, maxRunspaces: 4));
                services.AddSingleton<IPowerShellRunner, PowerShellRunner>();

                services.AddSingleton<IGraphConnection, GraphConnection>();
                services.AddSingleton<IExchangeConnection, ExchangeConnection>();
                services.AddSingleton<IConnectionStateMonitor, ConnectionStateMonitor>();
                services.AddSingleton<ICertValidator, CertValidator>();
                services.AddSingleton<ITenantLock, TenantLock>();

                services.AddSingleton<IPreferencesStore>(_ => new JsonPreferencesStore(configDir));
                services.AddSingleton<ICertConfigStore>(_ => new JsonCertConfigStore(configDir));

                services.AddSingleton<IUiLogSink, UiLogSink>();

                services.AddTransient<ConnectViewModel>();
                services.AddTransient<SettingsViewModel>();
                services.AddTransient<MainViewModel>();
                services.AddSingleton<MainWindow>();
                services.AddTransient<SettingsWindow>();
            })
            .Build();

        var monitor = Services.GetRequiredService<IConnectionStateMonitor>();
        monitor.Start();

        var main = Services.GetRequiredService<MainWindow>();
        main.Show();

        base.OnStartup(e);
    }

    private void WireGlobalExceptionHandlers()
    {
        DispatcherUnhandledException += (_, args) =>
        {
            Log.Error(args.Exception, "Unhandled UI exception");
            MessageBox.Show(args.Exception.Message, "Grex365 — Error", MessageBoxButton.OK, MessageBoxImage.Error);
            args.Handled = true;
        };

        AppDomain.CurrentDomain.UnhandledException += (_, args) =>
        {
            if (args.ExceptionObject is Exception ex)
            {
                Log.Fatal(ex, "Unhandled AppDomain exception (terminating={Terminating})", args.IsTerminating);
            }
        };

        TaskScheduler.UnobservedTaskException += (_, args) =>
        {
            Log.Error(args.Exception, "Unobserved task exception");
            args.SetObserved();
        };
    }

    private static void TryImportLegacyConfig(string targetConfigDir)
    {
        try
        {
            var exeDir = AppContext.BaseDirectory;
            var candidates = new List<string>
            {
                Path.Combine(exeDir, "..", "..", "..", "..", "..", "GREX365", "config"),
                Path.Combine(exeDir, "GREX365", "config"),
                Path.Combine(Directory.GetCurrentDirectory(), "GREX365", "config")
            };

            var importer = new LegacyPreferencesImporter(targetConfigDir);
            var result = importer.TryImportAsync(candidates).GetAwaiter().GetResult();
            if (result.PreferencesImported || result.CertConfigImported)
            {
                Log.Information(
                    "Imported legacy config (prefs={Prefs}, cert={Cert})",
                    result.PreferencesImported,
                    result.CertConfigImported);
            }
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Legacy config import failed (non-fatal)");
        }
    }

    protected override async void OnExit(ExitEventArgs e)
    {
        if (_host is not null)
        {
            var monitor = Services.GetService<IConnectionStateMonitor>();
            if (monitor is not null)
            {
                await monitor.DisposeAsync().ConfigureAwait(false);
            }

            var pool = Services.GetService<RunspacePoolHost>();
            pool?.Dispose();

            _host.Dispose();
        }
        Log.CloseAndFlush();
        base.OnExit(e);
    }
}
