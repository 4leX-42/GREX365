using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Grex365.App.Services;
using Grex365.App.ViewModels;
using Grex365.Core.Abstractions;
using Grex365.Core.Audit;
using Grex365.Core.Certificates;
using Grex365.Core.Connections;
using Grex365.Core.Models;
using Grex365.Core.DomainChecks;
using Grex365.Core.Groups;
using Grex365.Core.Health;
using Grex365.Core.Offboarding;
using Grex365.Core.Onboarding;
using Grex365.Core.Plugins;
using Grex365.Core.Preferences;
using Grex365.Core.Security;
using Grex365.Core.Users;
using Grex365.PowerShell;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Serilog;
using Serilog.Core;
using Serilog.Events;
using Serilog.Extensions.Logging;

namespace Grex365.App;

public partial class App : Application
{
    private IHost? _host;

    public static IServiceProvider Services => ((App)Current)._host!.Services;

    public static string DataDirectory { get; } =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Grex365");

    public static LoggingLevelSwitch LogLevelSwitch { get; } = new(LogEventLevel.Information);

    // Surfaced product version from Directory.Build.props (AssemblyInformationalVersion or fallback to FileVersion).
    public static string AppVersion { get; } = ResolveVersion();

    private static string ResolveVersion()
    {
        var asm = typeof(App).Assembly;
        var info = asm.GetCustomAttributes(typeof(System.Reflection.AssemblyInformationalVersionAttribute), false)
            .OfType<System.Reflection.AssemblyInformationalVersionAttribute>()
            .FirstOrDefault()?.InformationalVersion;
        if (!string.IsNullOrWhiteSpace(info))
        {
            // Strip build-metadata suffix introduced by SourceLink, if any.
            var plus = info.IndexOf('+');
            return plus >= 0 ? info.Substring(0, plus) : info;
        }
        var ver = asm.GetName().Version;
        return ver?.ToString(3) ?? "0.0.0";
    }

    protected override void OnStartup(StartupEventArgs e)
    {
        Directory.CreateDirectory(DataDirectory);
        var logsDir = Path.Combine(DataDirectory, "logs");
        var configDir = Path.Combine(DataDirectory, "config");
        Directory.CreateDirectory(logsDir);
        Directory.CreateDirectory(configDir);

        var bootPrefs = TryLoadBootPreferences(configDir);
        LogLevelSwitch.MinimumLevel = ParseLogLevel(bootPrefs.LogLevel);

        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.ControlledBy(LogLevelSwitch)
            .WriteTo.File(
                path: Path.Combine(logsDir, "grex365-.log"),
                rollingInterval: RollingInterval.Day,
                retainedFileCountLimit: 30,
                outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
            .CreateLogger();

        WireGlobalExceptionHandlers();
        TryImportLegacyConfig(configDir);

        var pluginsDir = Path.Combine(DataDirectory, "plugins");
        Directory.CreateDirectory(pluginsDir);

        var pluginReport = PluginLoader.LoadFrom(
            pluginsDir,
            disabledAssemblies: bootPrefs.DisabledPluginAssemblies);
        foreach (var failure in pluginReport.Failures)
        {
            Log.Warning("Plugin descartado {Path}: {Reason}", failure.AssemblyPath, failure.Message);
        }
        foreach (var plugin in pluginReport.Plugins)
        {
            Log.Information("Plugin {Asm} aporta {Count} módulo(s)", plugin.AssemblyName, plugin.Modules.Count);
        }
        foreach (var off in pluginReport.Disabled)
        {
            Log.Information("Plugin deshabilitado por preferencias: {File}", off.AssemblyFileName);
        }

        _host = Host.CreateDefaultBuilder()
            .ConfigureServices(services =>
            {
                services.AddSingleton<ILoggerFactory>(new SerilogLoggerFactory(Log.Logger, dispose: true));
                services.AddSingleton(typeof(ILogger<>), typeof(Logger<>));
                services.AddSingleton(LogLevelSwitch);

                services.AddSingleton(_ => new RunspacePoolHost(minRunspaces: 1, maxRunspaces: 4));
                services.AddSingleton<IPowerShellRunner, PowerShellRunner>();

                services.AddSingleton<IGraphConnection, GraphConnection>();
                services.AddSingleton<IExchangeConnection, ExchangeConnection>();
                services.AddSingleton<IConnectionStateMonitor, ConnectionStateMonitor>();
                services.AddSingleton<ICertValidator, CertValidator>();
                services.AddSingleton<ITenantLock, TenantLock>();
                services.AddSingleton<IGroupsService, GraphGroupsService>();
                services.AddSingleton<IDistributionListsService, DistributionListsService>();
                services.AddSingleton<ISharedMailboxService, SharedMailboxService>();
                services.AddSingleton<IMailboxRulesService, MailboxRulesService>();
                services.AddSingleton<IMailFlowRulesService, MailFlowRulesService>();
                services.AddSingleton<IAuditService, GraphAuditService>();
                services.AddSingleton<IExoForwardingAuditService, ExoForwardingAuditService>();
                services.AddSingleton<IAuditFindingsStore, InMemoryAuditFindingsStore>();
                services.AddSingleton<IUserDetailsHost, UserDetailsHost>();
                services.AddSingleton<ITenantHealthService, GraphTenantHealthService>();
                services.AddSingleton<IUsersService, GraphUsersService>();
                services.AddSingleton<IOffboardingService, OffboardingService>();
                services.AddSingleton<IOnboardingService, OnboardingService>();
                services.AddSingleton<ICertificateGenerator, SelfSignedCertificateGenerator>();
                services.AddSingleton<IAppRegistrationService, GraphAppRegistrationService>();
                services.AddSingleton<IDomainChecker, NslookupDomainChecker>();
                services.AddSingleton<IMembershipChecker, GraphMembershipChecker>();
                services.AddSingleton<IRbacGuard>(sp =>
                {
                    var checker = sp.GetRequiredService<IMembershipChecker>();
                    var prefsStore = sp.GetRequiredService<IPreferencesStore>();
                    return new RbacGuard(checker, () =>
                    {
                        try
                        {
                            return prefsStore.LoadAsync().GetAwaiter().GetResult().AuthorizationGroupId;
                        }
                        catch
                        {
                            return null;
                        }
                    });
                });

                services.AddSingleton<IPreferencesStore>(_ => new JsonPreferencesStore(configDir));
                services.AddSingleton<ICertConfigStore>(_ => new JsonCertConfigStore(configDir));

                services.AddSingleton<WpfUiNotifier>();
                services.AddSingleton<INotifier>(sp => sp.GetRequiredService<WpfUiNotifier>());
                services.AddSingleton<IDialogService, WpfDialogService>();
                services.AddSingleton<IClipboardService, WpfClipboardService>();
                services.AddSingleton<ISystemThemeProvider, WindowsRegistryThemeProvider>();

                var auditDir = Path.Combine(DataDirectory, "audit");
                Directory.CreateDirectory(auditDir);
                services.AddSingleton<IAuditLog>(_ => new FileAuditLog(auditDir));

                var aiConnectionString = bootPrefs.ApplicationInsightsConnectionString?.Trim();
                if (!string.IsNullOrWhiteSpace(aiConnectionString))
                {
                    try
                    {
                        var aiClient = new ApplicationInsightsTelemetry(aiConnectionString);
                        services.AddSingleton<ITelemetry>(aiClient);
                        Log.Information("Application Insights habilitado");
                    }
                    catch (Exception ex)
                    {
                        Log.Warning(ex, "No se pudo inicializar Application Insights, usando NullTelemetry");
                        services.AddSingleton<ITelemetry, NullTelemetry>();
                    }
                }
                else
                {
                    services.AddSingleton<ITelemetry, NullTelemetry>();
                }

                services.AddSingleton<IUiLogSink, UiLogSink>();

                // Page VMs are SINGLETON so state persists across nav changes
                // (search queries, selected item, last results, drafted input...).
                // SettingsViewModel + FirstRunWizardViewModel keep Transient because
                // they live in modal windows and should reset each time.
                services.AddSingleton<ConnectViewModel>();
                services.AddSingleton<DashboardViewModel>();
                services.AddSingleton<GroupsViewModel>();
                services.AddSingleton<SharedMailboxViewModel>();
                services.AddSingleton<AuditViewModel>();
                services.AddSingleton<TenantHealthViewModel>();
                services.AddSingleton<UsersViewModel>();
                services.AddSingleton<UserDetailsViewModel>();
                services.AddSingleton<OffboardingViewModel>();
                services.AddSingleton<OnboardingViewModel>();
                services.AddSingleton<MailboxRulesViewModel>();
                services.AddSingleton<MailFlowRulesViewModel>();
                services.AddSingleton<AuditLogViewModel>();
                services.AddSingleton<CertWizardViewModel>();
                services.AddSingleton<DomainCheckViewModel>();
                services.AddSingleton<PsConsoleViewModel>();
                services.AddTransient<SettingsViewModel>();
                services.AddSingleton<MainViewModel>();
                services.AddSingleton<MainWindow>();
                services.AddTransient<SettingsWindow>();
                services.AddTransient<AboutWindow>();
                services.AddTransient<FirstRunWizardViewModel>();
                services.AddTransient<FirstRunWizardWindow>();

                services.AddSingleton(pluginReport);
                foreach (var module in pluginReport.AllModules)
                {
                    module.RegisterServices(services);
                    services.AddTransient(module.ViewModelType);
                }
            })
            .Build();

        foreach (var module in pluginReport.AllModules)
        {
            try
            {
                var template = new System.Windows.DataTemplate
                {
                    DataType = module.ViewModelType,
                    VisualTree = new System.Windows.FrameworkElementFactory(module.ViewType)
                };
                Current.Resources.Add(new System.Windows.DataTemplateKey(module.ViewModelType), template);
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "No se pudo registrar template del módulo {Title}", module.Title);
            }
        }

        var monitor = Services.GetRequiredService<IConnectionStateMonitor>();
        monitor.Start();

        // Wire system theme detection: provider for "Auto" + live re-apply on OS theme change.
        ViewModels.SettingsViewModel.SystemThemeProvider = Services.GetRequiredService<ISystemThemeProvider>();
        Microsoft.Win32.SystemEvents.UserPreferenceChanged += OnSystemUserPreferenceChanged;

        TryApplySavedTheme();

        var main = Services.GetRequiredService<MainWindow>();
        main.Show();

        _ = ShowFirstRunWizardIfNeededAsync().ContinueWith(_ => TryAutoConnectAsync(), TaskScheduler.FromCurrentSynchronizationContext());

        base.OnStartup(e);
    }

    private async Task ShowFirstRunWizardIfNeededAsync()
    {
        try
        {
            var prefsStore = Services.GetRequiredService<IPreferencesStore>();
            var prefs = await prefsStore.LoadAsync().ConfigureAwait(true);
            if (prefs.FirstRunCompleted)
            {
                return;
            }

            var window = Services.GetRequiredService<FirstRunWizardWindow>();
            window.Owner = Current?.MainWindow;
            window.ShowDialog();

            // After wizard closes, re-apply theme in case user changed it.
            try
            {
                var updated = await prefsStore.LoadAsync().ConfigureAwait(true);
                ViewModels.SettingsViewModel.ApplyThemeFromPreferences(updated.Theme);
            }
            catch { /* non-critical */ }
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "First-run wizard failed (non-fatal)");
        }
    }

    private async Task TryAutoConnectAsync()
    {
        try
        {
            var certStore = Services.GetRequiredService<ICertConfigStore>();
            var certValidator = Services.GetRequiredService<ICertValidator>();
            var graph = Services.GetRequiredService<IGraphConnection>();
            var exchange = Services.GetRequiredService<IExchangeConnection>();
            var tenantLock = Services.GetRequiredService<ITenantLock>();
            var log = Services.GetRequiredService<IUiLogSink>();

            var config = await certStore.LoadAsync().ConfigureAwait(false);
            if (config is null)
            {
                Log.Information("Auto-connect skip: no cert config.");
                return;
            }
            var validation = certValidator.Validate(config);
            if (!validation.IsValid)
            {
                log.Progress.Report(LogEntry.Warn("AutoConnect", "Cert config presente pero inválido: " + validation.Message));
                return;
            }
            log.Progress.Report(LogEntry.Info("AutoConnect", "Cert válido detectado, conectando automáticamente..."));

            await graph.ConnectByCertificateAsync(config, log.Progress, CancellationToken.None).ConfigureAwait(false);

            try
            {
                await tenantLock.EnforceAsync(graph.TenantId ?? config.TenantId, CancellationToken.None).ConfigureAwait(false);
            }
            catch (TenantLockViolationException violation)
            {
                log.Progress.Report(LogEntry.Error("AutoConnect", "Tenant lock: " + violation.Message, violation));
                await graph.DisconnectAsync(CancellationToken.None).ConfigureAwait(false);
                return;
            }

            try
            {
                await exchange.ConnectByCertificateAsync(config, log.Progress, CancellationToken.None).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                log.Progress.Report(LogEntry.Warn("AutoConnect", "Exchange Online no se conectó: " + ex.Message));
            }
            log.Progress.Report(LogEntry.Ok("AutoConnect", "Conectado automáticamente."));
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Auto-connect falló");
            try
            {
                var log = Services.GetRequiredService<IUiLogSink>();
                log.Progress.Report(LogEntry.Warn("AutoConnect", "Auto-conexión falló: " + ex.Message));
            }
            catch { }
        }
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

    internal static LogEventLevel ParseLogLevel(string? value) => value?.Trim().ToLowerInvariant() switch
    {
        "debug" => LogEventLevel.Debug,
        "warning" or "warn" => LogEventLevel.Warning,
        "error" => LogEventLevel.Error,
        "fatal" => LogEventLevel.Fatal,
        "verbose" or "trace" => LogEventLevel.Verbose,
        _ => LogEventLevel.Information,
    };

    private static Grex365.Core.Models.UserPreferences TryLoadBootPreferences(string configDir)
    {
        try
        {
            var store = new Grex365.Core.Preferences.JsonPreferencesStore(configDir);
            return store.LoadAsync().GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "No se pudieron leer preferencias para arranque de plugins; usando defaults");
            return new Grex365.Core.Models.UserPreferences();
        }
    }

    private static void OnSystemUserPreferenceChanged(object? sender, Microsoft.Win32.UserPreferenceChangedEventArgs e)
    {
        if (e.Category != Microsoft.Win32.UserPreferenceCategory.General)
        {
            return;
        }
        // Only re-apply if user picked "Auto" — otherwise leave their explicit choice alone.
        try
        {
            var store = Services.GetRequiredService<IPreferencesStore>();
            var prefs = store.LoadAsync().GetAwaiter().GetResult();
            if (!string.Equals(prefs.Theme, "Auto", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }
            var dispatcher = Current?.Dispatcher;
            if (dispatcher is null) return;
            dispatcher.Invoke(() => ViewModels.SettingsViewModel.ApplyThemeFromPreferences("Auto"));
        }
        catch
        {
            // Non-critical; ignore.
        }
    }

    private static void TryApplySavedTheme()
    {
        try
        {
            var store = Services.GetRequiredService<IPreferencesStore>();
            var prefs = store.LoadAsync().GetAwaiter().GetResult();
            ViewModels.SettingsViewModel.ApplyThemeFromPreferences(prefs.Theme);
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "No se pudo aplicar tema guardado");
        }
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
        Microsoft.Win32.SystemEvents.UserPreferenceChanged -= OnSystemUserPreferenceChanged;
        if (_host is not null)
        {
            var monitor = Services.GetService<IConnectionStateMonitor>();
            if (monitor is not null)
            {
                await monitor.DisposeAsync().ConfigureAwait(false);
            }

            var pool = Services.GetService<RunspacePoolHost>();
            pool?.Dispose();

            if (Services.GetService<ITelemetry>() is ApplicationInsightsTelemetry ai)
            {
                ai.Flush();
                await Task.Delay(500).ConfigureAwait(false);
                ai.Dispose();
            }

            _host.Dispose();
        }
        Log.CloseAndFlush();
        base.OnExit(e);
    }
}
