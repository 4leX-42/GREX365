using System.Reflection;
using System.Text;
using System.Text.Json;
using Grex365.Core.Abstractions;
using Grex365.Core.Audit;
using Grex365.Core.Connections;
using Grex365.Core.Models;
using Grex365.Core.Operations;
using Grex365.Core.Operations.Builtin;
using Grex365.Core.Operations.Mcp;
using Grex365.Core.Preferences;
using Grex365.Core.Security;
using Grex365.Core.Users;
using Grex365.PowerShell;
using Microsoft.Extensions.DependencyInjection;

namespace Grex365.Cli;

// Headless host of the GREX365 operation catalogue: the same operations, gates and audit log
// as the desktop app, for scripts (ops run), scheduled tasks and AI agents (mcp).
public static class Program
{
    private const string Usage = """
        grex365 — operaciones de GREX365 sin interfaz gráfica

        Uso:
          grex365 ops list [--json]                 catálogo de operaciones
          grex365 ops describe <id>                 parámetros, permisos y comprobaciones de una operación
          grex365 ops run <id> [--arg nombre=valor]... [--dry-run] [--plan-id <id>]
          grex365 mcp [--allow-writes]              servidor MCP por stdio para agentes de IA

        Opciones globales:
          --data-dir <ruta>   carpeta de datos (por defecto %LOCALAPPDATA%\Grex365)

        Las operaciones que modifican el tenant van en dos pasos: --dry-run devuelve el plan y un
        plan_id; la ejecución exige ese --plan-id. Los agentes (mcp) solo leen salvo --allow-writes,
        y nunca ejecutan operaciones críticas.
        """;

    public static async Task<int> Main(string[] args)
    {
        Console.OutputEncoding = new UTF8Encoding(false);
        var rest = new List<string>(args);
        var dataDir = TakeOption(rest, "--data-dir")
            ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Grex365");

        if (rest.Count == 0 || rest[0] is "-h" or "--help" or "help")
        {
            Console.WriteLine(Usage);
            return 0;
        }

        try
        {
            switch (rest[0])
            {
                case "ops" when rest.Count >= 2:
                    return await RunOpsAsync(rest, dataDir).ConfigureAwait(false);
                case "mcp":
                    return await RunMcpAsync(rest, dataDir).ConfigureAwait(false);
                default:
                    Console.Error.WriteLine(Usage);
                    return 2;
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            return 1;
        }
    }

    private static async Task<int> RunOpsAsync(List<string> rest, string dataDir)
    {
        await using var services = BuildServices(dataDir, new OperationPolicy());
        var registry = services.GetRequiredService<OperationRegistry>();

        switch (rest[1])
        {
            case "list":
                if (rest.Contains("--json"))
                {
                    Console.WriteLine(JsonSerializer.Serialize(registry.Descriptors, McpServer.Json));
                }
                else
                {
                    foreach (var d in registry.Descriptors)
                    {
                        Console.WriteLine($"{d.Id,-24} {d.Impact,-9} {d.Title}");
                    }
                }

                return 0;

            case "describe" when rest.Count >= 3:
                var found = registry.Find(rest[2]);
                if (found is null)
                {
                    Console.Error.WriteLine($"No existe la operación '{rest[2]}'.");
                    return 1;
                }

                Console.WriteLine(JsonSerializer.Serialize(found.Descriptor, Indented));
                return 0;

            case "run" when rest.Count >= 3:
                var tail = rest.Skip(3).ToList();
                var planId = TakeOption(tail, "--plan-id");
                var dryRun = tail.Remove("--dry-run");
                var arguments = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
                while (TakeOption(tail, "--arg") is { } pair)
                {
                    var cut = pair.IndexOf('=');
                    if (cut <= 0)
                    {
                        Console.Error.WriteLine($"--arg espera nombre=valor, no '{pair}'.");
                        return 2;
                    }

                    arguments[pair[..cut]] = pair[(cut + 1)..];
                }

                if (tail.Count > 0)
                {
                    Console.Error.WriteLine($"Opción no reconocida: {tail[0]}");
                    return 2;
                }

                var progress = new Progress<LogEntry>(e => Console.Error.WriteLine($"[{e.Severity}] {e.Source}: {e.Message}"));
                var result = await services.GetRequiredService<OperationExecutor>().RunAsync(
                    new OperationRequest(rest[2], arguments, OperationChannel.Cli, Actor(), dryRun, planId),
                    progress).ConfigureAwait(false);
                Console.WriteLine(JsonSerializer.Serialize(result, Indented));
                return result.Succeeded ? 0 : 1;

            default:
                Console.Error.WriteLine(Usage);
                return 2;
        }
    }

    private static async Task<int> RunMcpAsync(List<string> rest, string dataDir)
    {
        var policy = new OperationPolicy(AllowAgentWrites: rest.Contains("--allow-writes"));
        await using var services = BuildServices(dataDir, policy);
        var server = new McpServer(
            services.GetRequiredService<OperationRegistry>(),
            services.GetRequiredService<OperationExecutor>(),
            Actor(),
            Version());

        // stdout carries the protocol and nothing else.
        using var input = new StreamReader(Console.OpenStandardInput(), new UTF8Encoding(false));
        await using var output = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false));
        await server.RunAsync(input, output).ConfigureAwait(false);
        return 0;
    }

    private static ServiceProvider BuildServices(string dataDir, OperationPolicy policy)
    {
        var configDir = Path.Combine(dataDir, "config");
        var services = new ServiceCollection();

        services.AddSingleton<IPreferencesStore>(_ => new JsonPreferencesStore(configDir));
        services.AddSingleton<ICertConfigStore>(_ => new JsonCertConfigStore(configDir));
        services.AddSingleton<IAuditLog>(_ => new FileAuditLog(Path.Combine(dataDir, "audit")));

        services.AddSingleton<IGraphConnection, GraphConnection>();
        services.AddSingleton<ITenantLock, TenantLock>();
        services.AddSingleton<IUsersService, GraphUsersService>();
        services.AddSingleton<IExternalExoRunner, ExternalExoRunner>();
        services.AddSingleton<IExternalExoOps, ExternalExoOps>();
        services.AddSingleton<IMembershipChecker, GraphMembershipChecker>();
        services.AddSingleton<IRbacGuard>(sp =>
        {
            var preferences = sp.GetRequiredService<IPreferencesStore>();
            return new RbacGuard(sp.GetRequiredService<IMembershipChecker>(), () =>
            {
                try
                {
                    return preferences.LoadAsync().GetAwaiter().GetResult().AuthorizationGroupId;
                }
                catch
                {
                    return null;
                }
            });
        });

        services.AddSingleton<IOperation, UsersSearchOperation>();
        services.AddSingleton<IOperation, UserGetOperation>();
        services.AddSingleton<IOperation, LicensesListOperation>();
        services.AddSingleton<IOperation, MailboxGetOperation>();
        services.AddSingleton<IOperation, UserLicenseRemoveOperation>();

        services.AddSingleton<OperationRegistry>();
        services.AddSingleton<IOperationSession, CertOperationSession>();
        services.AddSingleton(policy);
        services.AddSingleton(_ => PlanTokenService.FromKeyFile(Path.Combine(configDir, "plan.key")));
        services.AddSingleton(sp => new OperationExecutor(
            sp.GetRequiredService<OperationRegistry>(),
            sp.GetRequiredService<IOperationSession>(),
            sp.GetRequiredService<IRbacGuard>(),
            sp.GetRequiredService<IAuditLog>(),
            sp.GetRequiredService<PlanTokenService>(),
            sp.GetRequiredService<OperationPolicy>()));

        return services.BuildServiceProvider();
    }

    private static readonly JsonSerializerOptions Indented = new(McpServer.Json) { WriteIndented = true };

    private static string Actor() => $"{Environment.UserDomainName}\\{Environment.UserName}";

    private static string Version() =>
        typeof(Program).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? "0.0.0";

    private static string? TakeOption(List<string> args, string name)
    {
        var index = args.IndexOf(name);
        if (index < 0 || index + 1 >= args.Count)
        {
            return null;
        }

        var value = args[index + 1];
        args.RemoveRange(index, 2);
        return value;
    }
}
