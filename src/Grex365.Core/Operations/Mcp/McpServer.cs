using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;

namespace Grex365.Core.Operations.Mcp;

// Model Context Protocol server over a pair of text streams (stdio in practice). It exposes the
// operation catalogue as tools and runs every call through OperationExecutor, so an agent gets
// exactly the capabilities, gates and audit trail the operator gets — nothing of its own.
//
// Transport: one JSON-RPC 2.0 message per line. Implements initialize, ping, tools/list and
// tools/call, which is all a tools-only server needs.
public sealed class McpServer
{
    private static readonly string[] SupportedProtocols = ["2025-06-18", "2025-03-26", "2024-11-05"];

    public static readonly JsonSerializerOptions Json = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        Converters = { new JsonStringEnumConverter() },
    };

    private readonly OperationRegistry _registry;
    private readonly OperationExecutor _executor;
    private readonly string _actor;
    private readonly string _version;

    public McpServer(OperationRegistry registry, OperationExecutor executor, string actor, string version)
    {
        _registry = registry;
        _executor = executor;
        _actor = actor;
        _version = version;
    }

    public static string ToolName(string operationId) => operationId.Replace('.', '_');

    public async Task RunAsync(TextReader input, TextWriter output, CancellationToken cancellationToken = default)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            var line = await input.ReadLineAsync(cancellationToken).ConfigureAwait(false);
            if (line is null)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            var response = await HandleAsync(line, cancellationToken).ConfigureAwait(false);
            if (response is not null)
            {
                await output.WriteLineAsync(response.ToJsonString()).ConfigureAwait(false);
                await output.FlushAsync(cancellationToken).ConfigureAwait(false);
            }
        }
    }

    // Returns the response message, or null for notifications.
    public async Task<JsonObject?> HandleAsync(string message, CancellationToken cancellationToken = default)
    {
        JsonNode? parsed;
        try
        {
            parsed = JsonNode.Parse(message);
        }
        catch (JsonException)
        {
            return Error(null, -32700, "Parse error");
        }

        if (parsed is not JsonObject request || request["method"] is null)
        {
            return Error(null, -32600, "Invalid Request");
        }

        var id = request["id"]?.DeepClone();
        var method = request["method"]!.GetValue<string>();
        if (id is null)
        {
            return null;
        }

        switch (method)
        {
            case "initialize":
                var asked = request["params"]?["protocolVersion"]?.GetValue<string>();
                return Result(id, new JsonObject
                {
                    ["protocolVersion"] = SupportedProtocols.Contains(asked) ? asked : SupportedProtocols[0],
                    ["capabilities"] = new JsonObject { ["tools"] = new JsonObject() },
                    ["serverInfo"] = new JsonObject { ["name"] = "grex365", ["version"] = _version },
                    ["instructions"] =
                        "Operaciones de administración de Microsoft 365 del tenant configurado en GREX365. "
                        + "Las herramientas de lectura se ejecutan directamente. Las de escritura van en dos pasos: "
                        + "llama con dry_run=true, enseña el plan al usuario y repite con el plan_id devuelto. "
                        + "Si una herramienta responde Blocked o Denied, no busques otra vía: informa del motivo.",
                });

            case "ping":
                return Result(id, new JsonObject());

            case "tools/list":
                return Result(id, new JsonObject
                {
                    ["tools"] = new JsonArray(_registry.Descriptors.Select(d => (JsonNode)DescribeTool(d)).ToArray()),
                });

            case "tools/call":
                return Result(id, await CallAsync(request["params"] as JsonObject, cancellationToken).ConfigureAwait(false));

            default:
                return Error(id, -32601, $"Method not found: {method}");
        }
    }

    private static JsonObject DescribeTool(OperationDescriptor descriptor)
    {
        var properties = new JsonObject();
        var required = new JsonArray();
        foreach (var parameter in descriptor.Parameters)
        {
            properties[parameter.Name] = new JsonObject
            {
                ["type"] = parameter.Type switch
                {
                    OperationParameterType.Integer => "integer",
                    OperationParameterType.Boolean => "boolean",
                    _ => "string",
                },
                ["description"] = parameter.Description,
            };
            if (parameter.Required)
            {
                required.Add(parameter.Name);
            }
        }

        var description = descriptor.Description;
        var isWrite = descriptor.Impact != OperationImpact.Read;
        if (isWrite)
        {
            properties["dry_run"] = new JsonObject
            {
                ["type"] = "boolean",
                ["description"] = "true = solo comprobar y devolver el plan y su plan_id, sin cambiar nada.",
            };
            properties["plan_id"] = new JsonObject
            {
                ["type"] = "string",
                ["description"] = "plan_id devuelto por la simulación de esta misma llamada. Obligatorio para ejecutar.",
            };
            description += $" Impacto: {descriptor.Impact}.";
            if (descriptor.Safeguards is { Count: > 0 })
            {
                description += " Comprobaciones previas: " + string.Join(" ", descriptor.Safeguards);
            }
        }

        description += " Permisos: " + string.Join(", ", descriptor.RequiredPermissions) + ".";

        return new JsonObject
        {
            ["name"] = ToolName(descriptor.Id),
            ["title"] = descriptor.Title,
            ["description"] = description,
            ["inputSchema"] = new JsonObject
            {
                ["type"] = "object",
                ["properties"] = properties,
                ["required"] = required,
            },
            ["annotations"] = new JsonObject
            {
                ["readOnlyHint"] = !isWrite,
                ["destructiveHint"] = descriptor.Impact == OperationImpact.Critical,
            },
        };
    }

    private async Task<JsonObject> CallAsync(JsonObject? parameters, CancellationToken cancellationToken)
    {
        var toolName = parameters?["name"]?.GetValue<string>() ?? string.Empty;
        var descriptor = _registry.Descriptors.FirstOrDefault(d =>
            string.Equals(ToolName(d.Id), toolName, StringComparison.OrdinalIgnoreCase));
        if (descriptor is null)
        {
            return ToolResult(new OperationResult(OperationStatus.NotFound, $"No existe la herramienta '{toolName}'."));
        }

        var arguments = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        var dryRun = false;
        string? planId = null;
        if (parameters?["arguments"] is JsonObject supplied)
        {
            foreach (var (name, value) in supplied)
            {
                var text = value switch
                {
                    null => null,
                    JsonValue scalar when scalar.TryGetValue<string>(out var s) => s,
                    _ => value.ToJsonString(),
                };
                if (name == "dry_run")
                {
                    dryRun = string.Equals(text, "true", StringComparison.OrdinalIgnoreCase);
                }
                else if (name == "plan_id")
                {
                    planId = text;
                }
                else
                {
                    arguments[name] = text;
                }
            }
        }

        var result = await _executor.RunAsync(
            new OperationRequest(descriptor.Id, arguments, OperationChannel.Mcp, _actor, dryRun, planId),
            progress: null,
            cancellationToken).ConfigureAwait(false);
        return ToolResult(result);
    }

    private static JsonObject ToolResult(OperationResult result) => new()
    {
        ["content"] = new JsonArray(new JsonObject
        {
            ["type"] = "text",
            ["text"] = JsonSerializer.Serialize(result, Json),
        }),
        ["isError"] = !result.Succeeded,
    };

    private static JsonObject Result(JsonNode id, JsonObject result) => new()
    {
        ["jsonrpc"] = "2.0",
        ["id"] = id,
        ["result"] = result,
    };

    private static JsonObject Error(JsonNode? id, int code, string message) => new()
    {
        ["jsonrpc"] = "2.0",
        ["id"] = id,
        ["error"] = new JsonObject { ["code"] = code, ["message"] = message },
    };
}
