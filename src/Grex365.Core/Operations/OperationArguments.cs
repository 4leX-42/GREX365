using System.Globalization;

namespace Grex365.Core.Operations;

// Arguments after validation against the descriptor: every required one is present and every
// value parses as its declared type, so operations read them without re-checking.
public sealed class OperationArguments
{
    private readonly IReadOnlyDictionary<string, string?> _values;

    private OperationArguments(IReadOnlyDictionary<string, string?> values)
    {
        _values = values;
    }

    public string GetString(string name) =>
        _values.TryGetValue(name, out var value) && !string.IsNullOrWhiteSpace(value)
            ? value.Trim()
            : throw new InvalidOperationException($"Falta el argumento '{name}'.");

    public string? GetOptionalString(string name) =>
        _values.TryGetValue(name, out var value) && !string.IsNullOrWhiteSpace(value) ? value.Trim() : null;

    public int? GetOptionalInt(string name) =>
        GetOptionalString(name) is { } raw ? int.Parse(raw, CultureInfo.InvariantCulture) : null;

    public bool GetBool(string name, bool fallback = false) =>
        GetOptionalString(name) is { } raw ? bool.Parse(raw) : fallback;

    // Stable text form of the arguments: a plan confirms exactly these values and no others.
    public string Canonical() =>
        string.Join("\n", _values
            .Where(pair => !string.IsNullOrWhiteSpace(pair.Value))
            .OrderBy(pair => pair.Key, StringComparer.Ordinal)
            .Select(pair => $"{pair.Key}={pair.Value!.Trim()}"));

    public static OperationArguments Validate(
        OperationDescriptor descriptor,
        IReadOnlyDictionary<string, string?> raw,
        out IReadOnlyList<string> errors)
    {
        var problems = new List<string>();
        var known = descriptor.Parameters.ToDictionary(p => p.Name, StringComparer.OrdinalIgnoreCase);
        var values = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);

        foreach (var (name, value) in raw)
        {
            if (!known.TryGetValue(name, out var parameter))
            {
                problems.Add($"Argumento desconocido: '{name}'.");
                continue;
            }

            values[parameter.Name] = value;
        }

        foreach (var parameter in descriptor.Parameters)
        {
            values.TryGetValue(parameter.Name, out var value);
            if (string.IsNullOrWhiteSpace(value))
            {
                if (parameter.Required)
                {
                    problems.Add($"Falta el argumento obligatorio '{parameter.Name}'.");
                }

                continue;
            }

            var parses = parameter.Type switch
            {
                OperationParameterType.Integer => int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out _),
                OperationParameterType.Boolean => bool.TryParse(value, out _),
                _ => true,
            };
            if (!parses)
            {
                problems.Add($"El argumento '{parameter.Name}' debe ser de tipo {parameter.Type}.");
            }
        }

        errors = problems;
        return new OperationArguments(values);
    }
}
