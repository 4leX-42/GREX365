namespace Grex365.Core.Operations;

// The catalogue every channel reads. A module or plugin adds a capability by registering an
// IOperation in DI; nothing else in the core changes.
public sealed class OperationRegistry
{
    private readonly Dictionary<string, IOperation> _operations;

    public OperationRegistry(IEnumerable<IOperation> operations)
    {
        _operations = new Dictionary<string, IOperation>(StringComparer.OrdinalIgnoreCase);
        foreach (var operation in operations)
        {
            var id = operation.Descriptor.Id;
            if (!_operations.TryAdd(id, operation))
            {
                throw new InvalidOperationException($"Operación duplicada en el catálogo: '{id}'.");
            }
        }
    }

    public IReadOnlyList<OperationDescriptor> Descriptors =>
        _operations.Values
            .Select(o => o.Descriptor)
            .OrderBy(d => d.Id, StringComparer.Ordinal)
            .ToList();

    public IOperation? Find(string id) =>
        _operations.TryGetValue(id ?? string.Empty, out var operation) ? operation : null;
}
