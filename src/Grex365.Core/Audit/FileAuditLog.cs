using System.Text;
using System.Text.Json;
using Grex365.Core.Abstractions;

namespace Grex365.Core.Audit;

public sealed class FileAuditLog : IAuditLog
{
    private readonly string _baseDirectory;
    private readonly SemaphoreSlim _writeLock = new(1, 1);

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = false,
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
    };

    public FileAuditLog(string baseDirectory)
    {
        _baseDirectory = baseDirectory ?? throw new ArgumentNullException(nameof(baseDirectory));
        Directory.CreateDirectory(_baseDirectory);
    }

    public string GetMonthFilePath(int year, int month)
        => Path.Combine(_baseDirectory, $"audit-{year:D4}-{month:D2}.jsonl");

    public async Task WriteAsync(AuditRecord record, CancellationToken cancellationToken = default)
    {
        var path = GetMonthFilePath(record.Timestamp.Year, record.Timestamp.Month);
        var line = JsonSerializer.Serialize(record, JsonOptions);
        await _writeLock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            await File.AppendAllTextAsync(path, line + Environment.NewLine, Encoding.UTF8, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            _writeLock.Release();
        }
    }

    public async Task<IReadOnlyList<AuditRecord>> ReadMonthAsync(int year, int month, CancellationToken cancellationToken = default)
    {
        var path = GetMonthFilePath(year, month);
        if (!File.Exists(path))
        {
            return Array.Empty<AuditRecord>();
        }
        var lines = await File.ReadAllLinesAsync(path, Encoding.UTF8, cancellationToken).ConfigureAwait(false);
        var list = new List<AuditRecord>(lines.Length);
        foreach (var line in lines)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }
            try
            {
                var record = JsonSerializer.Deserialize<AuditRecord>(line, JsonOptions);
                if (record is not null)
                {
                    list.Add(record);
                }
            }
            catch
            {
                // skip malformed line
            }
        }
        return list;
    }
}
