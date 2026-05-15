namespace Grex365.Core.Models;

public sealed record LogEntry(
    DateTimeOffset Timestamp,
    LogSeverity Severity,
    string Source,
    string Message,
    Exception? Exception = null)
{
    public static LogEntry Info(string source, string message) =>
        new(DateTimeOffset.Now, LogSeverity.Info, source, message);

    public static LogEntry Ok(string source, string message) =>
        new(DateTimeOffset.Now, LogSeverity.Ok, source, message);

    public static LogEntry Warn(string source, string message) =>
        new(DateTimeOffset.Now, LogSeverity.Warning, source, message);

    public static LogEntry Error(string source, string message, Exception? ex = null) =>
        new(DateTimeOffset.Now, LogSeverity.Error, source, message, ex);

    public static LogEntry Debug(string source, string message) =>
        new(DateTimeOffset.Now, LogSeverity.Debug, source, message);
}

public enum LogSeverity
{
    Debug,
    Info,
    Ok,
    Warning,
    Error
}
