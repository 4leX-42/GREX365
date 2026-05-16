namespace Grex365.Core.Models;

public enum AutoReplyState
{
    Disabled,
    Enabled,
    Scheduled
}

public sealed record AutoReplyConfig(
    AutoReplyState State,
    string? InternalMessage,
    string? ExternalMessage,
    DateTime? StartTime,
    DateTime? EndTime);

public sealed record ForwardingConfig(
    string? ForwardingAddress,
    string? ForwardingSmtpAddress,
    bool DeliverToMailboxAndForward);

public sealed record CalendarPermissionEntry(string Principal, string AccessRights);

public static class CalendarAccessRights
{
    public const string Owner = "Owner";
    public const string PublishingEditor = "PublishingEditor";
    public const string Editor = "Editor";
    public const string PublishingAuthor = "PublishingAuthor";
    public const string Author = "Author";
    public const string NonEditingAuthor = "NonEditingAuthor";
    public const string Reviewer = "Reviewer";
    public const string Contributor = "Contributor";
    public const string LimitedDetails = "LimitedDetails";
    public const string AvailabilityOnly = "AvailabilityOnly";
    public const string None = "None";

    public static readonly IReadOnlyList<string> All = new[]
    {
        Owner, PublishingEditor, Editor, PublishingAuthor, Author,
        NonEditingAuthor, Reviewer, Contributor, LimitedDetails, AvailabilityOnly, None
    };
}
