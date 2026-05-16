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
