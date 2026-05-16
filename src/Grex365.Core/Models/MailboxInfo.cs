namespace Grex365.Core.Models;

public sealed record MailboxInfo(
    string Identity,
    string DisplayName,
    string PrimarySmtpAddress,
    string RecipientTypeDetails);

public sealed record MailboxPermissionEntry(
    string Permission,
    string Principal,
    string Detail);

public sealed record MailboxPermissionResult(
    string Action,
    string Permission,
    string Mailbox,
    string Principal,
    string Status,
    string Detail);
