namespace Grex365.Core.Models;

public sealed record MailboxInfo(
    string Identity,
    string DisplayName,
    string PrimarySmtpAddress,
    string RecipientTypeDetails,
    bool LitigationHoldEnabled = false,
    int InPlaceHoldCount = 0,
    bool ArchiveEnabled = false,
    long? TotalItemBytes = null)
{
    public bool IsSharedMailbox =>
        string.Equals(RecipientTypeDetails, "SharedMailbox", StringComparison.OrdinalIgnoreCase);

    public bool IsUserMailbox =>
        string.Equals(RecipientTypeDetails, "UserMailbox", StringComparison.OrdinalIgnoreCase);

    public double? TotalItemSizeGb =>
        TotalItemBytes is { } b ? Math.Round(b / 1024d / 1024d / 1024d, 2) : null;

    // Unlicensed shared mailboxes are capped at 50 GB; above that the mailbox keeps
    // requiring a license (Exchange Online Plan 2), so removing it is unsafe.
    public bool ExceedsUnlicensedSharedLimit =>
        TotalItemBytes is { } b && b > 50L * 1024 * 1024 * 1024;

    // A litigation/in-place hold satisfied by the current license blocks safe removal.
    public bool HasBlockingHold => LitigationHoldEnabled || InPlaceHoldCount > 0;
}

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
