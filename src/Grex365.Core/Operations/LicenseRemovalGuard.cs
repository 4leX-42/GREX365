using Grex365.Core.Models;

namespace Grex365.Core.Operations;

// Removing a license from a mailbox that still depends on it loses access or data, or breaks a
// retention obligation. This is the mandatory check before any license removal. It is pure so
// the rule can be tested without a tenant; the caller supplies the mailbox facts it read.
public static class LicenseRemovalGuard
{
    // mailbox == null means the user has no Exchange Online mailbox, so nothing depends on the license.
    public static IReadOnlyList<string> Evaluate(MailboxInfo? mailbox)
    {
        var reasons = new List<string>();
        if (mailbox is null)
        {
            return reasons;
        }

        if (mailbox.TotalItemBytes is null)
        {
            reasons.Add("No se pudo leer el tamaño del buzón, así que no se puede descartar que supere los 50 GB.");
        }
        else if (mailbox.ExceedsUnlicensedSharedLimit)
        {
            reasons.Add($"El buzón ocupa {mailbox.TotalItemSizeGb} GB: por encima de 50 GB sigue necesitando licencia.");
        }

        if (mailbox.LitigationHoldEnabled)
        {
            reasons.Add("El buzón tiene Litigation Hold activo.");
        }

        if (mailbox.InPlaceHoldCount > 0)
        {
            reasons.Add($"El buzón tiene {mailbox.InPlaceHoldCount} retención(es) aplicadas (In-Place Hold o directiva de retención).");
        }

        if (mailbox.ArchiveEnabled)
        {
            reasons.Add("El buzón tiene archivo online, que depende de la licencia.");
        }

        return reasons;
    }
}
