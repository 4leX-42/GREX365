namespace Grex365.Core.Audit;

public static class GraphPermissionErrorDetector
{
    // True iff the exception message indicates a missing Microsoft Graph permission.
    // Detection is heuristic: Graph SDK does not surface ODataError consistently,
    // so we substring-match on the message text. Used by GraphAuditService.RunIdentityAuditAsync
    // to retry without `signInActivity` in $select when `AuditLog.Read.All` is missing —
    // degrades gracefully instead of aborting the whole audit.
    public static bool IsAuditLogPermissionError(Exception? ex)
    {
        if (ex is null) return false;
        var msg = ex.Message ?? string.Empty;
        return msg.Contains("AuditLog.Read.All", StringComparison.OrdinalIgnoreCase)
            || msg.Contains("required Microsoft Graph permission", StringComparison.OrdinalIgnoreCase);
    }

    // Generic permission error detector — matches common Graph SDK error phrases.
    public static bool IsAnyPermissionError(Exception? ex)
    {
        if (ex is null) return false;
        var msg = ex.Message ?? string.Empty;
        return msg.Contains("required Microsoft Graph permission", StringComparison.OrdinalIgnoreCase)
            || msg.Contains("Insufficient privileges", StringComparison.OrdinalIgnoreCase)
            || msg.Contains("Authorization_RequestDenied", StringComparison.OrdinalIgnoreCase)
            || msg.Contains("Forbidden", StringComparison.OrdinalIgnoreCase);
    }
}
