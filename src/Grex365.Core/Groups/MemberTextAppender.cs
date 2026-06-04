namespace Grex365.Core.Groups;

// Pure helper for the "pick & append" UX on the Groups add-members box: append a chosen entry
// (UPN/email/id) to the newline-separated text, de-duplicated case-insensitively against the
// existing entries. Splits on the same delimiters GroupsViewModel uses to parse the box
// (newline / comma / semicolon) so a dupe is caught regardless of how the user typed the list.
// Existing formatting is preserved — the entry is only appended on a fresh line.
public static class MemberTextAppender
{
    private static readonly char[] Delimiters = { '\n', '\r', ',', ';' };

    public static string Append(string? existingText, string? entry)
    {
        var clean = (entry ?? string.Empty).Trim();
        var text = existingText ?? string.Empty;
        if (clean.Length == 0)
        {
            return text;
        }

        var existing = text.Split(Delimiters, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (existing.Any(e => string.Equals(e, clean, StringComparison.OrdinalIgnoreCase)))
        {
            return text; // already present — unchanged
        }

        if (text.Length == 0)
        {
            return clean;
        }
        // Append on a new line without doubling an existing trailing newline.
        var needsNewline = !(text.EndsWith('\n') || text.EndsWith('\r'));
        return needsNewline ? text + Environment.NewLine + clean : text + clean;
    }
}
