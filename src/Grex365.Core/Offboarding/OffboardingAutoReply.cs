namespace Grex365.Core.Offboarding;

/// <summary>
/// Pure auto-reply template rendering. Substitutes the supported placeholders so a template
/// like "{usuario} ya no trabaja aquí; contacte con {delegado}" becomes a per-user message.
/// Side-effect-free and fully unit-testable; the UI owns the template catalogue + wording.
/// </summary>
public static class OffboardingAutoReply
{
    public const string UserToken = "{usuario}";
    public const string DelegateToken = "{delegado}";

    /// <summary>
    /// Replaces {usuario} with the leaver's name and {delegado} with the delegate contact
    /// (case-insensitive). Empty/null substitutions collapse to "". Returns the template
    /// unchanged when it has no tokens (free-text message).
    /// </summary>
    public static string Render(string? template, string? user, string? delegateContact)
    {
        if (string.IsNullOrEmpty(template))
        {
            return string.Empty;
        }
        return template
            .Replace(UserToken, user ?? string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace(DelegateToken, delegateContact ?? string.Empty, StringComparison.OrdinalIgnoreCase);
    }
}
