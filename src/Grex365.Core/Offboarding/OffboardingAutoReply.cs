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

    /// <summary>
    /// Resolves the final auto-reply message for one leaver, applying the grouping/exception model:
    ///   1. a per-user override always wins (individual customisation);
    ///   2. otherwise, if auto-reply is enabled globally (globalTemplate non-empty), pick the
    ///      no-delegate template when the user has no delegate, else the global template;
    ///   3. if auto-reply is off and there is no override → null (skip the step).
    /// The chosen body is then rendered with {usuario}/{delegado}. This lets a batch share one
    /// config (set the delegate per group) while still honouring per-user exceptions.
    /// </summary>
    public static string? Resolve(
        string? globalTemplate,
        string? perUserOverride,
        string? noDelegateTemplate,
        string? user,
        string? delegateContact)
    {
        if (!string.IsNullOrWhiteSpace(perUserOverride))
        {
            return Render(perUserOverride, user, delegateContact);
        }
        if (string.IsNullOrWhiteSpace(globalTemplate))
        {
            return null; // auto-reply disabled and no per-user override
        }
        var hasDelegate = !string.IsNullOrWhiteSpace(delegateContact);
        var body = !hasDelegate && !string.IsNullOrWhiteSpace(noDelegateTemplate)
            ? noDelegateTemplate
            : globalTemplate;
        return Render(body, user, delegateContact);
    }
}
