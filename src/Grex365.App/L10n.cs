namespace Grex365.App;

public static class L10n
{
    private static readonly Dictionary<string, string> EmptyDict = new();
    private static Dictionary<string, string> _primary = EmptyDict;
    private static Dictionary<string, string> _fallback = EmptyDict;
    private static string _activeLanguage = "es";

    public const string DefaultLanguage = "es";
    public static IReadOnlyList<string> SupportedLanguages { get; } = new[] { "es", "en" };

    public static string ActiveLanguage => _activeLanguage;

    // Configurar desde dicts custom (tests + future i18n loader externo).
    public static void Configure(
        IDictionary<string, string> primary,
        IDictionary<string, string>? fallback = null,
        string activeLanguage = DefaultLanguage)
    {
        ArgumentNullException.ThrowIfNull(primary);
        _primary = new Dictionary<string, string>(primary, StringComparer.OrdinalIgnoreCase);
        _fallback = fallback is null
            ? EmptyDict
            : new Dictionary<string, string>(fallback, StringComparer.OrdinalIgnoreCase);
        _activeLanguage = activeLanguage;
    }

    // Inicializa con strings built-in. ES = canonical, EN parcial con fallback a ES.
    // Unknown / null / whitespace language → DefaultLanguage.
    public static void Initialize(string? languageCode)
    {
        var code = (languageCode ?? DefaultLanguage).Trim().ToLowerInvariant();
        if (!SupportedLanguages.Contains(code))
        {
            code = DefaultLanguage;
        }
        Configure(
            primary: code == "en" ? EnStrings : EsStrings,
            fallback: code == "en" ? EsStrings : null,
            activeLanguage: code);
    }

    public static string Get(string key)
    {
        if (string.IsNullOrEmpty(key)) return string.Empty;
        if (_primary.TryGetValue(key, out var v)) return v;
        if (_fallback.TryGetValue(key, out var f)) return f;
        return key;
    }

    public static void Reset()
    {
        _primary = EmptyDict;
        _fallback = EmptyDict;
        _activeLanguage = DefaultLanguage;
    }

    private static readonly Dictionary<string, string> EsStrings = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Nav.Dashboard"] = "Dashboard",
        ["Nav.Connection"] = "Conexión",
        ["Nav.Licenses"] = "Licencias",
        ["Nav.Users"] = "Usuarios",
        ["Nav.Groups"] = "Grupos",
        ["Nav.Onboarding"] = "Onboarding",
        ["Nav.Offboarding"] = "Offboarding",
        ["Nav.SharedMailbox"] = "Buzones",
        ["Nav.MailboxRules"] = "Reglas de buzón",
        ["Nav.MailFlow"] = "Flujo de correo",
        ["Nav.Audit"] = "Auditoría",
        ["Nav.AuditLog"] = "Registro de auditoría",
        ["Nav.PsConsole"] = "Consola PS",
        ["Nav.CertWizard"] = "Asistente cert",
        ["Nav.DnsCheck"] = "Comprobación DNS",

        ["NavCategory.Tenant"] = "Tenant",
        ["NavCategory.Identity"] = "Identidad",
        ["NavCategory.Mail"] = "Mail",
        ["NavCategory.Security"] = "Seguridad",
        ["NavCategory.Tools"] = "Herramientas",
        ["NavCategory.Plugins"] = "Plugins",
        ["NavCategory.Others"] = "Otros",

        ["Settings.Title"] = "Ajustes",
        ["Settings.Language"] = "Idioma",
        ["Settings.Language.Spanish"] = "Español",
        ["Settings.Language.English"] = "Inglés",
        ["Settings.Theme"] = "Tema",
        ["Settings.RestartRequired"] = "Reinicia la aplicación para aplicar el nuevo idioma.",

        ["Dialog.Ok"] = "Aceptar",
        ["Dialog.Cancel"] = "Cancelar",
        ["Dialog.Yes"] = "Sí",
        ["Dialog.No"] = "No",

        ["Common.Connect"] = "Conectar",
        ["Common.Disconnect"] = "Desconectar",
        ["Common.Refresh"] = "Refrescar",
        ["Common.Export"] = "Exportar",
        ["Common.Cancel"] = "Cancelar",
        ["Common.Apply"] = "Aplicar",
        ["Common.Clear"] = "Limpiar",
        ["Common.Search"] = "Buscar",
    };

    private static readonly Dictionary<string, string> EnStrings = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Nav.Dashboard"] = "Dashboard",
        ["Nav.Connection"] = "Connection",
        ["Nav.Licenses"] = "Licenses",
        ["Nav.Users"] = "Users",
        ["Nav.Groups"] = "Groups",
        ["Nav.Onboarding"] = "Onboarding",
        ["Nav.Offboarding"] = "Offboarding",
        ["Nav.SharedMailbox"] = "Shared mailboxes",
        ["Nav.MailboxRules"] = "Mailbox rules",
        ["Nav.MailFlow"] = "Mail flow",
        ["Nav.Audit"] = "Audit",
        ["Nav.AuditLog"] = "Audit log",
        ["Nav.PsConsole"] = "PS Console",
        ["Nav.CertWizard"] = "Cert wizard",
        ["Nav.DnsCheck"] = "DNS check",

        ["NavCategory.Tenant"] = "Tenant",
        ["NavCategory.Identity"] = "Identity",
        ["NavCategory.Mail"] = "Mail",
        ["NavCategory.Security"] = "Security",
        ["NavCategory.Tools"] = "Tools",
        ["NavCategory.Plugins"] = "Plugins",
        ["NavCategory.Others"] = "Others",

        ["Settings.Title"] = "Settings",
        ["Settings.Language"] = "Language",
        ["Settings.Language.Spanish"] = "Spanish",
        ["Settings.Language.English"] = "English",
        ["Settings.Theme"] = "Theme",
        ["Settings.RestartRequired"] = "Restart the app to apply the new language.",

        ["Dialog.Ok"] = "OK",
        ["Dialog.Cancel"] = "Cancel",
        ["Dialog.Yes"] = "Yes",
        ["Dialog.No"] = "No",

        ["Common.Connect"] = "Connect",
        ["Common.Disconnect"] = "Disconnect",
        ["Common.Refresh"] = "Refresh",
        ["Common.Export"] = "Export",
        ["Common.Cancel"] = "Cancel",
        ["Common.Apply"] = "Apply",
        ["Common.Clear"] = "Clear",
        ["Common.Search"] = "Search",
    };
}
