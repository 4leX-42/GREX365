namespace Grex365.App.ViewModels;

public static class NavTitleMigrator
{
    private static readonly Dictionary<string, string> RenamedNavTitles = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Conexion"] = "Conexión",
        ["Auditoria"] = "Auditoría",
        ["Reglas buzon"] = "Reglas de buzón",
        ["Salud tenant"] = "Licencias",
        ["Mail flow"] = "Flujo de correo",
        ["Audit log"] = "Registro de auditoría",
        ["Cert Wizard"] = "Asistente cert",
        ["DNS check"] = "Comprobación DNS",
    };

    public static string? Resolve(string? saved)
    {
        if (string.IsNullOrWhiteSpace(saved))
        {
            return null;
        }
        return RenamedNavTitles.TryGetValue(saved, out var renamed) ? renamed : saved;
    }

    public static IReadOnlyDictionary<string, string> RenameMap => RenamedNavTitles;
}
