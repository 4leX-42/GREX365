using Grex365.Core.Models;

namespace Grex365.Core.Groups;

public sealed record BulkGroupPlan(
    IReadOnlyList<BulkGroupRow> M365Rows,
    IReadOnlyList<BulkGroupRow> DlRows,
    int DistinctM365GroupCount,
    int DistinctDlGroupCount,
    string Breakdown,
    string TypeHint);

/// Splits and summarizes normalized bulk-group rows for the confirmation step in the UI.
/// Honors the user "Auto/M365/DL" choice: when forced, overrides per-row GroupType.
public static class BulkGroupPlanner
{
    public const string ForcedAuto = "Auto";
    public const string ForcedM365 = "M365";
    public const string ForcedDl = "DL";

    public const string TypeHintAuto = "Tipo detectado desde columna `GroupType` del CSV (default M365).";
    public const string TypeHintForcedTemplate = "Forzado por usuario: TODOS los grupos como {0}.";

    public static BulkGroupPlan Plan(IReadOnlyList<BulkGroupRow> rows, string bulkTypeChoice)
    {
        ArgumentNullException.ThrowIfNull(rows);

        var raw = (bulkTypeChoice ?? ForcedAuto).Trim();
        var isForcedM365 = string.Equals(raw, ForcedM365, StringComparison.OrdinalIgnoreCase);
        var isForcedDl = string.Equals(raw, ForcedDl, StringComparison.OrdinalIgnoreCase);
        var isForced = isForcedM365 || isForcedDl;

        IReadOnlyList<BulkGroupRow> effective;
        if (isForced)
        {
            var upper = isForcedM365 ? ForcedM365 : ForcedDl;
            effective = rows.Select(r => new BulkGroupRow(r.GroupName, r.Email, upper)).ToList();
        }
        else
        {
            effective = rows;
        }

        var m365 = effective.Where(r => string.Equals(r.GroupType, ForcedM365, StringComparison.OrdinalIgnoreCase)).ToList();
        var dl = effective.Where(r => string.Equals(r.GroupType, ForcedDl, StringComparison.OrdinalIgnoreCase)).ToList();

        var distinctM365 = m365.Select(r => r.GroupName).Distinct(StringComparer.OrdinalIgnoreCase).Count();
        var distinctDl = dl.Select(r => r.GroupName).Distinct(StringComparer.OrdinalIgnoreCase).Count();

        var breakdown = string.Join(" + ", new[]
        {
            distinctM365 > 0 ? $"{distinctM365} M365" : null,
            distinctDl > 0 ? $"{distinctDl} DL" : null,
        }.Where(s => s is not null));

        var typeHint = isForced
            ? string.Format(TypeHintForcedTemplate, isForcedM365 ? ForcedM365 : ForcedDl)
            : TypeHintAuto;

        return new BulkGroupPlan(m365, dl, distinctM365, distinctDl, breakdown, typeHint);
    }

    public static string BuildConfirmMessage(BulkGroupPlan plan, int rowCount, string domain)
    {
        ArgumentNullException.ThrowIfNull(plan);
        var trimmedDomain = (domain ?? string.Empty).Trim().TrimStart('@');
        return $"Se crearán/actualizarán {plan.Breakdown} ({rowCount} miembros) sobre @{trimmedDomain}.\n\n{plan.TypeHint}\n\n¿Continuar?";
    }

    public static string Summarize(IEnumerable<BulkGroupResult> results)
    {
        ArgumentNullException.ThrowIfNull(results);
        var list = results.ToList();
        var created = list.Count(r => string.Equals(r.Action, "Created", StringComparison.Ordinal));
        var existed = list.Count(r => string.Equals(r.Action, "Skipped", StringComparison.Ordinal));
        var added = list.Count(r => string.Equals(r.Action, "MemberAdded", StringComparison.Ordinal));
        var errors = list.Count(r => string.Equals(r.Action, "Error", StringComparison.Ordinal));
        return $"Grupos: Nuevos={created}  YaEstaban={existed}  Miembros={added}  Err={errors}";
    }
}
