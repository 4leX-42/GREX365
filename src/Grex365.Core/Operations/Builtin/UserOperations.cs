using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.Core.Operations.Builtin;

public abstract class ReadOperation : IOperation
{
    public abstract OperationDescriptor Descriptor { get; }

    public Task<OperationPreflight> PreflightAsync(OperationArguments arguments, CancellationToken cancellationToken) =>
        Task.FromResult(OperationPreflight.Allow(string.Empty));

    public abstract Task<object?> ExecuteAsync(
        OperationArguments arguments,
        IProgress<LogEntry>? progress,
        CancellationToken cancellationToken);
}

public sealed class UsersSearchOperation : ReadOperation
{
    private readonly IUsersService _users;

    public UsersSearchOperation(IUsersService users)
    {
        _users = users;
    }

    public override OperationDescriptor Descriptor { get; } = new(
        Id: "users.search",
        Module: "users",
        Title: "Buscar usuarios",
        Description: "Busca usuarios de Entra ID cuyo nombre, UPN o correo empieza por el texto dado. Devuelve como máximo 50.",
        Impact: OperationImpact.Read,
        Parameters: [new("query", OperationParameterType.String, "Inicio del nombre, UPN o correo.")],
        RequiredPermissions: ["User.Read.All"]);

    public override async Task<object?> ExecuteAsync(
        OperationArguments arguments,
        IProgress<LogEntry>? progress,
        CancellationToken cancellationToken) =>
        await _users.SearchAsync(arguments.GetString("query"), cancellationToken).ConfigureAwait(false);
}

public sealed record UserLicenseView(Guid SkuId, string SkuPartNumber);

public sealed record UserDetailsView(
    UserSummary User,
    IReadOnlyList<UserLicenseView> Licenses,
    IReadOnlyList<GroupSummary> Groups);

public sealed class UserGetOperation : ReadOperation
{
    private readonly IUsersService _users;

    public UserGetOperation(IUsersService users)
    {
        _users = users;
    }

    public override OperationDescriptor Descriptor { get; } = new(
        Id: "users.get",
        Module: "users",
        Title: "Ficha de usuario",
        Description: "Devuelve la ficha de un usuario: identidad, estado de la cuenta, licencias asignadas con su nombre de SKU y grupos a los que pertenece.",
        Impact: OperationImpact.Read,
        Parameters: [new("user", OperationParameterType.String, "UPN o identificador de objeto del usuario.")],
        RequiredPermissions: ["User.Read.All", "GroupMember.Read.All", "Organization.Read.All"]);

    public override async Task<object?> ExecuteAsync(
        OperationArguments arguments,
        IProgress<LogEntry>? progress,
        CancellationToken cancellationToken)
    {
        var key = arguments.GetString("user");
        var user = await _users.GetByIdAsync(key, cancellationToken).ConfigureAwait(false)
            ?? throw new InvalidOperationException($"No existe el usuario '{key}'.");

        var assigned = await _users.GetAssignedLicensesAsync(user.Id, cancellationToken).ConfigureAwait(false);
        var skus = await _users.ListSkusAsync(cancellationToken).ConfigureAwait(false);
        var names = skus.ToDictionary(s => s.SkuId, s => s.SkuPartNumber);
        var licenses = assigned
            .Select(id => new UserLicenseView(id, names.GetValueOrDefault(id, string.Empty)))
            .ToList();
        var groups = await _users.GetGroupMembershipsAsync(user.Id, cancellationToken).ConfigureAwait(false);

        return new UserDetailsView(user, licenses, groups);
    }
}

public sealed class LicensesListOperation : ReadOperation
{
    private readonly IUsersService _users;

    public LicensesListOperation(IUsersService users)
    {
        _users = users;
    }

    public override OperationDescriptor Descriptor { get; } = new(
        Id: "licenses.list",
        Module: "licenses",
        Title: "Licencias del tenant",
        Description: "Lista los SKU suscritos con unidades compradas, consumidas y libres.",
        Impact: OperationImpact.Read,
        Parameters: [],
        RequiredPermissions: ["Organization.Read.All"]);

    public override async Task<object?> ExecuteAsync(
        OperationArguments arguments,
        IProgress<LogEntry>? progress,
        CancellationToken cancellationToken) =>
        await _users.ListSkusAsync(cancellationToken).ConfigureAwait(false);
}

public sealed class MailboxGetOperation : ReadOperation
{
    private readonly IExternalExoOps _exo;

    public MailboxGetOperation(IExternalExoOps exo)
    {
        _exo = exo;
    }

    public override OperationDescriptor Descriptor { get; } = new(
        Id: "mailboxes.get",
        Module: "exchange",
        Title: "Datos de un buzón",
        Description: "Devuelve tipo de buzón, tamaño, retenciones (Litigation Hold, In-Place Hold) y si tiene archivo online. Estos son los datos que deciden si se puede retirar una licencia.",
        Impact: OperationImpact.Read,
        Parameters: [new("mailbox", OperationParameterType.String, "Dirección o UPN del buzón.")],
        RequiredPermissions: ["Exchange.ManageAsApp"],
        RequiresGraph: false,
        RequiresExchange: true);

    public override async Task<object?> ExecuteAsync(
        OperationArguments arguments,
        IProgress<LogEntry>? progress,
        CancellationToken cancellationToken)
    {
        var key = arguments.GetString("mailbox");
        return await _exo.GetMailboxFactsAsync(key, progress, cancellationToken).ConfigureAwait(false)
            ?? throw new InvalidOperationException($"No existe el buzón '{key}'.");
    }
}

public sealed record LicenseRemovalOutcome(string UserPrincipalName, Guid SkuId, string SkuPartNumber, bool StillAssigned);

public sealed class UserLicenseRemoveOperation : IOperation
{
    private readonly IUsersService _users;
    private readonly IExternalExoOps _exo;

    public UserLicenseRemoveOperation(IUsersService users, IExternalExoOps exo)
    {
        _users = users;
        _exo = exo;
    }

    public OperationDescriptor Descriptor { get; } = new(
        Id: "users.license.remove",
        Module: "licenses",
        Title: "Retirar una licencia a un usuario",
        Description: "Retira una licencia asignada directamente a un usuario, después de comprobar el estado de su buzón.",
        Impact: OperationImpact.Critical,
        Parameters:
        [
            new("user", OperationParameterType.String, "UPN o identificador de objeto del usuario."),
            new("sku", OperationParameterType.String, "SkuPartNumber (por ejemplo SPE_E3) o identificador del SKU."),
        ],
        RequiredPermissions: ["User.ReadWrite.All", "Organization.Read.All", "Exchange.ManageAsApp"],
        RequiresExchange: true,
        Safeguards:
        [
            "Se detiene si el buzón supera 50 GB o no se puede leer su tamaño.",
            "Se detiene si el buzón tiene Litigation Hold u otra retención.",
            "Se detiene si el buzón tiene archivo online.",
            "Se detiene si no se puede leer el estado del buzón.",
            "Solo retira asignaciones directas; una licencia heredada de un grupo se gestiona en el grupo.",
        ]);

    public async Task<OperationPreflight> PreflightAsync(OperationArguments arguments, CancellationToken cancellationToken)
    {
        var (user, sku) = await ResolveAsync(arguments, cancellationToken).ConfigureAwait(false);
        var plan = $"Retirar la licencia {sku.SkuPartNumber} ({sku.SkuId}) a {user.UserPrincipalName}.";

        var assigned = await _users.GetAssignedLicensesAsync(user.Id, cancellationToken).ConfigureAwait(false);
        if (!assigned.Contains(sku.SkuId))
        {
            return OperationPreflight.Block(plan, $"{user.UserPrincipalName} no tiene asignada la licencia {sku.SkuPartNumber}.");
        }

        // A failure reading the mailbox propagates: without its state the removal must not run.
        var mailbox = await _exo.GetMailboxFactsAsync(user.UserPrincipalName ?? user.Id, null, cancellationToken)
            .ConfigureAwait(false);
        var reasons = LicenseRemovalGuard.Evaluate(mailbox);
        var state = mailbox is null
            ? "El usuario no tiene buzón en Exchange Online."
            : $"Buzón {mailbox.RecipientTypeDetails}, {mailbox.TotalItemSizeGb?.ToString() ?? "?"} GB, sin retenciones ni archivo.";

        return reasons.Count > 0
            ? new OperationPreflight(plan, reasons)
            : OperationPreflight.Allow($"{plan} {state}");
    }

    public async Task<object?> ExecuteAsync(
        OperationArguments arguments,
        IProgress<LogEntry>? progress,
        CancellationToken cancellationToken)
    {
        var (user, sku) = await ResolveAsync(arguments, cancellationToken).ConfigureAwait(false);
        await _users.RemoveLicenseAsync(user.Id, sku.SkuId, progress, cancellationToken).ConfigureAwait(false);

        // Success is what the tenant says afterwards, not the absence of an error.
        var after = await _users.GetAssignedLicensesAsync(user.Id, cancellationToken).ConfigureAwait(false);
        var still = after.Contains(sku.SkuId);
        if (still)
        {
            throw new InvalidOperationException(
                $"La licencia {sku.SkuPartNumber} sigue asignada a {user.UserPrincipalName} tras la operación. Si viene de un grupo, hay que retirarla en el grupo.");
        }

        return new LicenseRemovalOutcome(user.UserPrincipalName ?? user.Id, sku.SkuId, sku.SkuPartNumber, still);
    }

    private async Task<(UserSummary User, SkuInfo Sku)> ResolveAsync(OperationArguments arguments, CancellationToken cancellationToken)
    {
        var userKey = arguments.GetString("user");
        var user = await _users.GetByIdAsync(userKey, cancellationToken).ConfigureAwait(false)
            ?? throw new InvalidOperationException($"No existe el usuario '{userKey}'.");

        var skuKey = arguments.GetString("sku");
        var skus = await _users.ListSkusAsync(cancellationToken).ConfigureAwait(false);
        var sku = skus.FirstOrDefault(s =>
                string.Equals(s.SkuPartNumber, skuKey, StringComparison.OrdinalIgnoreCase)
                || string.Equals(s.SkuId.ToString(), skuKey, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"El tenant no tiene ningún SKU '{skuKey}'.");

        return (user, sku);
    }
}
