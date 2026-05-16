using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.Core.Offboarding;

public sealed class OffboardingService : IOffboardingService
{
    private readonly IUsersService _users;
    private readonly ISharedMailboxService _mailboxes;

    public OffboardingService(IUsersService users, ISharedMailboxService mailboxes)
    {
        _users = users;
        _mailboxes = mailboxes;
    }

    public async Task<OffboardingResult> RunAsync(
        string upn,
        OffboardingOptions options,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var steps = new List<OffboardingStep>();
        var success = true;

        if (string.IsNullOrWhiteSpace(upn))
        {
            return new OffboardingResult(upn, false, new[]
            {
                new OffboardingStep("Validar", "ERROR", "UPN vacío")
            });
        }

        progress?.Report(LogEntry.Info("Offboarding", $"Iniciando offboarding de {upn}"));

        var user = await _users.GetByIdAsync(upn, cancellationToken).ConfigureAwait(false);
        if (user is null)
        {
            return new OffboardingResult(upn, false, new[]
            {
                new OffboardingStep("Buscar usuario", "ERROR", "Usuario no encontrado en Graph")
            });
        }
        steps.Add(new OffboardingStep("Buscar usuario", "OK",
            $"{user.DisplayName} (enabled={user.AccountEnabled}, lic={user.AssignedLicenseCount})"));

        if (options.DisableAccount)
        {
            try
            {
                await _users.SetAccountEnabledAsync(user.Id, false, progress, cancellationToken).ConfigureAwait(false);
                steps.Add(new OffboardingStep("Deshabilitar cuenta", "OK", "AccountEnabled=false"));
            }
            catch (Exception ex)
            {
                steps.Add(new OffboardingStep("Deshabilitar cuenta", "ERROR", ex.Message));
                success = false;
            }
        }

        if (options.RemoveLicenses)
        {
            try
            {
                await _users.RemoveAllLicensesAsync(user.Id, progress, cancellationToken).ConfigureAwait(false);
                steps.Add(new OffboardingStep("Quitar licencias", "OK",
                    user.AssignedLicenseCount == 0 ? "Sin licencias asignadas" : $"{user.AssignedLicenseCount} licencias retiradas"));
            }
            catch (Exception ex)
            {
                steps.Add(new OffboardingStep("Quitar licencias", "ERROR", ex.Message));
                success = false;
            }
        }

        if (options.ConvertMailboxToShared)
        {
            try
            {
                var info = await _mailboxes.ConvertToSharedAsync(upn, progress, cancellationToken).ConfigureAwait(false);
                var detail = info is null ? "Aplicado" : $"Tipo final: {info.RecipientTypeDetails}";
                steps.Add(new OffboardingStep("Mailbox->Shared", "OK", detail));
            }
            catch (Exception ex)
            {
                steps.Add(new OffboardingStep("Mailbox->Shared", "ERROR", ex.Message));
                success = false;
            }
        }

        return new OffboardingResult(upn, success, steps);
    }
}
