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

        // Step order follows Microsoft's "remove a former employee" guidance:
        //   block sign-in (+ revoke sessions) → convert mailbox to shared → remove licenses.
        // The mailbox MUST be converted to shared *while still licensed*; removing the
        // license first starts a 30-day deletion clock and hides the convert option. We
        // therefore convert before removing licenses and skip license removal if the
        // conversion failed, so the mailbox is never stranded for deletion.

        if (options.DisableAccount)
        {
            try
            {
                await _users.SetAccountEnabledAsync(user.Id, false, progress, cancellationToken).ConfigureAwait(false);
                // Disabling alone doesn't invalidate already-issued tokens — revoke sessions too.
                try
                {
                    await _users.RevokeSignInSessionsAsync(user.Id, progress, cancellationToken).ConfigureAwait(false);
                    steps.Add(new OffboardingStep("Deshabilitar cuenta", "OK", "AccountEnabled=false; sesiones revocadas"));
                }
                catch (Exception revokeEx)
                {
                    steps.Add(new OffboardingStep("Deshabilitar cuenta", "OK",
                        $"AccountEnabled=false; revoke sesiones falló: {revokeEx.Message}"));
                }
            }
            catch (Exception ex)
            {
                steps.Add(new OffboardingStep("Deshabilitar cuenta", "ERROR", ex.Message));
                success = false;
            }
        }

        var convertRequested = options.ConvertMailboxToShared;
        var convertSucceeded = false;
        if (convertRequested)
        {
            try
            {
                var info = await _mailboxes.ConvertToSharedAsync(upn, progress, cancellationToken).ConfigureAwait(false);
                var detail = info is null ? "Aplicado" : $"Tipo final: {info.RecipientTypeDetails}";
                convertSucceeded = info is null || info.IsSharedMailbox;
                steps.Add(new OffboardingStep("Mailbox->Shared", "OK", detail));
            }
            catch (Exception ex)
            {
                steps.Add(new OffboardingStep("Mailbox->Shared", "ERROR", ex.Message));
                success = false;
            }
        }

        if (options.RemoveLicenses)
        {
            // Safety gate: never strip the license if a requested shared conversion failed —
            // doing so would schedule the mailbox (and its data) for permanent deletion.
            if (convertRequested && !convertSucceeded)
            {
                steps.Add(new OffboardingStep("Quitar licencias", "OMITIDO",
                    "Conversión a buzón compartido falló; no se quitan licencias para evitar el borrado del buzón (30 días)."));
                success = false;
            }
            else
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
        }

        return new OffboardingResult(upn, success, steps);
    }
}
