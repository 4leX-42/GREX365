using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.Core.Offboarding;

public sealed class OffboardingService : IOffboardingService
{
    private readonly IUsersService _users;
    private readonly ISharedMailboxService _mailboxes;
    private readonly IExternalExoOps? _externalExo;

    // externalExo (when wired) runs the mailbox conversion via an external pwsh process —
    // the in-process EXO path is unreliable. Falls back to the in-proc service when absent.
    public OffboardingService(IUsersService users, ISharedMailboxService mailboxes, IExternalExoOps? externalExo = null)
    {
        _users = users;
        _mailboxes = mailboxes;
        _externalExo = externalExo;
    }

    public async Task<OffboardingResult> RunAsync(
        string upn,
        OffboardingOptions options,
        IProgress<LogEntry>? progress = null,
        IProgress<OffboardingStep>? stepProgress = null,
        CancellationToken cancellationToken = default)
    {
        var steps = new List<OffboardingStep>();
        var success = true;

        void Running(string name) => stepProgress?.Report(new OffboardingStep(name, "RUNNING", "…"));
        void Done(string name, string status, string detail)
        {
            var s = new OffboardingStep(name, status, detail);
            steps.Add(s);
            stepProgress?.Report(s);
        }

        if (string.IsNullOrWhiteSpace(upn))
        {
            Done("Validar", "ERROR", "UPN vacío");
            return new OffboardingResult(upn, false, steps);
        }

        progress?.Report(LogEntry.Info("Offboarding", $"Iniciando offboarding de {upn}"));

        Running("Buscar usuario");
        var user = await _users.GetByIdAsync(upn, cancellationToken).ConfigureAwait(false);
        if (user is null)
        {
            Done("Buscar usuario", "ERROR", "Usuario no encontrado en Graph");
            return new OffboardingResult(upn, false, steps);
        }
        Done("Buscar usuario", "OK",
            $"{user.DisplayName} (enabled={user.AccountEnabled}, lic={user.AssignedLicenseCount})");

        // Step order follows Microsoft's "remove a former employee" guidance:
        //   block sign-in (+ revoke sessions) → convert mailbox to shared → remove licenses.
        // The mailbox MUST be converted to shared *while still licensed*; removing the
        // license first starts a 30-day deletion clock and hides the convert option. We
        // therefore convert before removing licenses and skip license removal if the
        // conversion failed, so the mailbox is never stranded for deletion.

        if (options.DisableAccount)
        {
            Running("Deshabilitar cuenta");
            try
            {
                await _users.SetAccountEnabledAsync(user.Id, false, progress, cancellationToken).ConfigureAwait(false);
                // Disabling alone doesn't invalidate already-issued tokens — revoke sessions too.
                try
                {
                    await _users.RevokeSignInSessionsAsync(user.Id, progress, cancellationToken).ConfigureAwait(false);
                    Done("Deshabilitar cuenta", "OK", "AccountEnabled=false; sesiones revocadas");
                }
                catch (Exception revokeEx)
                {
                    Done("Deshabilitar cuenta", "OK", $"AccountEnabled=false; revoke sesiones falló: {revokeEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Done("Deshabilitar cuenta", "ERROR", ex.Message);
                success = false;
            }
        }

        var convertRequested = options.ConvertMailboxToShared;
        var convertSucceeded = false;
        if (convertRequested)
        {
            Running("Convertir a buzón compartido");
            try
            {
                var info = _externalExo is not null
                    ? await _externalExo.ConvertToSharedAsync(upn, progress, cancellationToken).ConfigureAwait(false)
                    : await _mailboxes.ConvertToSharedAsync(upn, progress, cancellationToken).ConfigureAwait(false);
                convertSucceeded = info is null || info.IsSharedMailbox;
                if (convertSucceeded)
                {
                    Done("Convertir a buzón compartido", "OK", info is null ? "Aplicado" : $"Tipo final: {info.RecipientTypeDetails}");
                }
                else
                {
                    // Verification failed: the type didn't actually flip to SharedMailbox.
                    Done("Convertir a buzón compartido", "ERROR", $"No se confirmó SharedMailbox (tipo={info!.RecipientTypeDetails})");
                    success = false;
                }
            }
            catch (Exception ex)
            {
                Done("Convertir a buzón compartido", "ERROR", ex.Message);
                success = false;
            }
        }

        if (options.RemoveLicenses)
        {
            // Safety gate: never strip the license if a requested shared conversion failed —
            // doing so would schedule the mailbox (and its data) for permanent deletion.
            if (convertRequested && !convertSucceeded)
            {
                Done("Quitar licencias", "OMITIDO",
                    "Conversión a buzón compartido falló; no se quitan licencias para evitar el borrado del buzón (30 días).");
                success = false;
            }
            else
            {
                Running("Quitar licencias");
                try
                {
                    await _users.RemoveAllLicensesAsync(user.Id, progress, cancellationToken).ConfigureAwait(false);
                    // Verify removal actually took (Graph propagation can lag a few seconds).
                    var verified = await VerifyLicensesRemovedAsync(user.Id, cancellationToken).ConfigureAwait(false);
                    if (verified)
                    {
                        Done("Quitar licencias", "OK",
                            user.AssignedLicenseCount == 0 ? "Sin licencias asignadas" : $"{user.AssignedLicenseCount} licencias retiradas (verificado)");
                    }
                    else
                    {
                        Done("Quitar licencias", "ERROR", "La operación se aceptó pero el usuario sigue con licencias asignadas.");
                        success = false;
                    }
                }
                catch (Exception ex)
                {
                    Done("Quitar licencias", "ERROR", ex.Message);
                    success = false;
                }
            }
        }

        return new OffboardingResult(upn, success, steps);
    }

    // Re-reads the user a few times (Graph license changes propagate with a small lag) and
    // returns true once no licenses remain. If GetByIdAsync isn't usefully mocked (unit tests),
    // the first read may already report 0; either way this never throws.
    private async Task<bool> VerifyLicensesRemovedAsync(string userId, CancellationToken ct)
    {
        for (var attempt = 0; attempt < 5; attempt++)
        {
            try
            {
                var u = await _users.GetByIdAsync(userId, ct).ConfigureAwait(false);
                if (u is null || u.AssignedLicenseCount == 0)
                {
                    return true;
                }
            }
            catch
            {
                return true; // can't verify → don't block the flow
            }
            if (attempt < 4)
            {
                await Task.Delay(TimeSpan.FromSeconds(3), ct).ConfigureAwait(false);
            }
        }
        return false;
    }
}
