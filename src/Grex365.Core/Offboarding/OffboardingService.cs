using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.Core.Offboarding;

public sealed class OffboardingService : IOffboardingService
{
    private readonly IUsersService _users;
    private readonly ISharedMailboxService _mailboxes;
    private readonly IExternalExoOps? _externalExo;

    // externalExo (when wired) runs the mailbox conversion + fact-gathering via an external
    // pwsh process — the in-process EXO path is unreliable. Falls back to the in-proc service
    // when absent.
    public OffboardingService(IUsersService users, ISharedMailboxService mailboxes, IExternalExoOps? externalExo = null)
    {
        _users = users;
        _mailboxes = mailboxes;
        _externalExo = externalExo;
    }

    // Delay between license-removal verification re-reads (Graph license changes propagate
    // with a small lag). Exposed as init-only tuning knobs so unit tests can run instantly.
    public TimeSpan VerifyPollDelay { get; init; } = TimeSpan.FromSeconds(3);
    public int VerifyAttempts { get; init; } = 5;

    public async Task<OffboardingResult> RunAsync(
        string upn,
        OffboardingOptions options,
        IProgress<LogEntry>? progress = null,
        IProgress<OffboardingStep>? stepProgress = null,
        CancellationToken cancellationToken = default)
    {
        var startedAt = DateTimeOffset.Now;
        var steps = new List<OffboardingStep>();
        var success = true;
        var dry = options.DryRun;

        void Running(string name) => stepProgress?.Report(new OffboardingStep(name, "RUNNING", "…"));
        void Done(string name, string status, string detail)
        {
            var s = new OffboardingStep(name, status, detail, DateTimeOffset.Now);
            steps.Add(s);
            stepProgress?.Report(s);
        }
        OffboardingResult Result(bool ok) => new(upn, ok, steps, dry, startedAt, DateTimeOffset.Now);

        if (string.IsNullOrWhiteSpace(upn))
        {
            Done("Validar", "ERROR", "UPN vacío");
            return Result(false);
        }

        progress?.Report(LogEntry.Info("Offboarding", $"Iniciando offboarding{(dry ? " (DRY-RUN)" : "")} de {upn}"));

        Running("Buscar usuario");
        var user = await _users.GetByIdAsync(upn, cancellationToken).ConfigureAwait(false);
        if (user is null)
        {
            Done("Buscar usuario", "ERROR", "Usuario no encontrado en Graph");
            return Result(false);
        }
        Done("Buscar usuario", "OK",
            $"{user.DisplayName} (enabled={user.AccountEnabled}, lic={user.AssignedLicenseCount})");

        // ---- Step 0: blocking pre-checks (read-only; always run, even in dry-run) ----
        // These drive idempotent skips (already-disabled / already-shared) and the license
        // safety gate below. The mailbox facts come from EXO (size / holds / archive); when
        // EXO isn't wired they degrade to null and the flow proceeds without those guards.
        Running("Verificaciones previas");
        var facts = await ReadMailboxFactsAsync(upn, progress, cancellationToken).ConfigureAwait(false);
        var alreadyDisabled = !user.AccountEnabled;
        var alreadyShared = facts?.IsSharedMailbox == true;
        var exceedsSize = facts?.ExceedsUnlicensedSharedLimit == true;
        var hasHold = facts?.HasBlockingHold == true;

        var notes = new List<string>();
        if (alreadyDisabled) notes.Add("cuenta ya deshabilitada");
        if (facts is null)
        {
            notes.Add("sin datos de buzón (EXO no conectado o sin buzón)");
        }
        else
        {
            notes.Add($"buzón {facts.RecipientTypeDetails}");
            if (facts.TotalItemSizeGb is { } gb) notes.Add($"{gb:N1} GB");
            if (alreadyShared) notes.Add("ya es compartido");
            if (exceedsSize) notes.Add(">50 GB: un buzón compartido sin licencia no puede superar ese límite");
            if (hasHold) notes.Add($"hold activo (litigation={facts.LitigationHoldEnabled}, in-place={facts.InPlaceHoldCount})");
            if (facts.ArchiveEnabled) notes.Add("archivo en línea habilitado");
        }
        var preWarn = exceedsSize || hasHold;
        Done("Verificaciones previas", preWarn ? "AVISO" : "OK", string.Join("; ", notes));
        if (preWarn)
        {
            progress?.Report(LogEntry.Warn("Offboarding", "Pre-check con avisos: " + string.Join("; ", notes)));
        }

        // ---- Step 1: block sign-in (+ revoke sessions) ----
        if (options.DisableAccount)
        {
            Running("Deshabilitar cuenta");
            if (dry)
            {
                Done("Deshabilitar cuenta", "SIMULADO",
                    alreadyDisabled
                        ? "ya estaba deshabilitada; se revocarían sesiones"
                        : "se deshabilitaría la cuenta y se revocarían sesiones");
            }
            else
            {
                try
                {
                    if (!alreadyDisabled)
                    {
                        await _users.SetAccountEnabledAsync(user.Id, false, progress, cancellationToken).ConfigureAwait(false);
                    }
                    // Revoke sessions even when it was already disabled — disabling alone doesn't
                    // invalidate tokens already issued.
                    var basePrefix = alreadyDisabled ? "Ya estaba deshabilitada" : "AccountEnabled=false";
                    try
                    {
                        await _users.RevokeSignInSessionsAsync(user.Id, progress, cancellationToken).ConfigureAwait(false);
                        Done("Deshabilitar cuenta", alreadyDisabled ? "OMITIDO" : "OK", $"{basePrefix}; sesiones revocadas");
                    }
                    catch (Exception revokeEx)
                    {
                        Done("Deshabilitar cuenta", alreadyDisabled ? "OMITIDO" : "OK", $"{basePrefix}; revoke sesiones falló: {revokeEx.Message}");
                    }
                }
                catch (Exception ex)
                {
                    Done("Deshabilitar cuenta", "ERROR", ex.Message);
                    success = false;
                }
            }
        }

        // ---- Step 2: convert mailbox to shared ----
        // The mailbox MUST be converted while still licensed; removing the license first starts
        // a 30-day deletion clock and hides the convert option. So convert before removing
        // licenses, and skip license removal if conversion failed (gate below).
        // EXO is reachable (external ops wired) but reported no readable mailbox → the user has
        // no EXO mailbox; skip the mailbox steps rather than erroring on each one.
        var noMailbox = facts is null && _externalExo is not null;

        var convertRequested = options.ConvertMailboxToShared;
        var convertSucceeded = false;
        if (convertRequested)
        {
            Running("Convertir a buzón compartido");
            if (noMailbox)
            {
                // No mailbox to convert — skip (license removal may still proceed; there is no
                // mailbox to strand). Matches the legacy "if ($mbox) { ... } else SKIP" flow.
                Done("Convertir a buzón compartido", "OMITIDO", "El usuario no tiene buzón en Exchange Online; se omiten los pasos de buzón.");
            }
            else if (alreadyShared)
            {
                // Idempotent: nothing to do; license removal may still proceed.
                convertSucceeded = true;
                Done("Convertir a buzón compartido", "OMITIDO", "El buzón ya es compartido");
            }
            else if (dry)
            {
                convertSucceeded = true; // simulate success so the license step can also be simulated
                Done("Convertir a buzón compartido", "SIMULADO",
                    facts is null ? "se intentaría convertir (sin datos de buzón)" : "se convertiría a buzón compartido");
            }
            else
            {
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
        }

        // ---- Step 3: remove licenses (gated) ----
        if (options.RemoveLicenses)
        {
            // Whether the mailbox will end up retained as shared (and therefore unlicensed).
            var willBeShared = alreadyShared || (convertRequested && convertSucceeded);

            // Safety gates: never strip the license when doing so would either strand the
            // mailbox for deletion or leave a non-compliant unlicensed mailbox.
            string? block = null;
            if (convertRequested && !convertSucceeded && !noMailbox)
            {
                // Convert was attempted on an existing mailbox and didn't take — stripping the
                // license now would schedule that mailbox for deletion.
                block = "Conversión a buzón compartido falló; no se quitan licencias para evitar el borrado del buzón (30 días).";
            }
            else if (willBeShared && exceedsSize)
            {
                block = $"El buzón supera 50 GB ({facts!.TotalItemSizeGb:N1} GB): un buzón compartido sin licencia no puede superar ese límite. No se quitan licencias (requiere Exchange Online Plan 2).";
            }
            else if (willBeShared && hasHold)
            {
                block = "El buzón tiene un hold activo (litigation/in-place): conservarlo requiere licencia. No se quitan licencias; valora un buzón inactivo (inactive mailbox).";
            }

            if (block is not null)
            {
                Done("Quitar licencias", "OMITIDO", block);
                if (!dry) success = false;
            }
            else if (dry)
            {
                Done("Quitar licencias", "SIMULADO",
                    user.AssignedLicenseCount == 0 ? "sin licencias asignadas" : $"se quitarían {user.AssignedLicenseCount} licencia(s) directa(s)");
            }
            else
            {
                Running("Quitar licencias");
                try
                {
                    await _users.RemoveAllLicensesAsync(user.Id, progress, cancellationToken).ConfigureAwait(false);
                    var remaining = await CountRemainingLicensesAsync(user.Id, cancellationToken).ConfigureAwait(false);
                    if (remaining <= 0)
                    {
                        Done("Quitar licencias", "OK",
                            user.AssignedLicenseCount == 0 ? "Sin licencias asignadas" : $"{user.AssignedLicenseCount} licencias retiradas (verificado)");
                    }
                    else
                    {
                        // Residual licenses after a successful call are almost always group-inherited:
                        // RemoveAllLicensesAsync only clears DIRECT assignments. Flag it clearly with the
                        // remediation instead of pretending the removal was complete.
                        Done("Quitar licencias", "AVISO",
                            $"Quedan {remaining} licencia(s) asignadas tras la retirada — probablemente heredadas de grupo. Quita al usuario del grupo de licencias en Entra ID.");
                    }
                }
                catch (Exception ex)
                {
                    Done("Quitar licencias", "ERROR", ex.Message);
                    success = false;
                }
            }
        }

        // ---- Step 4: optional mailbox finalization (EXO-only) ----
        // auto-reply / forwarding / hide-from-GAL on the (now shared) mailbox. Best-effort: a
        // failure here is reported AVISO (non-fatal) and never undoes the core offboarding.
        var wantsDelegate = !string.IsNullOrWhiteSpace(options.DelegateMailboxTo);
        var wantsAutoReply = !string.IsNullOrWhiteSpace(options.AutoReplyMessage);
        var wantsForward = !string.IsNullOrWhiteSpace(options.ForwardTo);
        var wantsHide = options.HideFromGal;
        if (wantsDelegate || wantsAutoReply || wantsForward || wantsHide)
        {
            // Mailbox exists if we have facts, it was already shared, or we just converted it.
            var mailboxExists = facts is not null || alreadyShared || (convertRequested && convertSucceeded);

            async Task FinalizeStepAsync(string name, bool requested, string simDetail, Func<Task<string>> action)
            {
                if (!requested) return;
                Running(name);
                if (_externalExo is null)
                {
                    Done(name, "OMITIDO", "requiere Exchange Online externo (no configurado)");
                    return;
                }
                if (!mailboxExists)
                {
                    Done(name, "OMITIDO", "el usuario no tiene buzón");
                    return;
                }
                if (dry)
                {
                    Done(name, "SIMULADO", simDetail);
                    return;
                }
                try
                {
                    Done(name, "OK", await action().ConfigureAwait(false));
                }
                catch (Exception ex)
                {
                    Done(name, "AVISO", ex.Message);
                }
            }

            await FinalizeStepAsync("Delegar buzón (FullAccess + SendAs)", wantsDelegate,
                $"se concedería FullAccess + SendAs a {options.DelegateMailboxTo}",
                async () => await _externalExo!.GrantDelegateAsync(upn, options.DelegateMailboxTo!, sendAs: true, progress, cancellationToken).ConfigureAwait(false)).ConfigureAwait(false);

            await FinalizeStepAsync("Auto-reply", wantsAutoReply, "se activaría un OOO permanente",
                async () =>
                {
                    await _externalExo!.SetAutoReplyAsync(upn, options.AutoReplyMessage!, progress, cancellationToken).ConfigureAwait(false);
                    return "OOO permanente activado";
                }).ConfigureAwait(false);

            await FinalizeStepAsync("Forward al delegado", wantsForward, $"se reenviaría el correo a {options.ForwardTo}",
                async () =>
                {
                    await _externalExo!.SetForwardingAsync(upn, options.ForwardTo!, progress, cancellationToken).ConfigureAwait(false);
                    return $"forward → {options.ForwardTo} (entrega doble)";
                }).ConfigureAwait(false);

            await FinalizeStepAsync("Ocultar de la GAL", wantsHide, "se ocultaría de la lista global de direcciones",
                async () => await _externalExo!.HideFromGalAsync(upn, progress, cancellationToken).ConfigureAwait(false)).ConfigureAwait(false);
        }

        return Result(success);
    }

    // Best-effort mailbox facts (size / holds / archive / type). Prefers the external EXO ops
    // (reliable) and falls back to the in-proc service. Never throws — pre-checks must degrade
    // gracefully rather than abort the flow.
    private async Task<MailboxInfo?> ReadMailboxFactsAsync(string upn, IProgress<LogEntry>? progress, CancellationToken ct)
    {
        try
        {
            return _externalExo is not null
                ? await _externalExo.GetMailboxFactsAsync(upn, progress, ct).ConfigureAwait(false)
                : await _mailboxes.GetMailboxAsync(upn, progress, ct).ConfigureAwait(false);
        }
        catch
        {
            return null;
        }
    }

    // Re-reads the user a few times (Graph license changes propagate with a small lag) and
    // returns the remaining assigned-license count: 0 once removal has taken effect, or the
    // last observed count if licenses persist (typically group-inherited). Returns 0 if the
    // user can't be re-read — never blocks the flow on a verification read failure.
    private async Task<int> CountRemainingLicensesAsync(string userId, CancellationToken ct)
    {
        var last = 0;
        for (var attempt = 0; attempt < VerifyAttempts; attempt++)
        {
            try
            {
                var u = await _users.GetByIdAsync(userId, ct).ConfigureAwait(false);
                last = u?.AssignedLicenseCount ?? 0;
                if (last == 0)
                {
                    return 0;
                }
            }
            catch
            {
                return 0; // can't verify → don't raise a false warning
            }
            if (attempt < VerifyAttempts - 1)
            {
                await Task.Delay(VerifyPollDelay, ct).ConfigureAwait(false);
            }
        }
        return last;
    }
}
