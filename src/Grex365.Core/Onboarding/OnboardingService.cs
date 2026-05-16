using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.Core.Onboarding;

public sealed class OnboardingService : IOnboardingService
{
    private readonly IUsersService _users;
    private readonly IGroupsService _groups;

    public OnboardingService(IUsersService users, IGroupsService groups)
    {
        _users = users;
        _groups = groups;
    }

    public async Task<OnboardingResult> RunAsync(
        OnboardingOptions options,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var steps = new List<OnboardingStep>();
        var upn = (options.Upn ?? string.Empty).Trim();

        var validation = OnboardingValidator.Validate(options);
        if (validation.Count > 0)
        {
            steps.Add(new OnboardingStep("Validar", "ERROR", string.Join("; ", validation)));
            return new OnboardingResult(upn, null, false, steps);
        }
        steps.Add(new OnboardingStep("Validar", "OK", "Entradas válidas."));

        progress?.Report(LogEntry.Info("Onboarding", $"Iniciando onboarding de {upn}"));

        UserSummary created;
        try
        {
            var spec = new NewUserSpec(
                DisplayName: options.DisplayName.Trim(),
                UserPrincipalName: upn,
                MailNickname: OnboardingValidator.DeriveMailNickname(upn, options.MailNickname),
                Password: options.InitialPassword,
                UsageLocation: options.UsageLocation.Trim().ToUpperInvariant(),
                ForceChangePasswordNextSignIn: options.ForceChangePasswordNextSignIn);
            created = await _users.CreateUserAsync(spec, progress, cancellationToken).ConfigureAwait(false);
            steps.Add(new OnboardingStep("Crear usuario", "OK",
                $"{created.DisplayName} · id={created.Id}"));
        }
        catch (Exception ex)
        {
            steps.Add(new OnboardingStep("Crear usuario", "ERROR", ex.Message));
            return new OnboardingResult(upn, null, false, steps);
        }

        var success = true;

        var skuIds = options.SkuIds ?? Array.Empty<Guid>();
        foreach (var sku in skuIds)
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                await _users.AssignLicenseAsync(created.Id, sku, progress, cancellationToken).ConfigureAwait(false);
                steps.Add(new OnboardingStep("Asignar licencia", "OK", sku.ToString()));
            }
            catch (Exception ex)
            {
                steps.Add(new OnboardingStep("Asignar licencia", "ERROR", $"{sku}: {ex.Message}"));
                success = false;
            }
        }

        var groupIds = options.GroupIdentifiers ?? Array.Empty<string>();
        foreach (var groupRaw in groupIds)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var groupKey = (groupRaw ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(groupKey))
            {
                continue;
            }
            try
            {
                var groupId = await ResolveGroupIdAsync(groupKey, cancellationToken).ConfigureAwait(false);
                if (groupId is null)
                {
                    steps.Add(new OnboardingStep("Añadir a grupo", "ERROR", $"No resuelto: {groupKey}"));
                    success = false;
                    continue;
                }
                var addResults = await _groups.AddMembersAsync(groupId, new[] { created.UserPrincipalName ?? upn }, progress, cancellationToken).ConfigureAwait(false);
                var r = addResults.FirstOrDefault();
                var status = r?.Status ?? "ERROR";
                if (status is "AGREGADO" or "YA_EXISTE")
                {
                    steps.Add(new OnboardingStep("Añadir a grupo", "OK", $"{groupKey} · {status}"));
                }
                else
                {
                    steps.Add(new OnboardingStep("Añadir a grupo", "ERROR", $"{groupKey} · {r?.Detail}"));
                    success = false;
                }
            }
            catch (Exception ex)
            {
                steps.Add(new OnboardingStep("Añadir a grupo", "ERROR", $"{groupKey}: {ex.Message}"));
                success = false;
            }
        }

        return new OnboardingResult(upn, created.Id, success, steps);
    }

    private async Task<string?> ResolveGroupIdAsync(string key, CancellationToken cancellationToken)
    {
        if (Guid.TryParse(key, out _))
        {
            return key;
        }
        var matches = await _groups.SearchAsync(key, cancellationToken).ConfigureAwait(false);
        var hit = matches.FirstOrDefault(g =>
            string.Equals(g.Mail, key, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(g.DisplayName, key, StringComparison.OrdinalIgnoreCase))
            ?? matches.FirstOrDefault();
        return hit?.Id;
    }
}
