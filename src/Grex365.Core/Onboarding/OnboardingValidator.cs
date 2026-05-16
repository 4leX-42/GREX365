using Grex365.Core.Abstractions;

namespace Grex365.Core.Onboarding;

public static class OnboardingValidator
{
    public static IReadOnlyList<string> Validate(OnboardingOptions options)
    {
        var errors = new List<string>();

        if (string.IsNullOrWhiteSpace(options.DisplayName))
        {
            errors.Add("DisplayName requerido.");
        }
        if (string.IsNullOrWhiteSpace(options.Upn))
        {
            errors.Add("UPN requerido.");
        }
        else if (!IsValidUpn(options.Upn))
        {
            errors.Add($"UPN inválido: {options.Upn}");
        }
        if (string.IsNullOrEmpty(options.InitialPassword) || options.InitialPassword.Length < 8)
        {
            errors.Add("Password requerido (mínimo 8 caracteres).");
        }
        if (string.IsNullOrWhiteSpace(options.UsageLocation) || options.UsageLocation.Length != 2)
        {
            errors.Add("UsageLocation requerido (código ISO-2, ej: ES).");
        }
        return errors;
    }

    public static string DeriveMailNickname(string upn, string? explicitNickname)
    {
        if (!string.IsNullOrWhiteSpace(explicitNickname))
        {
            return explicitNickname.Trim();
        }
        if (string.IsNullOrWhiteSpace(upn))
        {
            return string.Empty;
        }
        var at = upn.IndexOf('@');
        return at > 0 ? upn[..at] : upn;
    }

    private static bool IsValidUpn(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }
        var at = value.IndexOf('@');
        if (at <= 0 || at == value.Length - 1)
        {
            return false;
        }
        var dot = value.IndexOf('.', at);
        return dot > at + 1 && dot < value.Length - 1;
    }
}
