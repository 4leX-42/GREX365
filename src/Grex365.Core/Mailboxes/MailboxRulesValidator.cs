using Grex365.Core.Models;

namespace Grex365.Core.Mailboxes;

public static class MailboxRulesValidator
{
    public static IReadOnlyList<string> ValidateAutoReply(AutoReplyConfig config)
    {
        var errors = new List<string>();

        if (config.State == AutoReplyState.Disabled)
        {
            return errors;
        }

        var hasInternal = !string.IsNullOrWhiteSpace(config.InternalMessage);
        var hasExternal = !string.IsNullOrWhiteSpace(config.ExternalMessage);
        if (!hasInternal && !hasExternal)
        {
            errors.Add("Indica al menos un mensaje (interno o externo).");
        }

        if (config.State == AutoReplyState.Scheduled)
        {
            if (!config.StartTime.HasValue)
            {
                errors.Add("StartTime requerido cuando State=Scheduled.");
            }
            if (!config.EndTime.HasValue)
            {
                errors.Add("EndTime requerido cuando State=Scheduled.");
            }
            if (config.StartTime.HasValue && config.EndTime.HasValue && config.EndTime.Value <= config.StartTime.Value)
            {
                errors.Add("EndTime debe ser posterior a StartTime.");
            }
        }

        return errors;
    }

    public static IReadOnlyList<string> ValidateForwarding(string? smtpAddress)
    {
        var errors = new List<string>();
        if (string.IsNullOrWhiteSpace(smtpAddress))
        {
            errors.Add("SMTP de destino requerido.");
            return errors;
        }
        if (!IsEmail(smtpAddress))
        {
            errors.Add($"SMTP inválido: {smtpAddress}");
        }
        return errors;
    }

    private static bool IsEmail(string value)
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
