using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

public interface ICertValidator
{
    CertValidationResult Validate(CertConfig? config);
}

public enum CertValidationStatus
{
    Ok,
    MissingConfig,
    MissingFromStore,
    Expired,
    NotYetValid,
    NoPrivateKey
}

public sealed record CertValidationResult(
    CertValidationStatus Status,
    string Message,
    DateTimeOffset? NotAfter = null,
    DateTimeOffset? NotBefore = null)
{
    public bool IsValid => Status == CertValidationStatus.Ok;
}
