namespace Grex365.Core.Models;

public sealed record CertConfig(
    string AppId,
    string TenantId,
    string Organization,
    string CertThumbprint);
