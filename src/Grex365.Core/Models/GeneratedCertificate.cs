namespace Grex365.Core.Models;

public sealed record GeneratedCertificate(
    string Subject,
    string Thumbprint,
    DateTimeOffset NotBefore,
    DateTimeOffset NotAfter,
    string CerPath);
