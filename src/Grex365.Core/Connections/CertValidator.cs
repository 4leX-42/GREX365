using System.Security.Cryptography.X509Certificates;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.Core.Connections;

public sealed class CertValidator : ICertValidator
{
    public CertValidationResult Validate(CertConfig? config)
    {
        if (config is null
            || string.IsNullOrWhiteSpace(config.CertThumbprint)
            || string.IsNullOrWhiteSpace(config.AppId)
            || string.IsNullOrWhiteSpace(config.TenantId))
        {
            return new CertValidationResult(
                CertValidationStatus.MissingConfig,
                "Falta configuración (thumbprint, AppId o TenantId).");
        }

        using var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
        store.Open(OpenFlags.ReadOnly);
        var found = store.Certificates.Find(X509FindType.FindByThumbprint, config.CertThumbprint, validOnly: false);
        if (found.Count == 0)
        {
            return new CertValidationResult(
                CertValidationStatus.MissingFromStore,
                $"Certificado {config.CertThumbprint} no está en Cert:\\CurrentUser\\My.");
        }

        var cert = found[0];
        var now = DateTimeOffset.Now;

        if (now < cert.NotBefore)
        {
            return new CertValidationResult(
                CertValidationStatus.NotYetValid,
                $"Certificado aún no válido (NotBefore = {cert.NotBefore:yyyy-MM-dd}).",
                NotAfter: cert.NotAfter,
                NotBefore: cert.NotBefore);
        }

        if (now > cert.NotAfter)
        {
            return new CertValidationResult(
                CertValidationStatus.Expired,
                $"Certificado expirado el {cert.NotAfter:yyyy-MM-dd}.",
                NotAfter: cert.NotAfter,
                NotBefore: cert.NotBefore);
        }

        if (!cert.HasPrivateKey)
        {
            return new CertValidationResult(
                CertValidationStatus.NoPrivateKey,
                "Certificado encontrado pero sin clave privada.",
                NotAfter: cert.NotAfter,
                NotBefore: cert.NotBefore);
        }

        return new CertValidationResult(
            CertValidationStatus.Ok,
            $"Cert OK · expira {cert.NotAfter:yyyy-MM-dd}.",
            NotAfter: cert.NotAfter,
            NotBefore: cert.NotBefore);
    }
}
