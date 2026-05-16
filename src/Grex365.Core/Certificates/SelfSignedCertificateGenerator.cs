using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.Core.Certificates;

public sealed class SelfSignedCertificateGenerator : ICertificateGenerator
{
    public GeneratedCertificate GenerateAndStore(string commonName, int validDays, string exportDirectory, IProgress<LogEntry>? progress = null)
    {
        if (string.IsNullOrWhiteSpace(commonName))
        {
            throw new ArgumentException("CN no puede estar vacío.", nameof(commonName));
        }
        if (validDays <= 0)
        {
            throw new ArgumentException("validDays debe ser > 0.", nameof(validDays));
        }
        if (string.IsNullOrWhiteSpace(exportDirectory))
        {
            throw new ArgumentException("Directorio de exportación inválido.", nameof(exportDirectory));
        }

        Directory.CreateDirectory(exportDirectory);

        progress?.Report(LogEntry.Info("Cert", $"Generando RSA 2048 para CN={commonName}..."));

        using var rsa = RSA.Create(2048);
        var subject = $"CN={commonName}";
        var req = new CertificateRequest(subject, rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);

        req.CertificateExtensions.Add(new X509KeyUsageExtension(
            X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.KeyEncipherment,
            critical: true));
        req.CertificateExtensions.Add(new X509EnhancedKeyUsageExtension(
            new OidCollection { new Oid("1.3.6.1.5.5.7.3.2") }, // client auth
            critical: false));
        req.CertificateExtensions.Add(new X509SubjectKeyIdentifierExtension(req.PublicKey, critical: false));

        var notBefore = DateTimeOffset.Now.AddMinutes(-5);
        var notAfter = DateTimeOffset.Now.AddDays(validDays);

        using var cert = req.CreateSelfSigned(notBefore, notAfter);

        progress?.Report(LogEntry.Info("Cert", "Exportando a PFX (clave privada en memoria)..."));
        var pfx = cert.Export(X509ContentType.Pfx);
        using var importable = X509CertificateLoader.LoadPkcs12(pfx, password: null,
            keyStorageFlags: X509KeyStorageFlags.PersistKeySet | X509KeyStorageFlags.UserKeySet | X509KeyStorageFlags.Exportable);

        progress?.Report(LogEntry.Info("Cert", "Instalando en Cert:\\CurrentUser\\My..."));
        using (var store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
        {
            store.Open(OpenFlags.ReadWrite);
            store.Add(importable);
            store.Close();
        }

        var cerFileName = $"grex365-{SanitizeFileName(commonName)}-{importable.Thumbprint}.cer";
        var cerPath = Path.Combine(exportDirectory, cerFileName);
        var cerBytes = importable.Export(X509ContentType.Cert);
        File.WriteAllBytes(cerPath, cerBytes);

        progress?.Report(LogEntry.Ok("Cert", $"Cert generado. Thumbprint={importable.Thumbprint}"));
        progress?.Report(LogEntry.Info("Cert", $".cer exportado a {cerPath} (subir al App Registration en Azure AD)"));

        return new GeneratedCertificate(
            Subject: importable.Subject,
            Thumbprint: importable.Thumbprint,
            NotBefore: importable.NotBefore,
            NotAfter: importable.NotAfter,
            CerPath: cerPath);
    }

    private static string SanitizeFileName(string input)
    {
        foreach (var c in Path.GetInvalidFileNameChars())
        {
            input = input.Replace(c, '_');
        }
        return input;
    }
}
