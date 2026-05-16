using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

public interface ICertificateGenerator
{
    GeneratedCertificate GenerateAndStore(string commonName, int validDays, string exportDirectory, IProgress<LogEntry>? progress = null);
}
