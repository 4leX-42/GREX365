namespace Grex365.Core.Models;

public sealed record DomainCheck(
    string Domain,
    IReadOnlyList<DnsRecord> Records);

public sealed record DnsRecord(
    string Type,
    string Status,
    string Value);
