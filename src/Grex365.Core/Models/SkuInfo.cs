namespace Grex365.Core.Models;

public sealed record SkuInfo(
    Guid SkuId,
    string SkuPartNumber,
    int Enabled,
    int Consumed)
{
    public int Available => Math.Max(0, Enabled - Consumed);

    public string Display => string.IsNullOrEmpty(SkuPartNumber)
        ? SkuId.ToString()
        : $"{SkuPartNumber}  ({Available}/{Enabled} libres)";
}
