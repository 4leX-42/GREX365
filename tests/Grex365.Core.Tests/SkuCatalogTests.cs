using FluentAssertions;
using Grex365.Core.Models;

namespace Grex365.Core.Tests;

public class SkuCatalogTests
{
    [Theory]
    [InlineData("SPE_E5", "Microsoft 365 E5", LicenseCategory.Enterprise)]
    [InlineData("SPE_E3", "Microsoft 365 E3", LicenseCategory.Enterprise)]
    [InlineData("ENTERPRISEPREMIUM", "Office 365 E5", LicenseCategory.Enterprise)]
    [InlineData("SPB", "Microsoft 365 Business Premium", LicenseCategory.Business)]
    [InlineData("O365_BUSINESS_PREMIUM", "Microsoft 365 Business Standard", LicenseCategory.Business)]
    [InlineData("SPE_F3", "Microsoft 365 F3", LicenseCategory.Frontline)]
    [InlineData("AAD_PREMIUM_P2", "Microsoft Entra ID P2", LicenseCategory.Security)]
    [InlineData("VISIOCLIENT", "Visio Plan 2", LicenseCategory.AppOrAddOn)]
    public void Resolve_KnownSkus_ReturnsFriendlyName(string sku, string expectedName, LicenseCategory category)
    {
        var info = SkuCatalog.Resolve(sku);
        info.FriendlyName.Should().Be(expectedName);
        info.Category.Should().Be(category);
    }

    [Fact]
    public void Resolve_UnknownSku_FallsBackToHumanized()
    {
        var info = SkuCatalog.Resolve("CUSTOM_THING_FOR_TENANT");
        info.Category.Should().Be(LicenseCategory.Other);
        info.FriendlyName.Should().Be("Custom Thing For Tenant");
        info.Priority.Should().BeGreaterThan(5000);
    }

    [Fact]
    public void Resolve_EmptySku_GracefulFallback()
    {
        var info = SkuCatalog.Resolve("");
        info.FriendlyName.Should().Be("(SKU sin nombre)");
        info.Category.Should().Be(LicenseCategory.Other);
    }

    [Fact]
    public void Resolve_IsCaseInsensitive()
    {
        var lower = SkuCatalog.Resolve("spe_e5");
        var upper = SkuCatalog.Resolve("SPE_E5");
        lower.Should().Be(upper);
    }

    [Fact]
    public void Enterprise_HasHigherPriority_ThanBusiness()
    {
        SkuCatalog.Resolve("SPE_E5").Priority.Should().BeLessThan(SkuCatalog.Resolve("SPB").Priority);
        SkuCatalog.Resolve("SPB").Priority.Should().BeLessThan(SkuCatalog.Resolve("SPE_F3").Priority);
        SkuCatalog.Resolve("SPE_F3").Priority.Should().BeLessThan(SkuCatalog.Resolve("AAD_PREMIUM").Priority);
    }

    [Fact]
    public void CategoryLabel_Mapping()
    {
        SkuCatalog.CategoryLabel(LicenseCategory.Enterprise).Should().Be("Enterprise");
        SkuCatalog.CategoryLabel(LicenseCategory.Business).Should().Be("Business");
        SkuCatalog.CategoryLabel(LicenseCategory.Frontline).Should().Be("Frontline");
        SkuCatalog.CategoryLabel(LicenseCategory.Security).Should().Be("Security e Identity");
        SkuCatalog.CategoryLabel(LicenseCategory.AppOrAddOn).Should().Be("Apps y complementos");
        SkuCatalog.CategoryLabel(LicenseCategory.Other).Should().Be("Otros");
    }
}
