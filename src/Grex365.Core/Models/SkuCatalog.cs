namespace Grex365.Core.Models;

public enum LicenseCategory
{
    Enterprise,
    Business,
    Frontline,
    Security,
    AppOrAddOn,
    Other,
}

public sealed record LicenseSkuInfo(
    string SkuPartNumber,
    string FriendlyName,
    LicenseCategory Category,
    int Priority);

public static class SkuCatalog
{
    // Curated map of common Microsoft 365 SKU part numbers to friendly names + category.
    // Reference: https://learn.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
    private static readonly Dictionary<string, LicenseSkuInfo> Map = new(StringComparer.OrdinalIgnoreCase)
    {
        // Enterprise plans (highest priority)
        ["SPE_E5"]            = new("SPE_E5",            "Microsoft 365 E5",                  LicenseCategory.Enterprise, 10),
        ["SPE_E3"]            = new("SPE_E3",            "Microsoft 365 E3",                  LicenseCategory.Enterprise, 20),
        ["ENTERPRISEPREMIUM"] = new("ENTERPRISEPREMIUM", "Office 365 E5",                     LicenseCategory.Enterprise, 30),
        ["ENTERPRISEPACK"]    = new("ENTERPRISEPACK",    "Office 365 E3",                     LicenseCategory.Enterprise, 40),
        ["STANDARDPACK"]      = new("STANDARDPACK",      "Office 365 E1",                     LicenseCategory.Enterprise, 50),
        ["DEVELOPERPACK_E5"]  = new("DEVELOPERPACK_E5",  "Office 365 E5 Developer",           LicenseCategory.Enterprise, 60),

        // Business plans
        ["SPB"]                       = new("SPB",                       "Microsoft 365 Business Premium",  LicenseCategory.Business, 100),
        ["O365_BUSINESS_PREMIUM"]     = new("O365_BUSINESS_PREMIUM",     "Microsoft 365 Business Standard", LicenseCategory.Business, 110),
        ["O365_BUSINESS_ESSENTIALS"]  = new("O365_BUSINESS_ESSENTIALS",  "Microsoft 365 Business Basic",    LicenseCategory.Business, 120),
        ["O365_BUSINESS"]             = new("O365_BUSINESS",             "Microsoft 365 Apps for Business", LicenseCategory.Business, 130),
        ["MICROSOFT_BUSINESS_CENTER"] = new("MICROSOFT_BUSINESS_CENTER", "Microsoft Business Center",       LicenseCategory.Business, 140),

        // Frontline
        ["SPE_F1"] = new("SPE_F1", "Microsoft 365 F1",     LicenseCategory.Frontline, 200),
        ["SPE_F3"] = new("SPE_F3", "Microsoft 365 F3",     LicenseCategory.Frontline, 210),

        // Security / Identity / EMS
        ["EMS"]                = new("EMS",                "Enterprise Mobility + Security E3",     LicenseCategory.Security, 300),
        ["EMSPREMIUM"]         = new("EMSPREMIUM",         "Enterprise Mobility + Security E5",     LicenseCategory.Security, 310),
        ["AAD_PREMIUM"]        = new("AAD_PREMIUM",        "Microsoft Entra ID P1",                 LicenseCategory.Security, 320),
        ["AAD_PREMIUM_P2"]     = new("AAD_PREMIUM_P2",     "Microsoft Entra ID P2",                 LicenseCategory.Security, 330),
        ["INTUNE_A"]           = new("INTUNE_A",           "Microsoft Intune",                      LicenseCategory.Security, 340),
        ["INTUNE_A_D"]         = new("INTUNE_A_D",         "Microsoft Intune Device",               LicenseCategory.Security, 345),
        ["DEFENDER_ENDPOINT_P1"] = new("DEFENDER_ENDPOINT_P1", "Defender for Endpoint P1",          LicenseCategory.Security, 350),
        ["DEFENDER_ENDPOINT_P2"] = new("DEFENDER_ENDPOINT_P2", "Defender for Endpoint P2",          LicenseCategory.Security, 351),
        ["ATP_ENTERPRISE"]     = new("ATP_ENTERPRISE",     "Defender for Office 365 P1",            LicenseCategory.Security, 360),
        ["THREAT_INTELLIGENCE"]= new("THREAT_INTELLIGENCE","Defender for Office 365 P2",            LicenseCategory.Security, 361),
        ["WIN10_PRO_ENT_SUB"]  = new("WIN10_PRO_ENT_SUB",  "Windows 10/11 Enterprise E3",           LicenseCategory.Security, 370),
        ["WIN10_VDA_E5"]       = new("WIN10_VDA_E5",       "Windows 10/11 Enterprise E5",           LicenseCategory.Security, 371),

        // Apps / add-ons
        ["EXCHANGESTANDARD"]    = new("EXCHANGESTANDARD",    "Exchange Online (Plan 1)",      LicenseCategory.AppOrAddOn, 500),
        ["EXCHANGEENTERPRISE"]  = new("EXCHANGEENTERPRISE",  "Exchange Online (Plan 2)",      LicenseCategory.AppOrAddOn, 510),
        ["EXCHANGE_S_ARCHIVE_ADDON_GOV"] = new("EXCHANGE_S_ARCHIVE_ADDON_GOV", "Exchange Online Archiving", LicenseCategory.AppOrAddOn, 515),
        ["TEAMS_EXPLORATORY"]   = new("TEAMS_EXPLORATORY",   "Microsoft Teams Exploratory",   LicenseCategory.AppOrAddOn, 520),
        ["MCOMEETADV"]          = new("MCOMEETADV",          "Audio Conferencing",            LicenseCategory.AppOrAddOn, 530),
        ["MCOEV"]               = new("MCOEV",               "Microsoft Teams Phone",         LicenseCategory.AppOrAddOn, 531),
        ["POWER_BI_PRO"]        = new("POWER_BI_PRO",        "Power BI Pro",                  LicenseCategory.AppOrAddOn, 540),
        ["POWER_BI_STANDARD"]   = new("POWER_BI_STANDARD",   "Power BI (free)",               LicenseCategory.AppOrAddOn, 541),
        ["PBI_PREMIUM_PER_USER"]= new("PBI_PREMIUM_PER_USER","Power BI Premium per User",     LicenseCategory.AppOrAddOn, 542),
        ["VISIOCLIENT"]         = new("VISIOCLIENT",         "Visio Plan 2",                  LicenseCategory.AppOrAddOn, 550),
        ["VISIO_PLAN1_DEPT"]    = new("VISIO_PLAN1_DEPT",    "Visio Plan 1",                  LicenseCategory.AppOrAddOn, 551),
        ["PROJECTPROFESSIONAL"] = new("PROJECTPROFESSIONAL", "Project Plan 3",                LicenseCategory.AppOrAddOn, 560),
        ["PROJECTPREMIUM"]      = new("PROJECTPREMIUM",      "Project Plan 5",                LicenseCategory.AppOrAddOn, 561),
        ["FLOW_FREE"]           = new("FLOW_FREE",           "Power Automate Free",           LicenseCategory.AppOrAddOn, 570),
        ["POWERAPPS_VIRAL"]     = new("POWERAPPS_VIRAL",     "Power Apps (Viral)",            LicenseCategory.AppOrAddOn, 571),
        ["WINDOWS_STORE"]       = new("WINDOWS_STORE",       "Windows Store for Business",    LicenseCategory.AppOrAddOn, 580),
    };

    public static LicenseSkuInfo Resolve(string skuPartNumber)
    {
        if (Map.TryGetValue(skuPartNumber, out var info))
        {
            return info;
        }
        return new LicenseSkuInfo(
            SkuPartNumber: skuPartNumber,
            FriendlyName: HumanizeSkuPart(skuPartNumber),
            Category: LicenseCategory.Other,
            Priority: 9000);
    }

    public static string CategoryLabel(LicenseCategory category) => category switch
    {
        LicenseCategory.Enterprise => "Enterprise",
        LicenseCategory.Business => "Business",
        LicenseCategory.Frontline => "Frontline",
        LicenseCategory.Security => "Security e Identity",
        LicenseCategory.AppOrAddOn => "Apps y complementos",
        _ => "Otros",
    };

    private static string HumanizeSkuPart(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "(SKU sin nombre)";
        // Normalize: replace underscores with spaces, lowercase, title case words.
        var words = raw.Replace('_', ' ').ToLowerInvariant().Split(' ', StringSplitOptions.RemoveEmptyEntries);
        for (var i = 0; i < words.Length; i++)
        {
            words[i] = words[i] switch
            {
                "spe" or "sp" or "spb" => words[i].ToUpperInvariant(),
                "e5" or "e3" or "e1" or "f1" or "f3" or "p1" or "p2" => words[i].ToUpperInvariant(),
                _ => char.ToUpperInvariant(words[i][0]) + words[i][1..],
            };
        }
        return string.Join(' ', words);
    }
}
