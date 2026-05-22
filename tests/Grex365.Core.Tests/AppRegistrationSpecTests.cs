using FluentAssertions;
using Grex365.Core.Connections;

namespace Grex365.Core.Tests;

public class AppRegistrationSpecTests
{
    [Fact]
    public void BuildRequiredResourceAccess_ReturnsGraphAndExoEntries()
    {
        var access = AppRegistrationSpec.BuildRequiredResourceAccess();

        access.Should().HaveCount(2);
        access[0].ResourceAppId.Should().Be(AppRegistrationSpec.GraphResourceId);
        access[1].ResourceAppId.Should().Be(AppRegistrationSpec.ExoResourceId);
    }

    [Fact]
    public void BuildRequiredResourceAccess_GraphHasNineAppRoles()
    {
        var access = AppRegistrationSpec.BuildRequiredResourceAccess();
        var graph = access[0].ResourceAccess!;

        graph.Should().HaveCount(9);
        graph.Should().AllSatisfy(r => r.Type.Should().Be("Role"));
        graph.Select(r => r.Id).Should().OnlyHaveUniqueItems();
    }

    [Fact]
    public void BuildRequiredResourceAccess_ExoHasManageAsAppOnly()
    {
        var access = AppRegistrationSpec.BuildRequiredResourceAccess();
        var exo = access[1].ResourceAccess!;

        exo.Should().HaveCount(1);
        exo[0].Type.Should().Be("Role");
        exo[0].Id.Should().Be(Guid.Parse("dc50a0fb-09a3-484d-be87-e023b12c6440"));
    }

    [Fact]
    public void GraphAppRoles_AllIdsAreValidGuids()
    {
        foreach (var (id, _) in AppRegistrationSpec.GraphAppRoles)
        {
            Guid.TryParse(id, out _).Should().BeTrue($"role id '{id}' should be a valid GUID");
        }
    }

    [Fact]
    public void GraphAppRoles_IncludesCriticalPermissions()
    {
        var names = AppRegistrationSpec.GraphAppRoles.Select(r => r.Name).ToList();

        names.Should().Contain(new[]
        {
            "User.ReadWrite.All",
            "Group.ReadWrite.All",
            "Organization.Read.All",
            "AuditLog.Read.All",
            "Directory.ReadWrite.All",
            "Reports.Read.All",
            "Policy.Read.All",
            "Application.Read.All",
        });
    }

    [Theory]
    [InlineData("00000000-0000-0000-0000-000000000001", "11111111-2222-3333-4444-555555555555")]
    [InlineData("contoso.onmicrosoft.com", "abc-def")]
    public void BuildAdminConsentUrl_FormatsExpectedShape(string tenantId, string appId)
    {
        var url = AppRegistrationSpec.BuildAdminConsentUrl(tenantId, appId);

        url.Should().Be($"https://login.microsoftonline.com/{tenantId}/adminconsent?client_id={appId}");
    }

    [Fact]
    public void BuildAdminConsentUrl_EmptyTenant_Throws()
    {
        var act = () => AppRegistrationSpec.BuildAdminConsentUrl("", "appid");
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void BuildAdminConsentUrl_EmptyAppId_Throws()
    {
        var act = () => AppRegistrationSpec.BuildAdminConsentUrl("tenant", "  ");
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void BuildApplication_SetsExpectedShape()
    {
        var cer = new byte[] { 1, 2, 3, 4 };
        var app = AppRegistrationSpec.BuildApplication("Grex365", cer, "ABCDEF1234567890");

        app.DisplayName.Should().Be("Grex365");
        app.SignInAudience.Should().Be("AzureADMyOrg");
        app.RequiredResourceAccess.Should().HaveCount(2);
        app.KeyCredentials.Should().ContainSingle();
        var key = app.KeyCredentials![0];
        key.Type.Should().Be("AsymmetricX509Cert");
        key.Usage.Should().Be("Verify");
        key.Key.Should().BeEquivalentTo(cer);
        key.DisplayName.Should().Contain("ABCDEF12"); // first 8 chars of thumbprint
    }

    [Fact]
    public void BuildApplication_ShortThumbprint_TruncatesGracefully()
    {
        var app = AppRegistrationSpec.BuildApplication("X", new byte[] { 1 }, "ABC");
        app.KeyCredentials![0].DisplayName.Should().Contain("ABC");
    }

    [Fact]
    public void BuildApplication_EmptyDisplayName_Throws()
    {
        var act = () => AppRegistrationSpec.BuildApplication("  ", new byte[] { 1 }, "thumb");
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void BuildApplication_NullCert_Throws()
    {
        var act = () => AppRegistrationSpec.BuildApplication("X", null!, "thumb");
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void BuildApplication_EmptyCertArray_Throws()
    {
        var act = () => AppRegistrationSpec.BuildApplication("X", Array.Empty<byte>(), "thumb");
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void BuildApplication_EmptyThumbprint_Throws()
    {
        var act = () => AppRegistrationSpec.BuildApplication("X", new byte[] { 1 }, "");
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void BuildCertLabel_ShorterThan8Chars_ReturnsFullString()
    {
        AppRegistrationSpec.BuildCertLabel("ABC").Should().Be("ABC");
    }

    [Fact]
    public void BuildCertLabel_LongerThan8Chars_TruncatesTo8()
    {
        AppRegistrationSpec.BuildCertLabel("ABCDEF1234567890").Should().Be("ABCDEF12");
    }
}
