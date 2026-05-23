using FluentAssertions;
using Grex365.App.ViewModels;

namespace Grex365.App.Tests;

public class NavTitleMigratorTests
{
    [Fact]
    public void Resolve_Null_ReturnsNull()
    {
        NavTitleMigrator.Resolve(null).Should().BeNull();
    }

    [Fact]
    public void Resolve_Empty_ReturnsNull()
    {
        NavTitleMigrator.Resolve(string.Empty).Should().BeNull();
    }

    [Fact]
    public void Resolve_Whitespace_ReturnsNull()
    {
        NavTitleMigrator.Resolve("   ").Should().BeNull();
    }

    [Theory]
    [InlineData("Conexion", "Conexión")]
    [InlineData("Auditoria", "Auditoría")]
    [InlineData("Reglas buzon", "Reglas de buzón")]
    [InlineData("Salud tenant", "Licencias")]
    [InlineData("Mail flow", "Flujo de correo")]
    [InlineData("Audit log", "Registro de auditoría")]
    [InlineData("Cert Wizard", "Asistente cert")]
    [InlineData("DNS check", "Comprobación DNS")]
    public void Resolve_RenamedTitle_ReturnsNewTitle(string saved, string expected)
    {
        NavTitleMigrator.Resolve(saved).Should().Be(expected);
    }

    [Theory]
    [InlineData("conexion")]
    [InlineData("CONEXION")]
    [InlineData("Salud Tenant")]
    [InlineData("CERT WIZARD")]
    public void Resolve_RenamedTitle_CaseInsensitive(string saved)
    {
        NavTitleMigrator.Resolve(saved).Should().NotBeNullOrEmpty().And.NotBe(saved);
    }

    [Theory]
    [InlineData("Dashboard")]
    [InlineData("Usuarios")]
    [InlineData("Grupos")]
    [InlineData("Conexión")]
    [InlineData("Licencias")]
    public void Resolve_NotRenamed_ReturnsAsIs(string saved)
    {
        NavTitleMigrator.Resolve(saved).Should().Be(saved);
    }

    [Fact]
    public void Resolve_UnknownTitle_ReturnsAsIs()
    {
        NavTitleMigrator.Resolve("My Custom Plugin").Should().Be("My Custom Plugin");
    }

    [Fact]
    public void RenameMap_ContainsAllExpectedKeys()
    {
        var keys = NavTitleMigrator.RenameMap.Keys;
        keys.Should().Contain([
            "Conexion", "Auditoria", "Reglas buzon", "Salud tenant",
            "Mail flow", "Audit log", "Cert Wizard", "DNS check"
        ]);
    }

    [Fact]
    public void RenameMap_NewValues_AllResolveToSelf()
    {
        foreach (var (_, newTitle) in NavTitleMigrator.RenameMap)
        {
            NavTitleMigrator.Resolve(newTitle).Should().Be(newTitle,
                because: $"the new title '{newTitle}' should not be mapped again");
        }
    }

    [Fact]
    public void RenameMap_NoCycles()
    {
        foreach (var (oldTitle, newTitle) in NavTitleMigrator.RenameMap)
        {
            NavTitleMigrator.RenameMap.ContainsKey(newTitle).Should().BeFalse(
                because: $"the new title '{newTitle}' (from '{oldTitle}') should not appear as another old title");
        }
    }
}
