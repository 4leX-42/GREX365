using FluentAssertions;
using Grex365.App;
using Xunit;

namespace Grex365.App.Tests;

// Locks the wording + format behaviour of the ViewModel status / dialog keys
// introduced in Sprint AM. Parity (EN exists for every ES key) is already covered
// by L10nTests.EnDict_HasTranslationForEveryEsKey; this guards specific values so a
// future edit can't silently change a message a VM test relies on.
[Collection("L10n")]
public class VmStatusL10nTests : IDisposable
{
    public VmStatusL10nTests() => L10n.Initialize("es");
    public void Dispose() => L10n.Initialize("es");

    [Theory]
    [InlineData("Common.Status.Cancelled", "Cancelado.")]
    [InlineData("Common.Status.CancelledByUser", "Cancelado por el usuario.")]
    [InlineData("Common.Status.NoResultsToExport", "Sin resultados para exportar.")]
    [InlineData("Common.Confirm.Title", "Confirmar")]
    [InlineData("Users.Status.SelectUser", "Selecciona un usuario.")]
    [InlineData("Users.Status.Enabled", "Habilitado.")]
    [InlineData("Users.Status.Disabled", "Deshabilitado.")]
    [InlineData("Groups.Status.SelectMember", "Selecciona un miembro.")]
    [InlineData("MailboxRules.Status.EmptyMailbox", "Buzón vacío.")]
    [InlineData("MailboxRules.Status.ForwardingApplied", "Reenvío aplicado.")]
    [InlineData("Offboarding.Status.EmptyUpn", "UPN vacío.")]
    [InlineData("PsConsole.Status.EmptyCommand", "Comando vacío.")]
    [InlineData("PsConsole.Status.OutputCleared", "Salida limpiada.")]
    [InlineData("DomainCheck.Status.EmptyDomain", "Dominio vacío.")]
    [InlineData("SharedMailbox.Status.NotFound", "Buzón no encontrado.")]
    [InlineData("UserDetails.Status.NotFound", "Usuario no encontrado.")]
    [InlineData("AuditLog.Status.FolderNotExist", "Carpeta de auditoría aún no existe.")]
    [InlineData("Audit.Status.NoFindingsToExport", "Sin hallazgos para exportar.")]
    public void EsValue_IsExact(string key, string expected)
    {
        L10n.Get(key).Should().Be(expected);
    }

    [Fact]
    public void ErrorKey_FormatsWithMessage()
    {
        L10n.Format("Common.Status.Error", "boom").Should().Be("Error: boom");
        L10n.Initialize("en");
        L10n.Format("Common.Status.Error", "boom").Should().Be("Error: boom");
    }

    [Fact]
    public void ExportedKey_FormatsWithFileName()
    {
        L10n.Format("Common.Status.Exported", "out.csv").Should().Be("Exportado: out.csv");
    }

    [Theory]
    [InlineData("Offboarding.Status.SuccessSummary", 3, "Offboarding OK · 3 pasos")]
    [InlineData("Onboarding.Status.SuccessSummary", 5, "Onboarding OK · 5 pasos")]
    public void SummaryKey_FormatsWithCount(string key, int count, string expected)
    {
        L10n.Format(key, count).Should().Be(expected);
    }

    [Fact]
    public void ToggleBody_ComposesLocalizedVerb()
    {
        var enable = L10n.Format("UserDetails.Confirm.ToggleBody",
            L10n.Get("UserDetails.Verb.Enable"), "user@contoso.com");
        enable.Should().Be("Habilitar la cuenta user@contoso.com?");

        L10n.Initialize("en");
        var disable = L10n.Format("UserDetails.Confirm.ToggleBody",
            L10n.Get("UserDetails.Verb.Disable"), "user@contoso.com");
        disable.Should().Be("Disable account user@contoso.com?");
    }

    [Fact]
    public void EnValues_DifferFromSpanish_ForTranslatableKeys()
    {
        L10n.Initialize("es");
        var es = L10n.Get("Users.Status.SelectUser");
        L10n.Initialize("en");
        var en = L10n.Get("Users.Status.SelectUser");
        en.Should().Be("Select a user.").And.NotBe(es);
    }
}
