using FluentAssertions;
using Grex365.Core.Offboarding;

namespace Grex365.Core.Tests;

public class OffboardingAutoReplyTests
{
    [Fact]
    public void Render_SubstitutesUserAndDelegate()
    {
        var msg = OffboardingAutoReply.Render(
            "{usuario} ya no forma parte de la organización. Contacte con {delegado}.",
            "Jane Doe", "deleg@a");

        msg.Should().Be("Jane Doe ya no forma parte de la organización. Contacte con deleg@a.");
    }

    [Fact]
    public void Render_IsCaseInsensitive()
    {
        OffboardingAutoReply.Render("Contacte con {Delegado}", null, "x@a")
            .Should().Be("Contacte con x@a");
    }

    [Fact]
    public void Render_MissingValues_CollapseToEmpty()
    {
        OffboardingAutoReply.Render("[{usuario}] contacte {delegado}", null, null)
            .Should().Be("[] contacte ");
    }

    [Fact]
    public void Render_NoTokens_ReturnsAsIs()
    {
        OffboardingAutoReply.Render("Mensaje libre sin tokens", "Jane", "d@a")
            .Should().Be("Mensaje libre sin tokens");
    }

    [Fact]
    public void Render_EmptyTemplate_ReturnsEmpty()
    {
        OffboardingAutoReply.Render("", "Jane", "d@a").Should().BeEmpty();
        OffboardingAutoReply.Render(null, "Jane", "d@a").Should().BeEmpty();
    }

    private const string Global = "Contacte con {delegado}.";
    private const string NoDel = "{usuario} ya no trabaja aquí.";

    [Fact]
    public void Resolve_PerUserOverride_AlwaysWins()
    {
        OffboardingAutoReply.Resolve(Global, "Mensaje propio para {delegado}", NoDel, "Jane", "p@a")
            .Should().Be("Mensaje propio para p@a");
    }

    [Fact]
    public void Resolve_WithDelegate_UsesGlobalTemplate()
    {
        OffboardingAutoReply.Resolve(Global, null, NoDel, "Jane", "pepe@a")
            .Should().Be("Contacte con pepe@a.");
    }

    [Fact]
    public void Resolve_NoDelegate_UsesNoDelegateTemplate()
    {
        OffboardingAutoReply.Resolve(Global, null, NoDel, "Jane", "")
            .Should().Be("Jane ya no trabaja aquí.");
    }

    [Fact]
    public void Resolve_AutoReplyOff_NoOverride_ReturnsNull()
    {
        OffboardingAutoReply.Resolve("", null, NoDel, "Jane", "pepe@a").Should().BeNull();
    }

    // Grouping: one shared global template, different delegates per user → different messages.
    [Theory]
    [InlineData("pepe@a", "Contacte con pepe@a.")]
    [InlineData("marta@a", "Contacte con marta@a.")]
    public void Resolve_SharedTemplate_PerUserDelegate(string del, string expected)
    {
        OffboardingAutoReply.Resolve(Global, null, NoDel, "Jane", del).Should().Be(expected);
    }
}
