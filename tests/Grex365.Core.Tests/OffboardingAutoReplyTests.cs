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
}
