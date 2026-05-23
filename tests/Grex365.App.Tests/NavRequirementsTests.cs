using FluentAssertions;
using Grex365.App.ViewModels;

namespace Grex365.App.Tests;

public class NavRequirementsTests
{
    [Fact]
    public void NoRequirements_AlwaysEnabled()
    {
        NavRequirements.IsEnabled(requiresGraph: false, requiresExchange: false,
            graphConnected: false, exchangeConnected: false).Should().BeTrue();
        NavRequirements.IsEnabled(requiresGraph: false, requiresExchange: false,
            graphConnected: true, exchangeConnected: true).Should().BeTrue();
    }

    [Fact]
    public void RequiresGraph_NeedsGraphConnected()
    {
        NavRequirements.IsEnabled(requiresGraph: true, requiresExchange: false,
            graphConnected: true, exchangeConnected: false).Should().BeTrue();
        NavRequirements.IsEnabled(requiresGraph: true, requiresExchange: false,
            graphConnected: false, exchangeConnected: false).Should().BeFalse();
        NavRequirements.IsEnabled(requiresGraph: true, requiresExchange: false,
            graphConnected: false, exchangeConnected: true).Should().BeFalse();
    }

    [Fact]
    public void RequiresExchange_NeedsExchangeConnected()
    {
        NavRequirements.IsEnabled(requiresGraph: false, requiresExchange: true,
            graphConnected: false, exchangeConnected: true).Should().BeTrue();
        NavRequirements.IsEnabled(requiresGraph: false, requiresExchange: true,
            graphConnected: false, exchangeConnected: false).Should().BeFalse();
    }

    [Fact]
    public void RequiresBoth_NeedsBothConnected()
    {
        NavRequirements.IsEnabled(requiresGraph: true, requiresExchange: true,
            graphConnected: true, exchangeConnected: true).Should().BeTrue();
        NavRequirements.IsEnabled(requiresGraph: true, requiresExchange: true,
            graphConnected: true, exchangeConnected: false).Should().BeFalse();
        NavRequirements.IsEnabled(requiresGraph: true, requiresExchange: true,
            graphConnected: false, exchangeConnected: true).Should().BeFalse();
        NavRequirements.IsEnabled(requiresGraph: true, requiresExchange: true,
            graphConnected: false, exchangeConnected: false).Should().BeFalse();
    }

    [Fact]
    public void Overload_OnItem_HonoursRequirements()
    {
        var dashboard = new NavigationItem("Dashboard", string.Empty, typeof(DashboardViewModel));
        var users = new NavigationItem("Usuarios", string.Empty, typeof(UsersViewModel)) { RequiresGraph = true };
        var mailflow = new NavigationItem("Flujo de correo", string.Empty, typeof(MailFlowRulesViewModel)) { RequiresExchange = true };
        var sharedMbx = new NavigationItem("Buzones", string.Empty, typeof(SharedMailboxViewModel))
        {
            RequiresGraph = true,
            RequiresExchange = true,
        };

        NavRequirements.IsEnabled(dashboard, graphConnected: false, exchangeConnected: false).Should().BeTrue();
        NavRequirements.IsEnabled(users, graphConnected: true, exchangeConnected: false).Should().BeTrue();
        NavRequirements.IsEnabled(users, graphConnected: false, exchangeConnected: true).Should().BeFalse();
        NavRequirements.IsEnabled(mailflow, graphConnected: true, exchangeConnected: false).Should().BeFalse();
        NavRequirements.IsEnabled(mailflow, graphConnected: false, exchangeConnected: true).Should().BeTrue();
        NavRequirements.IsEnabled(sharedMbx, graphConnected: true, exchangeConnected: true).Should().BeTrue();
        NavRequirements.IsEnabled(sharedMbx, graphConnected: true, exchangeConnected: false).Should().BeFalse();
    }

    [Fact]
    public void Overload_NullItem_Throws()
    {
        Action act = () => NavRequirements.IsEnabled(null!, true, true);
        act.Should().Throw<ArgumentNullException>();
    }

    [Theory]
    [InlineData(false, false, true, true, true)]
    [InlineData(true, false, true, false, true)]
    [InlineData(true, false, false, false, false)]
    [InlineData(false, true, false, true, true)]
    [InlineData(false, true, true, false, false)]
    [InlineData(true, true, true, true, true)]
    [InlineData(true, true, false, true, false)]
    [InlineData(true, true, true, false, false)]
    [InlineData(true, true, false, false, false)]
    public void TruthTable(bool reqG, bool reqE, bool g, bool e, bool expected)
    {
        NavRequirements.IsEnabled(reqG, reqE, g, e).Should().Be(expected);
    }
}
