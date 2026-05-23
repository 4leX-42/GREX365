using System.ComponentModel;
using FluentAssertions;
using Grex365.Core.Users;

namespace Grex365.Core.Tests;

public class UserDetailsHostTests
{
    [Fact]
    public void Initial_State_IsClosedAndNull()
    {
        var host = new UserDetailsHost();

        host.IsOpen.Should().BeFalse();
        host.CurrentUserId.Should().BeNull();
    }

    [Fact]
    public void RequestOpen_ValidId_SetsStateAndFiresOpenRequested()
    {
        var host = new UserDetailsHost();
        string? opened = null;
        host.OpenRequested += (_, uid) => opened = uid;

        host.RequestOpen("uid-1");

        host.IsOpen.Should().BeTrue();
        host.CurrentUserId.Should().Be("uid-1");
        opened.Should().Be("uid-1");
    }

    [Fact]
    public void RequestOpen_RaisesPropertyChanged_ForIsOpen_AndCurrentUserId()
    {
        var host = new UserDetailsHost();
        var props = new List<string>();
        host.PropertyChanged += (_, e) => props.Add(e.PropertyName ?? string.Empty);

        host.RequestOpen("uid-9");

        props.Should().Contain(nameof(UserDetailsHost.IsOpen));
        props.Should().Contain(nameof(UserDetailsHost.CurrentUserId));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void RequestOpen_WhitespaceOrNull_Noop(string? uid)
    {
        var host = new UserDetailsHost();
        var openedFired = false;
        host.OpenRequested += (_, _) => openedFired = true;

        host.RequestOpen(uid!);

        host.IsOpen.Should().BeFalse();
        host.CurrentUserId.Should().BeNull();
        openedFired.Should().BeFalse();
    }

    [Fact]
    public void RequestOpen_SameId_DoesNotFireDuplicatePropertyChanged()
    {
        var host = new UserDetailsHost();
        host.RequestOpen("uid-1");

        var props = new List<string>();
        host.PropertyChanged += (_, e) => props.Add(e.PropertyName ?? string.Empty);
        host.RequestOpen("uid-1");

        // OpenRequested fires every time (it's an action), but PropertyChanged
        // is gated by value-changed check → CurrentUserId and IsOpen both
        // already match, so no PropertyChanged should be raised.
        props.Should().BeEmpty();
    }

    [Fact]
    public void RequestClose_FromOpen_ResetsAndFiresCloseRequested()
    {
        var host = new UserDetailsHost();
        host.RequestOpen("uid-1");
        var closedFired = false;
        host.CloseRequested += (_, _) => closedFired = true;

        host.RequestClose();

        host.IsOpen.Should().BeFalse();
        host.CurrentUserId.Should().BeNull();
        closedFired.Should().BeTrue();
    }

    [Fact]
    public void RequestClose_FromClosed_NoPropertyChanged_ButCloseRequestedFires()
    {
        var host = new UserDetailsHost();
        var props = new List<string>();
        var closedFired = false;
        host.PropertyChanged += (_, e) => props.Add(e.PropertyName ?? string.Empty);
        host.CloseRequested += (_, _) => closedFired = true;

        host.RequestClose();

        host.IsOpen.Should().BeFalse();
        host.CurrentUserId.Should().BeNull();
        props.Should().BeEmpty(because: "fields already at default; no value-changed transitions");
        closedFired.Should().BeTrue(because: "RequestClose always fires, value-gated PropertyChanged is separate");
    }

    [Fact]
    public void Open_Then_Close_Then_Open_WorksCorrectly()
    {
        var host = new UserDetailsHost();

        host.RequestOpen("uid-1");
        host.RequestClose();
        host.RequestOpen("uid-2");

        host.IsOpen.Should().BeTrue();
        host.CurrentUserId.Should().Be("uid-2");
    }
}
