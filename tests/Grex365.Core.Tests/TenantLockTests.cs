using FluentAssertions;
using Grex365.Core.Abstractions;
using Grex365.Core.Connections;
using Grex365.Core.Models;
using Moq;

namespace Grex365.Core.Tests;

public class TenantLockTests
{
    private static Mock<IPreferencesStore> Store(UserPreferences prefs)
    {
        var mock = new Mock<IPreferencesStore>();
        mock.Setup(s => s.LoadAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(prefs);
        return mock;
    }

    [Fact]
    public async Task NoEnforce_DoesNothing()
    {
        var prefs = new UserPreferences { EnforceTenantLock = false, ExpectedTenantId = "ABC" };
        var sut = new TenantLock(Store(prefs).Object);

        await sut.EnforceAsync("DIFFERENT");
    }

    [Fact]
    public async Task Enforce_Match_DoesNothing()
    {
        var prefs = new UserPreferences { EnforceTenantLock = true, ExpectedTenantId = "abc-123" };
        var sut = new TenantLock(Store(prefs).Object);

        await sut.EnforceAsync("ABC-123"); // case-insensitive
    }

    [Fact]
    public async Task Enforce_Mismatch_Throws()
    {
        var prefs = new UserPreferences { EnforceTenantLock = true, ExpectedTenantId = "expected" };
        var sut = new TenantLock(Store(prefs).Object);

        Func<Task> act = () => sut.EnforceAsync("other");

        var ex = await act.Should().ThrowAsync<TenantLockViolationException>();
        ex.Which.Expected.Should().Be("expected");
        ex.Which.Actual.Should().Be("other");
    }

    [Fact]
    public async Task NullActualTenant_DoesNothing()
    {
        var prefs = new UserPreferences { EnforceTenantLock = true, ExpectedTenantId = "abc" };
        var sut = new TenantLock(Store(prefs).Object);

        await sut.EnforceAsync(string.Empty);
    }

    [Fact]
    public async Task EnforceButNoExpectedConfigured_DoesNothing()
    {
        var prefs = new UserPreferences { EnforceTenantLock = true, ExpectedTenantId = null };
        var sut = new TenantLock(Store(prefs).Object);

        await sut.EnforceAsync("anything");
    }
}
