using FluentAssertions;
using Grex365.App;
using Grex365.App.ViewModels;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Moq;

namespace Grex365.App.Tests;

[Collection("L10n")]
public class FirstRunWizardViewModelTests : IDisposable
{
    public FirstRunWizardViewModelTests() => L10n.Initialize("es");
    public void Dispose() => L10n.Reset();

    private static Mock<IPreferencesStore> MakePrefsMock(UserPreferences? initial = null)
    {
        var prefs = initial ?? new UserPreferences();
        var mock = new Mock<IPreferencesStore>();
        mock.Setup(p => p.LoadAsync(It.IsAny<CancellationToken>())).ReturnsAsync(prefs);
        mock.Setup(p => p.SaveAsync(It.IsAny<UserPreferences>(), It.IsAny<CancellationToken>())).Returns(Task.CompletedTask);
        return mock;
    }

    [Fact]
    public void InitialState_IsWelcome()
    {
        var vm = new FirstRunWizardViewModel(MakePrefsMock().Object);
        vm.CurrentStep.Should().Be(FirstRunStep.Welcome);
        vm.IsWelcome.Should().BeTrue();
        vm.CanGoBack.Should().BeFalse();
        vm.CanGoNext.Should().BeTrue();
    }

    [Fact]
    public void Next_AdvancesThroughAllSteps()
    {
        var vm = new FirstRunWizardViewModel(MakePrefsMock().Object);

        vm.NextCommand.Execute(null);
        vm.CurrentStep.Should().Be(FirstRunStep.Connection);

        vm.NextCommand.Execute(null);
        vm.CurrentStep.Should().Be(FirstRunStep.TenantLock);

        vm.NextCommand.Execute(null);
        vm.CurrentStep.Should().Be(FirstRunStep.Theme);

        vm.NextCommand.Execute(null);
        vm.CurrentStep.Should().Be(FirstRunStep.Summary);

        vm.IsSummary.Should().BeTrue();
        vm.CanGoNext.Should().BeFalse();
    }

    [Fact]
    public void Next_OnSummary_IsNoOp()
    {
        var vm = new FirstRunWizardViewModel(MakePrefsMock().Object);
        vm.CurrentStep = FirstRunStep.Summary;

        vm.NextCommand.Execute(null);

        vm.CurrentStep.Should().Be(FirstRunStep.Summary);
    }

    [Fact]
    public void Back_OnWelcome_IsNoOp()
    {
        var vm = new FirstRunWizardViewModel(MakePrefsMock().Object);

        vm.BackCommand.Execute(null);

        vm.CurrentStep.Should().Be(FirstRunStep.Welcome);
    }

    [Fact]
    public void Back_FromConnection_GoesToWelcome()
    {
        var vm = new FirstRunWizardViewModel(MakePrefsMock().Object);
        vm.CurrentStep = FirstRunStep.Connection;

        vm.BackCommand.Execute(null);

        vm.CurrentStep.Should().Be(FirstRunStep.Welcome);
        vm.IsWelcome.Should().BeTrue();
    }

    [Fact]
    public async Task Skip_SetsFirstRunCompleted_AndMarksSkippedCompleted()
    {
        var saved = new UserPreferences();
        var mock = new Mock<IPreferencesStore>();
        mock.Setup(p => p.LoadAsync(It.IsAny<CancellationToken>())).ReturnsAsync(saved);
        mock.Setup(p => p.SaveAsync(It.IsAny<UserPreferences>(), It.IsAny<CancellationToken>()))
            .Callback<UserPreferences, CancellationToken>((u, _) => saved = u)
            .Returns(Task.CompletedTask);
        var vm = new FirstRunWizardViewModel(mock.Object);

        await vm.SkipCommand.ExecuteAsync(null);

        vm.Skipped.Should().BeTrue();
        vm.Completed.Should().BeTrue();
        saved.FirstRunCompleted.Should().BeTrue();
        mock.Verify(p => p.SaveAsync(It.Is<UserPreferences>(u => u.FirstRunCompleted),
            It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task Finish_PersistsAllPreferences()
    {
        var saved = new UserPreferences();
        var mock = new Mock<IPreferencesStore>();
        mock.Setup(p => p.LoadAsync(It.IsAny<CancellationToken>())).ReturnsAsync(saved);
        mock.Setup(p => p.SaveAsync(It.IsAny<UserPreferences>(), It.IsAny<CancellationToken>()))
            .Callback<UserPreferences, CancellationToken>((u, _) => saved = u)
            .Returns(Task.CompletedTask);
        var vm = new FirstRunWizardViewModel(mock.Object)
        {
            ConnectionMethod = "cert",
            EnforceTenantLock = true,
            ExpectedTenantId = "  abc-tenant-id  ",
            ExpectedTenantDomain = "  contoso.onmicrosoft.com  ",
            Theme = "Light"
        };

        await vm.FinishCommand.ExecuteAsync(null);

        vm.Completed.Should().BeTrue();
        vm.Skipped.Should().BeFalse();
        saved.FirstRunCompleted.Should().BeTrue();
        saved.ConnectionMethod.Should().Be("cert");
        saved.EnforceTenantLock.Should().BeTrue();
        saved.ExpectedTenantId.Should().Be("abc-tenant-id");
        saved.ExpectedTenantDomain.Should().Be("contoso.onmicrosoft.com");
        saved.Theme.Should().Be("Light");
    }

    [Fact]
    public async Task Finish_EmptyTenantValues_PersistsAsNull()
    {
        var saved = new UserPreferences();
        var mock = new Mock<IPreferencesStore>();
        mock.Setup(p => p.LoadAsync(It.IsAny<CancellationToken>())).ReturnsAsync(saved);
        mock.Setup(p => p.SaveAsync(It.IsAny<UserPreferences>(), It.IsAny<CancellationToken>()))
            .Callback<UserPreferences, CancellationToken>((u, _) => saved = u)
            .Returns(Task.CompletedTask);
        var vm = new FirstRunWizardViewModel(mock.Object)
        {
            ExpectedTenantId = "   ",
            ExpectedTenantDomain = string.Empty
        };

        await vm.FinishCommand.ExecuteAsync(null);

        saved.ExpectedTenantId.Should().BeNull();
        saved.ExpectedTenantDomain.Should().BeNull();
    }

    [Fact]
    public async Task Finish_SaveThrows_SetsErrorStatus_AndDoesNotMarkCompleted()
    {
        var mock = new Mock<IPreferencesStore>();
        mock.Setup(p => p.LoadAsync(It.IsAny<CancellationToken>())).ReturnsAsync(new UserPreferences());
        mock.Setup(p => p.SaveAsync(It.IsAny<UserPreferences>(), It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("disk full"));
        var vm = new FirstRunWizardViewModel(mock.Object);

        await vm.FinishCommand.ExecuteAsync(null);

        vm.Completed.Should().BeFalse();
        vm.SaveStatus.Should().StartWith("Error:");
    }

    [Fact]
    public void Labels_ReflectCurrentValues()
    {
        var vm = new FirstRunWizardViewModel(MakePrefsMock().Object);

        vm.ConnectionMethod = "cert";
        vm.ConnectionMethodLabel.Should().Contain("Certificado");

        vm.ConnectionMethod = "devicecode";
        vm.ConnectionMethodLabel.Should().Contain("Device code");

        vm.EnforceTenantLock = false;
        vm.TenantLockLabel.Should().Contain("Desactivado");

        vm.EnforceTenantLock = true;
        vm.ExpectedTenantId = "tenant-xyz";
        vm.TenantLockLabel.Should().Contain("Activado").And.Contain("tenant-xyz");
    }

    [Fact]
    public void Labels_English_TranslatedViaL10n()
    {
        L10n.Initialize("en");
        var vm = new FirstRunWizardViewModel(MakePrefsMock().Object);

        vm.ConnectionMethod = "cert";
        vm.ConnectionMethodLabel.Should().Contain("Certificate");

        vm.ConnectionMethod = "devicecode";
        vm.ConnectionMethodLabel.Should().Contain("Device code");

        vm.EnforceTenantLock = false;
        vm.TenantLockLabel.Should().Contain("Disabled");

        vm.EnforceTenantLock = true;
        vm.ExpectedTenantId = "tenant-en";
        vm.TenantLockLabel.Should().Contain("Enabled").And.Contain("tenant-en");
    }

    [Fact]
    public void TenantLockLabel_EmptyValues_UsesEmptyMarker()
    {
        var vm = new FirstRunWizardViewModel(MakePrefsMock().Object)
        {
            EnforceTenantLock = true,
            ExpectedTenantId = string.Empty,
            ExpectedTenantDomain = "   "
        };

        vm.TenantLockLabel.Should().Contain("—");
    }

    [Fact]
    public async Task Skip_SetsLocalizedSavingStatus_DuringExecution()
    {
        var prefsMock = new Mock<IPreferencesStore>();
        var loadTcs = new TaskCompletionSource<UserPreferences>();
        prefsMock.Setup(p => p.LoadAsync(It.IsAny<CancellationToken>())).Returns(loadTcs.Task);
        prefsMock.Setup(p => p.SaveAsync(It.IsAny<UserPreferences>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);
        var vm = new FirstRunWizardViewModel(prefsMock.Object);

        var task = vm.SkipCommand.ExecuteAsync(null);
        vm.SaveStatus.Should().Contain("Saltando");
        loadTcs.SetResult(new UserPreferences());
        await task;
    }

    [Fact]
    public async Task Finish_SetsLocalizedSavedStatus_OnSuccess()
    {
        var mock = MakePrefsMock();
        var vm = new FirstRunWizardViewModel(mock.Object);

        await vm.FinishCommand.ExecuteAsync(null);

        vm.SaveStatus.Should().Be("Guardado.");
    }

    [Fact]
    public async Task Finish_English_LocalizedSavedStatus()
    {
        L10n.Initialize("en");
        var mock = MakePrefsMock();
        var vm = new FirstRunWizardViewModel(mock.Object);

        await vm.FinishCommand.ExecuteAsync(null);

        vm.SaveStatus.Should().Be("Saved.");
    }
}
