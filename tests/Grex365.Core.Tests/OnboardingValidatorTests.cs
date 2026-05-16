using FluentAssertions;
using Grex365.Core.Abstractions;
using Grex365.Core.Onboarding;

namespace Grex365.Core.Tests;

public class OnboardingValidatorTests
{
    private static OnboardingOptions Valid(
        string? displayName = "Ana Garcia",
        string? upn = "ana.garcia@contoso.com",
        string? password = "Sup3rSecret!",
        string? usage = "ES")
        => new(
            DisplayName: displayName!,
            Upn: upn!,
            InitialPassword: password!,
            UsageLocation: usage!,
            MailNickname: null,
            SkuIds: Array.Empty<Guid>(),
            GroupIdentifiers: Array.Empty<string>());

    [Fact]
    public void Valid_Options_NoErrors()
    {
        OnboardingValidator.Validate(Valid()).Should().BeEmpty();
    }

    [Theory]
    [InlineData("", "DisplayName")]
    [InlineData("   ", "DisplayName")]
    public void Empty_DisplayName_Rejected(string name, string token)
    {
        var errs = OnboardingValidator.Validate(Valid(displayName: name));
        errs.Should().Contain(e => e.Contains(token));
    }

    [Theory]
    [InlineData("")]
    [InlineData("notanemail")]
    [InlineData("user@")]
    [InlineData("@domain.com")]
    [InlineData("user@domain")]
    public void Invalid_Upn_Rejected(string upn)
    {
        var errs = OnboardingValidator.Validate(Valid(upn: upn));
        errs.Should().Contain(e => e.Contains("UPN", StringComparison.OrdinalIgnoreCase));
    }

    [Theory]
    [InlineData("")]
    [InlineData("1234567")]
    public void Short_Password_Rejected(string pw)
    {
        var errs = OnboardingValidator.Validate(Valid(password: pw));
        errs.Should().Contain(e => e.Contains("Password", StringComparison.OrdinalIgnoreCase));
    }

    [Theory]
    [InlineData("")]
    [InlineData("E")]
    [InlineData("ESP")]
    public void Bad_UsageLocation_Rejected(string usage)
    {
        var errs = OnboardingValidator.Validate(Valid(usage: usage));
        errs.Should().Contain(e => e.Contains("UsageLocation"));
    }

    [Fact]
    public void DeriveMailNickname_FromUpn_TakesLocalPart()
    {
        OnboardingValidator.DeriveMailNickname("ana.garcia@contoso.com", null)
            .Should().Be("ana.garcia");
    }

    [Fact]
    public void DeriveMailNickname_Explicit_Wins()
    {
        OnboardingValidator.DeriveMailNickname("ana.garcia@contoso.com", " custom-alias ")
            .Should().Be("custom-alias");
    }

    [Fact]
    public void DeriveMailNickname_Empty_When_Both_Empty()
    {
        OnboardingValidator.DeriveMailNickname("", null).Should().BeEmpty();
    }
}
