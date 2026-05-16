using FluentAssertions;
using Grex365.Core.Mailboxes;
using Grex365.Core.Models;

namespace Grex365.Core.Tests;

public class MailboxRulesValidatorTests
{
    [Fact]
    public void Disabled_State_PassesEvenWithoutMessages()
    {
        var cfg = new AutoReplyConfig(AutoReplyState.Disabled, null, null, null, null);
        MailboxRulesValidator.ValidateAutoReply(cfg).Should().BeEmpty();
    }

    [Fact]
    public void Enabled_Without_AnyMessage_Fails()
    {
        var cfg = new AutoReplyConfig(AutoReplyState.Enabled, "  ", null, null, null);
        var errs = MailboxRulesValidator.ValidateAutoReply(cfg);
        errs.Should().Contain(e => e.Contains("mensaje", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Enabled_With_OnlyInternal_Passes()
    {
        var cfg = new AutoReplyConfig(AutoReplyState.Enabled, "Estoy fuera", null, null, null);
        MailboxRulesValidator.ValidateAutoReply(cfg).Should().BeEmpty();
    }

    [Fact]
    public void Scheduled_Without_StartOrEnd_Fails()
    {
        var cfg = new AutoReplyConfig(AutoReplyState.Scheduled, "x", "y", null, null);
        var errs = MailboxRulesValidator.ValidateAutoReply(cfg);
        errs.Should().Contain(e => e.Contains("StartTime"));
        errs.Should().Contain(e => e.Contains("EndTime"));
    }

    [Fact]
    public void Scheduled_End_BeforeStart_Fails()
    {
        var start = new DateTime(2026, 5, 20);
        var end = new DateTime(2026, 5, 10);
        var cfg = new AutoReplyConfig(AutoReplyState.Scheduled, "x", "y", start, end);
        MailboxRulesValidator.ValidateAutoReply(cfg)
            .Should().Contain(e => e.Contains("posterior", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Scheduled_Equal_Dates_Fails()
    {
        var d = new DateTime(2026, 5, 20);
        var cfg = new AutoReplyConfig(AutoReplyState.Scheduled, "x", "y", d, d);
        MailboxRulesValidator.ValidateAutoReply(cfg).Should().NotBeEmpty();
    }

    [Fact]
    public void Scheduled_Valid_Range_Passes()
    {
        var start = new DateTime(2026, 5, 10);
        var end = new DateTime(2026, 5, 20);
        var cfg = new AutoReplyConfig(AutoReplyState.Scheduled, "x", "y", start, end);
        MailboxRulesValidator.ValidateAutoReply(cfg).Should().BeEmpty();
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    public void Forwarding_Empty_Fails(string? input)
    {
        MailboxRulesValidator.ValidateForwarding(input).Should().NotBeEmpty();
    }

    [Theory]
    [InlineData("notanemail")]
    [InlineData("a@b")]
    [InlineData("@b.com")]
    public void Forwarding_Bad_Shape_Fails(string input)
    {
        MailboxRulesValidator.ValidateForwarding(input).Should().NotBeEmpty();
    }

    [Theory]
    [InlineData("user@example.com")]
    [InlineData("a.b.c@sub.domain.io")]
    public void Forwarding_Valid_Passes(string input)
    {
        MailboxRulesValidator.ValidateForwarding(input).Should().BeEmpty();
    }
}
