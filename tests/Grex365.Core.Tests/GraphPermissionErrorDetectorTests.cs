using FluentAssertions;
using Grex365.Core.Audit;

namespace Grex365.Core.Tests;

public class GraphPermissionErrorDetectorTests
{
    [Fact]
    public void IsAuditLogPermissionError_NullException_ReturnsFalse()
    {
        GraphPermissionErrorDetector.IsAuditLogPermissionError(null).Should().BeFalse();
    }

    [Fact]
    public void IsAuditLogPermissionError_EmptyMessage_ReturnsFalse()
    {
        var ex = new InvalidOperationException(string.Empty);
        GraphPermissionErrorDetector.IsAuditLogPermissionError(ex).Should().BeFalse();
    }

    [Fact]
    public void IsAuditLogPermissionError_ContainsAuditLogScope_ReturnsTrue()
    {
        var ex = new InvalidOperationException("Caller needs scope AuditLog.Read.All to read signInActivity.");
        GraphPermissionErrorDetector.IsAuditLogPermissionError(ex).Should().BeTrue();
    }

    [Fact]
    public void IsAuditLogPermissionError_ContainsGenericPhrase_ReturnsTrue()
    {
        var ex = new InvalidOperationException("The required Microsoft Graph permission is missing.");
        GraphPermissionErrorDetector.IsAuditLogPermissionError(ex).Should().BeTrue();
    }

    [Theory]
    [InlineData("auditlog.read.all")]
    [InlineData("AUDITLOG.READ.ALL")]
    [InlineData("AuditLog.READ.All")]
    [InlineData("required microsoft graph permission")]
    [InlineData("REQUIRED MICROSOFT GRAPH PERMISSION")]
    public void IsAuditLogPermissionError_CaseInsensitive(string message)
    {
        var ex = new InvalidOperationException(message);
        GraphPermissionErrorDetector.IsAuditLogPermissionError(ex).Should().BeTrue();
    }

    [Fact]
    public void IsAuditLogPermissionError_UnrelatedMessage_ReturnsFalse()
    {
        var ex = new InvalidOperationException("Tenant not found.");
        GraphPermissionErrorDetector.IsAuditLogPermissionError(ex).Should().BeFalse();
    }

    [Fact]
    public void IsAnyPermissionError_NullException_ReturnsFalse()
    {
        GraphPermissionErrorDetector.IsAnyPermissionError(null).Should().BeFalse();
    }

    [Theory]
    [InlineData("The required Microsoft Graph permission is missing.")]
    [InlineData("Insufficient privileges to complete the operation.")]
    [InlineData("Authorization_RequestDenied")]
    [InlineData("Forbidden: app does not have role")]
    [InlineData("INSUFFICIENT PRIVILEGES")]
    [InlineData("forbidden")]
    public void IsAnyPermissionError_RecognizedPhrases_ReturnsTrue(string message)
    {
        var ex = new InvalidOperationException(message);
        GraphPermissionErrorDetector.IsAnyPermissionError(ex).Should().BeTrue();
    }

    [Theory]
    [InlineData("Tenant lock mismatch")]
    [InlineData("Request timed out")]
    [InlineData("Network error")]
    [InlineData("")]
    public void IsAnyPermissionError_UnrelatedMessages_ReturnsFalse(string message)
    {
        var ex = new InvalidOperationException(message);
        GraphPermissionErrorDetector.IsAnyPermissionError(ex).Should().BeFalse();
    }
}
