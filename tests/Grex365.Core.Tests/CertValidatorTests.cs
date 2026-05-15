using FluentAssertions;
using Grex365.Core.Abstractions;
using Grex365.Core.Connections;
using Grex365.Core.Models;

namespace Grex365.Core.Tests;

public class CertValidatorTests
{
    [Fact]
    public void Validate_Null_Returns_MissingConfig()
    {
        var sut = new CertValidator();
        sut.Validate(null).Status.Should().Be(CertValidationStatus.MissingConfig);
    }

    [Fact]
    public void Validate_EmptyThumbprint_Returns_MissingConfig()
    {
        var sut = new CertValidator();
        var cfg = new CertConfig(AppId: "x", TenantId: "y", Organization: "z", CertThumbprint: "");
        sut.Validate(cfg).Status.Should().Be(CertValidationStatus.MissingConfig);
    }

    [Fact]
    public void Validate_NonExistentThumbprint_Returns_MissingFromStore()
    {
        var sut = new CertValidator();
        var cfg = new CertConfig(
            AppId: "x",
            TenantId: "y",
            Organization: "z",
            CertThumbprint: "0000000000000000000000000000000000000000");
        sut.Validate(cfg).Status.Should().Be(CertValidationStatus.MissingFromStore);
    }
}
