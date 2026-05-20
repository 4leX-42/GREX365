using FluentAssertions;
using Grex365.Core.Certificates;

namespace Grex365.Core.Tests;

public class CertificateGeneratorTests
{
    [Fact]
    public void ExportPfx_Throws_When_Thumbprint_Empty()
    {
        var sut = new SelfSignedCertificateGenerator();
        var act = () => sut.ExportPfx(string.Empty, "out.pfx", "p");
        act.Should().Throw<ArgumentException>().WithParameterName("thumbprint");
    }

    [Fact]
    public void ExportPfx_Throws_When_OutputPath_Empty()
    {
        var sut = new SelfSignedCertificateGenerator();
        var act = () => sut.ExportPfx("ABCDEF", string.Empty, "p");
        act.Should().Throw<ArgumentException>().WithParameterName("outputPath");
    }

    [Fact]
    public void ExportPfx_Throws_When_Password_Empty()
    {
        var sut = new SelfSignedCertificateGenerator();
        var act = () => sut.ExportPfx("ABCDEF", "out.pfx", string.Empty);
        act.Should().Throw<ArgumentException>().WithParameterName("password");
    }

    [Fact]
    public void ExportPfx_Throws_When_Thumbprint_NotFound_In_Store()
    {
        var sut = new SelfSignedCertificateGenerator();
        var bogus = new string('A', 40); // 40-char hex-like but unlikely to match real cert
        var act = () => sut.ExportPfx(bogus, Path.Combine(Path.GetTempPath(), "should-not-exist.pfx"), "p");
        act.Should().Throw<InvalidOperationException>()
            .Which.Message.Should().Contain("no encontrado");
    }
}
