using System.Diagnostics;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.SkillGeneration.Tests;

[Collection("Sequential")]
[Trait("Category", "Integration")]
[Trait("Feature", "Distribution")]
[Trait("RequiresExcel", "false")]
public sealed class MacVbaHelperPackageTests
{
    private static readonly string RepoRoot = FindRepoRoot();
    private static readonly string ScriptPath = Path.Combine(
        RepoRoot,
        "scripts",
        "Build-MacVbaHelperPackage.ps1");

    [Fact]
    public async Task BuildPackage_RecordsOpaqueArtifactAndSelfCertProvenance()
    {
        var sandbox = Path.Combine(Path.GetTempPath(), $"ExcelMcpVbaHelper-{Guid.NewGuid():N}");
        var output = Path.Combine(sandbox, "package");
        Directory.CreateDirectory(sandbox);

        try
        {
            var helperPath = Path.Combine(sandbox, "ExcelMcpHelper.xlam");
            var sourcePath = Path.Combine(sandbox, "ExcelMcpHelper.bas");
            var certificatePath = Path.Combine(sandbox, "ExcelMcpHelper.cer");
            await File.WriteAllBytesAsync(helperPath, [0x50, 0x4B, 0x03, 0x04]);
            await File.WriteAllTextAsync(
                sourcePath,
                """
                Private Const HELPER_VERSION As String = "1.4.0"
                Private Const PROTOCOL_VERSION As Long = 1
                """);

            using var rsa = RSA.Create(2048);
            var request = new CertificateRequest(
                "CN=ExcelMcp SelfCert Test",
                rsa,
                HashAlgorithmName.SHA256,
                RSASignaturePadding.Pkcs1);
            var usages = new OidCollection
            {
                new("1.3.6.1.5.5.7.3.3")
            };
            request.CertificateExtensions.Add(new X509EnhancedKeyUsageExtension(usages, false));
            using var certificate = request.CreateSelfSigned(
                DateTimeOffset.UtcNow.AddDays(-1),
                DateTimeOffset.UtcNow.AddYears(1));
            await File.WriteAllBytesAsync(
                certificatePath,
                certificate.Export(X509ContentType.Cert));

            var result = await RunScriptAsync(
                "-HelperPath", helperPath,
                "-SourcePath", sourcePath,
                "-PublicCertificatePath", certificatePath,
                "-OutputDirectory", output,
                "-SourceCommit", "deadbeef",
                "-ExcelSignatureVerifiedConfirmed",
                "-SelfCertTrustModelConfirmed");

            Assert.True(result.ExitCode == 0, result.CombinedOutput);
            Assert.True(File.Exists(Path.Combine(output, "ExcelMcpHelper.xlam")));
            Assert.True(File.Exists(Path.Combine(output, "ExcelMcpHelper.bas")));
            Assert.True(File.Exists(Path.Combine(output, "ExcelMcpHelper.cer")));

            using var manifest = JsonDocument.Parse(await File.ReadAllTextAsync(
                Path.Combine(output, "ExcelMcpHelper.manifest.json")));
            var root = manifest.RootElement;
            Assert.Equal("1.4.0", root.GetProperty("helperVersion").GetString());
            Assert.Equal(1, root.GetProperty("protocolVersion").GetInt32());
            Assert.Equal("deadbeef", root.GetProperty("sourceCommit").GetString());
            Assert.Equal(
                "self-signed-manual-per-machine",
                root.GetProperty("trustModel").GetString());
            Assert.True(root.GetProperty("excelSignatureVerified").GetBoolean());
            Assert.Equal(
                certificate.Thumbprint,
                root.GetProperty("certificateThumbprint").GetString(),
                ignoreCase: true);
            Assert.Matches(
                "^[A-F0-9]{64}$",
                root.GetProperty("artifactSha256").GetString());
            Assert.Matches(
                "^[A-F0-9]{64}$",
                root.GetProperty("sourceSha256").GetString());
            using var packagedCertificate = X509CertificateLoader.LoadCertificateFromFile(
                Path.Combine(output, "ExcelMcpHelper.cer"));
            Assert.False(packagedCertificate.HasPrivateKey);
        }
        finally
        {
            if (Directory.Exists(sandbox))
            {
                Directory.Delete(sandbox, recursive: true);
            }
        }
    }

    [Fact]
    public async Task BuildPackage_RejectsCertificateContainingPrivateKey()
    {
        var sandbox = Path.Combine(Path.GetTempPath(), $"ExcelMcpVbaHelper-{Guid.NewGuid():N}");
        Directory.CreateDirectory(sandbox);

        try
        {
            var helperPath = Path.Combine(sandbox, "ExcelMcpHelper.xlam");
            var sourcePath = Path.Combine(sandbox, "ExcelMcpHelper.bas");
            var certificatePath = Path.Combine(sandbox, "ExcelMcpHelper.pfx");
            await File.WriteAllBytesAsync(helperPath, [0x50, 0x4B, 0x03, 0x04]);
            await File.WriteAllTextAsync(
                sourcePath,
                """
                Private Const HELPER_VERSION As String = "1.4.0"
                Private Const PROTOCOL_VERSION As Long = 1
                """);

            using var rsa = RSA.Create(2048);
            var request = new CertificateRequest(
                "CN=ExcelMcp Private Key Test",
                rsa,
                HashAlgorithmName.SHA256,
                RSASignaturePadding.Pkcs1);
            var usages = new OidCollection
            {
                new("1.3.6.1.5.5.7.3.3")
            };
            request.CertificateExtensions.Add(new X509EnhancedKeyUsageExtension(usages, false));
            using var certificate = request.CreateSelfSigned(
                DateTimeOffset.UtcNow.AddDays(-1),
                DateTimeOffset.UtcNow.AddYears(1));
            await File.WriteAllBytesAsync(
                certificatePath,
                certificate.Export(X509ContentType.Pfx, string.Empty));

            var result = await RunScriptAsync(
                "-HelperPath", helperPath,
                "-SourcePath", sourcePath,
                "-PublicCertificatePath", certificatePath,
                "-OutputDirectory", Path.Combine(sandbox, "package"),
                "-SourceCommit", "deadbeef",
                "-ExcelSignatureVerifiedConfirmed",
                "-SelfCertTrustModelConfirmed");

            Assert.NotEqual(0, result.ExitCode);
            Assert.False(Directory.Exists(Path.Combine(sandbox, "package")));
        }
        finally
        {
            if (Directory.Exists(sandbox))
            {
                Directory.Delete(sandbox, recursive: true);
            }
        }
    }

    private static async Task<ScriptResult> RunScriptAsync(params string[] arguments)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = "pwsh",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            WorkingDirectory = RepoRoot
        };
        startInfo.ArgumentList.Add("-NoProfile");
        startInfo.ArgumentList.Add("-File");
        startInfo.ArgumentList.Add(ScriptPath);
        foreach (var argument in arguments)
        {
            startInfo.ArgumentList.Add(argument);
        }

        using var process = Process.Start(startInfo);
        Assert.NotNull(process);
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
        await process.WaitForExitAsync(timeout.Token);
        return new ScriptResult(process.ExitCode, await stdout, await stderr);
    }

    private static string FindRepoRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Sbroenne.ExcelMcp.sln")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException("Could not locate repository root.");
    }

    private sealed record ScriptResult(int ExitCode, string Stdout, string Stderr)
    {
        public string CombinedOutput => $"{Stdout}{Environment.NewLine}{Stderr}";
    }
}
