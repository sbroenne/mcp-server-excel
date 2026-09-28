using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Sbroenne.ExcelMcp.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("Feature", "File")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Medium")]
public sealed class SessionLifecycleContractTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDirectory;

    public SessionLifecycleContractTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDirectory = Path.Join(Path.GetTempPath(), $"SessionLifecycleContractTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDirectory);
    }

    [Fact]
    public async Task SessionHelp_AdvertisesOnlyCanonicalLifecycleActions()
    {
        var result = await CliProcessHelper.RunAsync(["session", "--help"], timeoutMs: 10_000);
        var output = result.Stdout + result.Stderr;

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("create", output, StringComparison.Ordinal);
        Assert.Contains("open", output, StringComparison.Ordinal);
        Assert.Contains("close", output, StringComparison.Ordinal);
        Assert.Contains("list", output, StringComparison.Ordinal);
        Assert.Contains("test", output, StringComparison.Ordinal);
        Assert.DoesNotContain("save a session", output, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("session save", output, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ServiceStatus_ResponsiveServiceWithoutDaemonMutex_ReportsRunning()
    {
        var (result, json) = await CliProcessHelper.RunJsonAsync(
            ["service", "status"],
            timeoutMs: 15_000,
            diagnosticLabel: "service-status-in-process-host");
        using (json)
        {
            Assert.Equal(0, result.ExitCode);
            Assert.True(json.RootElement.GetProperty("success").GetBoolean());
            Assert.True(json.RootElement.GetProperty("running").GetBoolean());
            Assert.Equal(
                "running",
                json.RootElement.GetProperty("daemonState").GetString());
        }
    }

    [Fact]
    public async Task SessionSave_IsRejectedAsAnUnknownCommand()
    {
        var result = await CliProcessHelper.RunAsync(
            ["session", "save", "--session", "missing-session"],
            timeoutMs: 10_000);
        var output = result.Stdout + result.Stderr;

        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("Unknown command", output, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("save", output, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task SessionTest_RelativePath_ReturnsSharedValidationError()
    {
        var result = await InProcessCliHelper.RunWithServiceAsync(
            ["session", "test", @"relative\book.xlsx"]);
        var output = result.Stdout + result.Stderr;

        Assert.Equal(1, result.ExitCode);
        Assert.Contains("absolute Windows path", output, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("open")]
    [InlineData("create")]
    public async Task SessionOpenAndCreate_RelativePath_ReturnSharedValidationError(
        string action)
    {
        var result = await CliProcessHelper.RunAsync(
            ["session", action, @"relative\book.txt"],
            timeoutMs: 10_000);
        var output = result.Stdout + result.Stderr;

        Assert.Equal(1, result.ExitCode);
        Assert.Contains("absolute Windows path", output, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("missing.xlsx", false, false, false, false, false, false)]
    [InlineData("invalid.txt", true, false, false, false, false, false)]
    [InlineData("corrupt.xlsx", true, false, false, false, false, false)]
    [InlineData("normal.xlsx", true, false, false, false, false, false)]
    [InlineData("protected.xlsx", true, true, false, false, true, true)]
    [InlineData("protected-modern.xlsx", true, true, false, false, true, true)]
    public async Task SessionTest_ReturnsCanonicalFileMetadata(
        string fileName,
        bool createFile,
        bool irmSignature,
        bool expectedValid,
        bool expectedCanOpen,
        bool expectedReadOnly,
        bool expectedVisible)
    {
        var path = Path.Join(_tempDirectory, fileName);
        if (createFile)
        {
            if (irmSignature)
            {
                OleDataSpaceTestFile.Write(
                    path,
                    fileName.Contains("modern", StringComparison.Ordinal)
                        ? "DRMEncryptedDataSpace"
                        : "\tDRMDataSpace");
            }
            else
            {
                if (string.Equals(fileName, "normal.xlsx", StringComparison.Ordinal))
                {
                    CopyOpaqueWorkbook(path);
                }
                else
                {
                    await File.WriteAllTextAsync(path, "not an Excel workbook");
                }
            }
        }

        var (result, json) = await InProcessCliHelper.RunJsonWithServiceAsync(
            ["session", "test", path]);
        using (json)
        {
            _output.WriteLine(result.Stdout);
            Assert.Equal(expectedCanOpen ? 0 : 1, result.ExitCode);
            Assert.Equal(expectedCanOpen, json.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal(createFile, json.RootElement.GetProperty("exists").GetBoolean());
            Assert.Equal(expectedValid, json.RootElement.GetProperty("isValid").GetBoolean());
            Assert.Equal(expectedCanOpen, json.RootElement.GetProperty("canOpen").GetBoolean());
            Assert.Equal(irmSignature, json.RootElement.GetProperty("isIrmProtected").GetBoolean());
            Assert.Equal(expectedReadOnly, json.RootElement.GetProperty("willOpenReadOnly").GetBoolean());
            Assert.Equal(expectedVisible, json.RootElement.GetProperty("requiresVisibleSession").GetBoolean());
            Assert.Equal(Path.GetFullPath(path), json.RootElement.GetProperty("filePath").GetString());
            Assert.Equal(Path.GetExtension(path), json.RootElement.GetProperty("extension").GetString());
            Assert.True(json.RootElement.TryGetProperty("size", out _));
            Assert.True(json.RootElement.TryGetProperty("lastModified", out _));
            Assert.False(json.RootElement.TryGetProperty("isError", out _),
                "File preflight is a diagnostic result, not a tool execution failure.");
            if (!expectedCanOpen)
            {
                Assert.False(string.IsNullOrWhiteSpace(json.RootElement.GetProperty("message").GetString()));
            }
        }
    }

    private static void CopyOpaqueWorkbook(string path)
    {
        using var source = typeof(ExcelMcpService).Assembly.GetManifestResourceStream(
            "Sbroenne.ExcelMcp.Service.Mac.Blank.xlsx");
        Assert.NotNull(source);
        using var destination = File.Create(path);
        source!.CopyTo(destination);
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(_tempDirectory, recursive: true);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }
}
