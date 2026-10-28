using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Tests.Infrastructure;
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
        Assert.Contains("save", output, StringComparison.OrdinalIgnoreCase);
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
            Assert.True(json.RootElement.GetProperty("data").GetProperty("running").GetBoolean());
            Assert.Equal(
                "running",
                json.RootElement.GetProperty("data").GetProperty("daemonState").GetString());
        }
    }

    [Fact]
    public async Task SessionSave_MissingSession_ReturnsSessionNotFound()
    {
        var result = await CliProcessHelper.RunAsync(
            ["session", "save", "--session", "missing-session"],
            timeoutMs: 10_000);
        var output = result.Stdout + result.Stderr;

        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("SessionNotFound", output, StringComparison.Ordinal);
        Assert.Contains("missing-session", output, StringComparison.Ordinal);
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
    [InlineData("normal.xlsx", true, false, true, true, false, false)]
    [InlineData("protected.xlsx", true, true, false, false, false, true)]
    [InlineData("protected-modern.xlsx", true, true, false, false, false, true)]
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
                    SavedWorkbookTemplates.CopyBlankTo(path);
                }
                else
                {
                    await File.WriteAllTextAsync(path, "not an Excel workbook");
                }
            }
        }

        var originalBytes = createFile ? await File.ReadAllBytesAsync(path) : null;
        var originalModified = createFile ? File.GetLastWriteTime(path) : DateTime.MinValue;

        var (result, json) = await InProcessCliHelper.RunJsonWithServiceAsync(
            ["session", "test", path]);
        using (json)
        {
            _output.WriteLine(result.Stdout);
            Assert.Equal(expectedCanOpen ? 0 : 1, result.ExitCode);
            Assert.Equal(expectedCanOpen, json.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal(createFile, json.RootElement.GetProperty("data").GetProperty("exists").GetBoolean());
            Assert.Equal(expectedValid, json.RootElement.GetProperty("data").GetProperty("isValid").GetBoolean());
            Assert.Equal(expectedCanOpen, json.RootElement.GetProperty("data").GetProperty("canOpen").GetBoolean());
            Assert.Equal(irmSignature, json.RootElement.GetProperty("data").GetProperty("isIrmProtected").GetBoolean());
            Assert.Equal(expectedReadOnly, json.RootElement.GetProperty("data").GetProperty("willOpenReadOnly").GetBoolean());
            Assert.Equal(expectedVisible, json.RootElement.GetProperty("data").GetProperty("requiresVisibleSession").GetBoolean());
            Assert.Equal(Path.GetFullPath(path), json.RootElement.GetProperty("data").GetProperty("filePath").GetString());
            Assert.Equal(Path.GetExtension(path), json.RootElement.GetProperty("data").GetProperty("extension").GetString());
            Assert.Equal(originalBytes?.LongLength ?? 0,
                json.RootElement.GetProperty("data").GetProperty("size").GetInt64());
            Assert.Equal(originalModified,
                json.RootElement.GetProperty("data").GetProperty("lastModified").GetDateTime());
            Assert.False(json.RootElement.TryGetProperty("isError", out _),
                "File preflight is a diagnostic result, not a tool execution failure.");
            if (!expectedCanOpen)
            {
                Assert.False(string.IsNullOrWhiteSpace(json.RootElement.GetProperty("data").GetProperty("message").GetString()));
            }

            Assert.Equal(createFile, File.Exists(path));
            if (originalBytes is not null)
            {
                Assert.Equal(originalBytes, await File.ReadAllBytesAsync(path));
                Assert.Equal(originalModified, File.GetLastWriteTime(path));
            }
        }
    }

    public void Dispose()
    {
        Directory.Delete(_tempDirectory, recursive: true);
    }
}
