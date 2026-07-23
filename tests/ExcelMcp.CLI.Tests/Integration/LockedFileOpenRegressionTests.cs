using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// End-to-end regressions for locked workbook open failures through the CLI surface.
/// Verifies the CLI returns an actionable error and the workbook can be opened
/// successfully immediately after the external lock is released.
/// </summary>
[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "CLI")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "true")]
public sealed class LockedFileOpenRegressionTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _testFile;

    public LockedFileOpenRegressionTests(ITestOutputHelper output)
    {
        _output = output;
        _testFile = Path.Combine(Path.GetTempPath(), $"LockedFileOpen_{Guid.NewGuid():N}.xlsx");

        ExcelSession.CreateNew(
            _testFile,
            isMacroEnabled: false,
            (ctx, ct) => 0,
            CancellationToken.None);
    }

    [Fact]
    public async Task SessionOpen_FileLockedByAnotherProcess_ReturnsActionableError_AndNextOpenSucceeds()
    {
        using (var fileLock = new FileStream(_testFile, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
        {
            var (lockedResult, lockedJson) = await CliProcessHelper.RunJsonAsync(
                ["session", "open", _testFile],
                timeoutMs: 30000,
                diagnosticLabel: "locked-session-open");

            _output.WriteLine($"[locked-session-open] Stdout: {lockedResult.Stdout}");
            _output.WriteLine($"[locked-session-open] Stderr: {lockedResult.Stderr}");

            Assert.Equal(1, lockedResult.ExitCode);
            Assert.False(lockedJson.RootElement.GetProperty("success").GetBoolean());
            Assert.True(lockedJson.RootElement.GetProperty("isError").GetBoolean());

            var error = lockedJson.RootElement.GetProperty("error").GetString();
            var errorMessage = lockedJson.RootElement.GetProperty("errorMessage").GetString();
            Assert.NotNull(error);
            Assert.Equal(error, errorMessage);
            Assert.Contains("already open", error, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("close the file", error, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("exclusive access", error, StringComparison.OrdinalIgnoreCase);
            Assert.False(string.IsNullOrWhiteSpace(lockedJson.RootElement.GetProperty("exceptionType").GetString()));
        }

        var (listAfterFailureResult, listAfterFailureJson) = await CliProcessHelper.RunJsonAsync(
            ["session", "list"],
            timeoutMs: 10000,
            diagnosticLabel: "locked-session-list-after-failure");

        _output.WriteLine($"[locked-session-list-after-failure] Stdout: {listAfterFailureResult.Stdout}");
        _output.WriteLine($"[locked-session-list-after-failure] Stderr: {listAfterFailureResult.Stderr}");

        Assert.Equal(0, listAfterFailureResult.ExitCode);
        Assert.Equal(0, listAfterFailureJson.RootElement.GetProperty("sessions").GetArrayLength());

        string? sessionId = null;
        try
        {
            var (openResult, openJson) = await CliProcessHelper.RunJsonAsync(
                ["session", "open", _testFile],
                timeoutMs: 30000,
                diagnosticLabel: "locked-session-open-after-release");

            _output.WriteLine($"[locked-session-open-after-release] Stdout: {openResult.Stdout}");
            _output.WriteLine($"[locked-session-open-after-release] Stderr: {openResult.Stderr}");

            Assert.Equal(0, openResult.ExitCode);
            Assert.True(openJson.RootElement.GetProperty("success").GetBoolean());

            sessionId = openJson.RootElement.GetProperty("sessionId").GetString();
            Assert.False(string.IsNullOrWhiteSpace(sessionId));
        }
        finally
        {
            if (!string.IsNullOrWhiteSpace(sessionId))
            {
                var closeResult = await CliProcessHelper.RunAsync(
                    ["session", "close", "--session", sessionId, "--save", "false"],
                    timeoutMs: 30000,
                    diagnosticLabel: "locked-session-close-after-release");

                _output.WriteLine($"[locked-session-close-after-release] Stdout: {closeResult.Stdout}");
                _output.WriteLine($"[locked-session-close-after-release] Stderr: {closeResult.Stderr}");
            }
        }
    }

    public void Dispose()
    {
        if (File.Exists(_testFile))
        {
#pragma warning disable CA1031
            try
            {
                File.Delete(_testFile);
            }
            catch
            {
            }
#pragma warning restore CA1031
        }

        GC.SuppressFinalize(this);
    }
}
