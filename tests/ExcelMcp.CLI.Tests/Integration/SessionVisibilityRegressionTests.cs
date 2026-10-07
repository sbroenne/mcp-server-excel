using System.Diagnostics;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Regressions for CLI session visibility wiring. These tests prove the CLI forwards
/// --show to the daemon/service layer and that session list exposes the effective flag.
/// </summary>
[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "CLI")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "true")]
public sealed class SessionVisibilityRegressionTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly List<string> _filesToDelete = [];

    public SessionVisibilityRegressionTests(ITestOutputHelper output) => _output = output;

    [Fact]
    public async Task SessionOpen_WithShow_ReportsVisibleSession()
    {
        var workbookPath = CreateExistingWorkbookPath("session-open-visible");
        string? sessionId = null;
        Exception? failure = null;

        try
        {
            var (openResult, openJsonDocument) = await CliProcessHelper.RunJsonAsync(
                ["session", "open", workbookPath, "--show"],
                timeoutMs: 30000,
                diagnosticLabel: "session-open-visible");
            using var openJson = openJsonDocument;

            _output.WriteLine($"[session-open-visible] Stdout: {openResult.Stdout}");
            _output.WriteLine($"[session-open-visible] Stderr: {openResult.Stderr}");

            Assert.Equal(0, openResult.ExitCode);
            Assert.True(openJson.RootElement.GetProperty("success").GetBoolean());

            sessionId = openJson.RootElement.GetProperty("sessionId").GetString();
            Assert.False(string.IsNullOrWhiteSpace(sessionId));

            await AssertSessionVisibilityAsync(sessionId!, expectedVisible: true, "session-open-visible-list");
            await AssertLiveExcelVisibilityAsync(sessionId!, expectedVisible: true, "session-open-visible-window");
        }
        catch (Exception exception)
        {
            failure = exception;
        }
        finally
        {
            var cleanup = await Record.ExceptionAsync(() => CloseSessionIfNeededAsync(sessionId, "session-open-visible-close"));
            if (cleanup is not null)
                failure = failure is null ? cleanup : new AggregateException(failure, cleanup);
        }
        if (failure is not null)
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
    }

    [ConfiguredIrmFact]
    [Trait("RunType", "OnDemand")]
    public async Task SessionOpen_RealIrmWorkbookWithShow_ReturnsWithinTimeoutBudget_WhenConfigured()
    {
        var irmTestFile = GetConfiguredIrmTestFilePath()
            ?? throw new InvalidOperationException("Configured IRM test fixture was unavailable after test discovery.");

        string? sessionId = null;
        Exception? failure = null;
        try
        {
            var stopwatch = Stopwatch.StartNew();
            var (openResult, openJsonDocument) = await CliProcessHelper.RunJsonAsync(
                ["session", "open", irmTestFile, "--show", "--timeout-seconds", "15"],
                timeoutMs: 20000,
                diagnosticLabel: "session-open-irm-show");
            stopwatch.Stop();
            using var openJson = openJsonDocument;

            _output.WriteLine($"[session-open-irm-show] Elapsed: {stopwatch.Elapsed.TotalSeconds:F1}s");
            _output.WriteLine($"[session-open-irm-show] Stdout: {openResult.Stdout}");
            _output.WriteLine($"[session-open-irm-show] Stderr: {openResult.Stderr}");

            Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(20),
                "CLI session open must return within the requested timeout budget for protected workbooks.");
            Assert.True(openJson.RootElement.TryGetProperty("success", out var successProp));

            if (successProp.GetBoolean())
            {
                sessionId = openJson.RootElement.GetProperty("sessionId").GetString();
                Assert.False(string.IsNullOrWhiteSpace(sessionId));
                await AssertSessionVisibilityAsync(sessionId!, expectedVisible: true, "session-open-irm-show-list");
            }
            else
            {
                var error = openJson.RootElement.GetProperty("error").GetString();
                Assert.False(string.IsNullOrWhiteSpace(error));
            }

            var (listResult, listJsonDocument) = await CliProcessHelper.RunJsonAsync(
                ["session", "list"],
                timeoutMs: 10000,
                diagnosticLabel: "session-open-irm-show-list-all");
            using var listJson = listJsonDocument;

            Assert.Equal(0, listResult.ExitCode);
            Assert.True(listJson.RootElement.GetProperty("success").GetBoolean());
        }
        catch (Exception exception)
        {
            failure = exception;
        }
        finally
        {
            var cleanup = await Record.ExceptionAsync(() => CloseSessionIfNeededAsync(sessionId, "session-open-irm-show-close"));
            if (cleanup is not null)
                failure = failure is null ? cleanup : new AggregateException(failure, cleanup);
        }
        if (failure is not null)
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
    }

    public void Dispose()
    {
        foreach (var file in _filesToDelete)
        {
            if (!File.Exists(file))
            {
                continue;
            }

            File.Delete(file);
        }

        GC.SuppressFinalize(this);
    }

    private async Task AssertSessionVisibilityAsync(string sessionId, bool expectedVisible, string diagnosticLabel)
    {
        var (listResult, listJsonDocument) = await CliProcessHelper.RunJsonAsync(
            ["session", "list"],
            timeoutMs: 10000,
            diagnosticLabel: diagnosticLabel);
        using var listJson = listJsonDocument;

        _output.WriteLine($"[{diagnosticLabel}] Stdout: {listResult.Stdout}");
        _output.WriteLine($"[{diagnosticLabel}] Stderr: {listResult.Stderr}");

        Assert.Equal(0, listResult.ExitCode);
        Assert.True(listJson.RootElement.GetProperty("success").GetBoolean());

        var session = listJson.RootElement
            .GetProperty("sessions")
            .EnumerateArray()
            .FirstOrDefault(item => string.Equals(
                item.GetProperty("sessionId").GetString(),
                sessionId,
                StringComparison.OrdinalIgnoreCase));

        Assert.Equal(JsonValueKind.Object, session.ValueKind);
        Assert.Equal(expectedVisible, session.GetProperty("isExcelVisible").GetBoolean());
    }

    private static async Task AssertLiveExcelVisibilityAsync(
        string sessionId,
        bool expectedVisible,
        string diagnosticLabel)
    {
        var (result, jsonDocument) = await CliProcessHelper.RunJsonAsync(
            ["window", "get-info", "--session", sessionId],
            timeoutMs: 10000,
            diagnosticLabel: diagnosticLabel);
        using var json = jsonDocument;

        Assert.Equal(0, result.ExitCode);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal(expectedVisible, json.RootElement.GetProperty("isVisible").GetBoolean());
    }

    private async Task CloseSessionIfNeededAsync(string? sessionId, string diagnosticLabel)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return;
        }

        var closeResult = await CliProcessHelper.RunAsync(
            ["session", "close", "--session", sessionId],
            timeoutMs: 30000,
            diagnosticLabel: diagnosticLabel);

        _output.WriteLine($"[{diagnosticLabel}] Stdout: {closeResult.Stdout}");
        _output.WriteLine($"[{diagnosticLabel}] Stderr: {closeResult.Stderr}");
        Assert.True(closeResult.ExitCode == 0, closeResult.Stdout + closeResult.Stderr);
    }

    private string CreateExistingWorkbookPath(string prefix)
    {
        var workbookPath = Path.Combine(Path.GetTempPath(), $"{prefix}-{Guid.NewGuid():N}.xlsx");
        var sourceWorkbookPath = Path.Combine(
            GetRepositoryRoot(),
            "tests",
            "ExcelMcp.ComInterop.Tests",
            "Integration",
            "Session",
            "TestFiles",
            "batch-test-static.xlsx");
        Assert.True(File.Exists(sourceWorkbookPath), $"Static workbook fixture missing: {sourceWorkbookPath}");

        File.Copy(sourceWorkbookPath, workbookPath, overwrite: true);
        _filesToDelete.Add(workbookPath);
        return workbookPath;
    }

    private static string? GetConfiguredIrmTestFilePath()
    {
        var irmTestFile = Environment.GetEnvironmentVariable("TEST_IRM_FILE");
        return !string.IsNullOrWhiteSpace(irmTestFile) && File.Exists(irmTestFile)
            ? Path.GetFullPath(irmTestFile)
            : null;
    }

    private static string GetRepositoryRoot() =>
        Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
}
