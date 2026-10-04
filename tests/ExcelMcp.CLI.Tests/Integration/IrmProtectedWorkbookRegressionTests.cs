using System.Diagnostics;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// CLI regressions for IRM/AIP-protected workbook handling.
/// Uses a deterministic fake-signature file for fail-fast guidance coverage and
/// an opt-in real protected workbook via TEST_IRM_FILE for local startup validation.
/// </summary>
[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "CLI")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "true")]
public sealed class IrmProtectedWorkbookRegressionTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _fakeIrmFile;

    public IrmProtectedWorkbookRegressionTests(ITestOutputHelper output)
    {
        _output = output;
        _fakeIrmFile = Path.Combine(Path.GetTempPath(), $"CliFakeIrm_{Guid.NewGuid():N}.xlsx");
        OleDataSpaceTestFile.Write(_fakeIrmFile, "\tDRMDataSpace");
    }

    private static string? GetConfiguredIrmTestFilePath()
    {
        var irmTestFile = Environment.GetEnvironmentVariable("TEST_IRM_FILE");
        return !string.IsNullOrWhiteSpace(irmTestFile) && File.Exists(irmTestFile)
            ? Path.GetFullPath(irmTestFile)
            : null;
    }

    [Fact]
    public async Task SessionOpenHelp_ListsShowOption_ForProtectedWorkbooks()
    {
        var result = await CliProcessHelper.RunAsync(
            ["session", "open", "--help"],
            timeoutMs: 10000,
            diagnosticLabel: "irm-session-open-help");

        _output.WriteLine($"[irm-session-open-help] Stdout: {result.Stdout}");
        _output.WriteLine($"[irm-session-open-help] Stderr: {result.Stderr}");

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("--show", result.Stdout, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("IRM", result.Stdout, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    public async Task SessionOpen_IrmSignatureFile_WithoutShow_FailsFastWithInteractiveGuidance()
    {
        var original = File.ReadAllBytes(_fakeIrmFile);
        var sessionsBefore = await ReadSessionsAsync();
        var stopwatch = Stopwatch.StartNew();

        var (result, json) = await CliProcessHelper.RunJsonAsync(
            ["session", "open", _fakeIrmFile, "--timeout", "15"],
            timeoutMs: 20000,
            diagnosticLabel: "irm-session-open-headless");

        stopwatch.Stop();

        _output.WriteLine($"[irm-session-open-headless] Stdout: {result.Stdout}");
        _output.WriteLine($"[irm-session-open-headless] Stderr: {result.Stderr}");

        Assert.Equal(1, result.ExitCode);
        Assert.False(json.RootElement.GetProperty("success").GetBoolean());

        var error = json.RootElement.GetProperty("error").GetString();
        Assert.NotNull(error);
        Assert.Contains("IRM/AIP-protected workbook", error, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("show=true", error, StringComparison.OrdinalIgnoreCase);
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(20),
            "CLI session open must fail fast for protected workbooks when --show is omitted.");
        Assert.Equal(original, File.ReadAllBytes(_fakeIrmFile));
        Assert.Equal(sessionsBefore, await ReadSessionsAsync());
    }

    [ConfiguredIrmFact]
    [Trait("RunType", "OnDemand")]
    public async Task SessionOpen_RealIrmWorkbook_WithShow_CompletesWithinTimeoutBudget_WhenConfigured()
    {
        var irmTestFile = GetConfiguredIrmTestFilePath()
            ?? throw new InvalidOperationException("Configured IRM test fixture was unavailable after test discovery.");

        Assert.True(FileAccessValidator.IsIrmProtected(irmTestFile),
            "TEST_IRM_FILE must point to a real IRM/AIP-protected workbook for this regression.");

        var stopwatch = Stopwatch.StartNew();
        string? sessionId = null;
        Exception? failure = null;

        try
        {
            var (result, json) = await CliProcessHelper.RunJsonAsync(
                ["session", "open", irmTestFile, "--show", "--timeout", "15"],
                timeoutMs: 20000,
                diagnosticLabel: "irm-session-open-visible");

            stopwatch.Stop();

            _output.WriteLine($"[irm-session-open-visible] Stdout: {result.Stdout}");
            _output.WriteLine($"[irm-session-open-visible] Stderr: {result.Stderr}");

            Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(20),
                "CLI session open must return within the requested timeout budget for protected workbooks.");

            if (result.ExitCode == 0)
            {
                Assert.True(json.RootElement.GetProperty("success").GetBoolean());
                sessionId = json.RootElement.GetProperty("sessionId").GetString();
                Assert.False(string.IsNullOrWhiteSpace(sessionId));
            }
            else
            {
                Assert.False(json.RootElement.GetProperty("success").GetBoolean());
                var error = json.RootElement.GetProperty("error").GetString();
                Assert.False(string.IsNullOrWhiteSpace(error));
                Assert.DoesNotContain("show=true", error, StringComparison.OrdinalIgnoreCase);
            }
        }
        catch (Exception exception)
        {
            failure = exception;
        }
        finally
        {
            if (!string.IsNullOrWhiteSpace(sessionId))
            {
                var cleanup = await Record.ExceptionAsync(async () =>
                {
                    var closeResult = await CliProcessHelper.RunAsync(
                        ["session", "close", "--session", sessionId, "--save", "false"],
                        timeoutMs: 30000,
                        diagnosticLabel: "irm-session-close-visible");
                    Assert.True(closeResult.ExitCode == 0, closeResult.Stdout + closeResult.Stderr);
                });
                if (cleanup is not null)
                    failure = failure is null ? cleanup : new AggregateException(failure, cleanup);
            }
        }
        if (failure is not null)
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
    }

    public void Dispose()
    {
        if (File.Exists(_fakeIrmFile))
        {
            File.Delete(_fakeIrmFile);
        }

        GC.SuppressFinalize(this);
    }

    private static async Task<string[]> ReadSessionsAsync()
    {
        var result = await CliProcessHelper.RunAsync(["session", "list"]);
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        using var document = JsonDocument.Parse(result.Stdout);
        Assert.True(document.RootElement.GetProperty("success").GetBoolean());
        return document.RootElement.GetProperty("sessions").EnumerateArray()
            .Select(session => session.GetProperty("sessionId").GetString()!)
            .Order(StringComparer.Ordinal).ToArray();
    }
}
