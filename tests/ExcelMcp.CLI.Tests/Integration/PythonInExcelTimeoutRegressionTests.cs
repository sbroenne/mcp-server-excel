using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Regression coverage for Python-in-Excel polling deadlines that exceed the session timeout.
/// </summary>
[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "PythonInExcel")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "true")]
public sealed class PythonInExcelTimeoutRegressionTests : IDisposable
{
    private readonly string _workbookPath = Path.Combine(
        Path.GetTempPath(),
        $"excelmcp-python-timeout-{Guid.NewGuid():N}.xlsx");

    [Fact]
    public async Task GetResult_WaitExceedsSessionTimeout_IsRejectedAndSessionRemainsClosable()
    {
        string? sessionId = null;
        Exception? failure = null;
        try
        {
            var (createResult, createJsonDocument) = await CliProcessHelper.RunJsonAsync(
                ["session", "create", _workbookPath, "--timeout-seconds", "30"],
                timeoutMs: 45000,
                diagnosticLabel: "python-timeout-create");
            using var createJson = createJsonDocument;

            Assert.Equal(0, createResult.ExitCode);
            Assert.True(createJson.RootElement.GetProperty("success").GetBoolean());
            sessionId = createJson.RootElement.GetProperty("sessionId").GetString();
            Assert.False(string.IsNullOrWhiteSpace(sessionId));
            var seed = await CliProcessHelper.RunAsync(
                ["range", "set-values", "--session", sessionId!, "--sheet-name", "Sheet1",
                    "--range-address", "A1", "--values", "[[17]]"]);
            Assert.True(seed.ExitCode == 0, seed.Stdout + seed.Stderr);

            var (getResult, getJsonDocument) = await CliProcessHelper.RunJsonAsync(
                [
                    "pythoninexcel", "get-result",
                    "--sheet", "Sheet1",
                    "--range", "A1",
                    "--max-wait-seconds", "60",
                    "--session", sessionId!
                ],
                timeoutMs: 10000,
                diagnosticLabel: "python-timeout-get-result");
            using var getJson = getJsonDocument;

            Assert.Equal(1, getResult.ExitCode);
            Assert.False(getJson.RootElement.GetProperty("success").GetBoolean());
            Assert.Contains("session operation timeout", getJson.RootElement.GetProperty("error").GetString());

            var (listResult, listJsonDocument) = await CliProcessHelper.RunJsonAsync(
                ["session", "list"],
                timeoutMs: 10000,
                diagnosticLabel: "python-timeout-list");
            using var listJson = listJsonDocument;
            Assert.Equal(0, listResult.ExitCode);
            Assert.True(listJson.RootElement.GetProperty("success").GetBoolean());
            var session = listJson.RootElement.GetProperty("sessions")
                .EnumerateArray()
                .Single(item => item.GetProperty("sessionId").GetString() == sessionId);

            Assert.Equal(0, session.GetProperty("activeOperations").GetInt32());
            Assert.True(session.GetProperty("canClose").GetBoolean());
            var (readResult, readDocument) = await CliProcessHelper.RunJsonAsync(
                ["range", "get-values", "--session", sessionId!, "--sheet-name", "Sheet1",
                    "--range-address", "A1"]);
            using (readDocument)
            {
                Assert.Equal(0, readResult.ExitCode);
                Assert.True(readDocument.RootElement.GetProperty("success").GetBoolean());
                Assert.Equal(17, readDocument.RootElement.GetProperty("values")[0][0].GetInt32());
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
                    var closed = await CliProcessHelper.RunAsync(
                        ["session", "close", "--session", sessionId, "--save", "false"],
                        timeoutMs: 30000,
                        diagnosticLabel: "python-timeout-close");
                    Assert.True(closed.ExitCode == 0, closed.Stdout + closed.Stderr);
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
        if (File.Exists(_workbookPath))
        {
            File.Delete(_workbookPath);
        }
    }
}
