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
        try
        {
            var (createResult, createJsonDocument) = await CliProcessHelper.RunJsonAsync(
                ["session", "create", _workbookPath, "--timeout", "30"],
                timeoutMs: 45000,
                diagnosticLabel: "python-timeout-create");
            using var createJson = createJsonDocument;

            Assert.Equal(0, createResult.ExitCode);
            sessionId = createJson.RootElement.GetProperty("sessionId").GetString();
            Assert.False(string.IsNullOrWhiteSpace(sessionId));

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

            var (_, listJsonDocument) = await CliProcessHelper.RunJsonAsync(
                ["session", "list"],
                timeoutMs: 10000,
                diagnosticLabel: "python-timeout-list");
            using var listJson = listJsonDocument;
            var session = listJson.RootElement.GetProperty("sessions")
                .EnumerateArray()
                .Single(item => item.GetProperty("sessionId").GetString() == sessionId);

            Assert.Equal(0, session.GetProperty("activeOperations").GetInt32());
            Assert.True(session.GetProperty("canClose").GetBoolean());
        }
        finally
        {
            if (!string.IsNullOrWhiteSpace(sessionId))
            {
                await CliProcessHelper.RunAsync(
                    ["session", "close", "--session", sessionId],
                    timeoutMs: 30000,
                    diagnosticLabel: "python-timeout-close");
            }
        }
    }

    public void Dispose()
    {
        if (File.Exists(_workbookPath))
        {
            File.Delete(_workbookPath);
        }
    }
}
