using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "SessionLifecycle")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Medium")]
[Collection("ServiceWorkflow")]
public sealed class ServiceSessionRecoveryTests : IDisposable
{
    private readonly string _tempDirectory = Path.Combine(
        Path.GetTempPath(),
        $"ServiceSessionRecoveryTests_{Guid.NewGuid():N}");

    public ServiceSessionRecoveryTests()
    {
        Directory.CreateDirectory(_tempDirectory);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Open_LockedWorkbook_FailsWithoutSessionAndSucceedsAfterUnlock(bool slowStatusRead)
    {
        var workbookPath = Path.Combine(_tempDirectory, "LockedWorkbook.xlsx");
        File.Copy(
            Path.Combine(AppContext.BaseDirectory, "TestFiles", "batch-test-static.xlsx"),
            workbookPath);

        var service = new ExcelMcpService();
        using var statusReadFinished = new ManualResetEventSlim(true);
        ExcelBatch? reopenedBatch = null;
        Exception? failure = null;
        string? sessionId = null;
        try
        {
            using (var fileLock = new FileStream(
                       workbookPath,
                       FileMode.Open,
                       FileAccess.ReadWrite,
                       FileShare.None))
            {
                var lockedResponse = await OpenAsync(service, workbookPath);

                Assert.False(lockedResponse.Success, "Opening the locked workbook unexpectedly succeeded.");
                Assert.Contains("already open", lockedResponse.ErrorMessage, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("close the file", lockedResponse.ErrorMessage, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("exclusive access", lockedResponse.ErrorMessage, StringComparison.OrdinalIgnoreCase);
                Assert.False(string.IsNullOrWhiteSpace(lockedResponse.ExceptionType),
                    "The locked-workbook error did not identify its exception type.");

                using var emptyList = ParseSuccessfulResult(
                    await service.ProcessAsync(new ServiceRequest { Command = "session.list" }));
                Assert.Empty(emptyList.RootElement.GetProperty("sessions").EnumerateArray());
            }

            var openResponse = await OpenAsync(service, workbookPath);
            using var openResult = ParseSuccessfulResult(openResponse);
            sessionId = openResult.RootElement.GetProperty("sessionId").GetString();
            Assert.False(string.IsNullOrWhiteSpace(sessionId),
                "Opening the unlocked workbook returned no session ID.");

            if (slowStatusRead)
            {
                reopenedBatch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(sessionId));
                reopenedBatch.BeforeRefreshStateReadHookForTests = () =>
                {
                    reopenedBatch.BeforeRefreshStateReadHookForTests = null;
                    statusReadFinished.Reset();
                    try
                    {
                        Thread.Sleep(400);
                    }
                    finally
                    {
                        statusReadFinished.Set();
                    }
                };
            }

            using var listResult = JsonDocument.Parse(
                await ServiceWorkbookLifecycleTests.WaitForSettledSessionListAsync(service, sessionId));
            var session = Assert.Single(listResult.RootElement.GetProperty("sessions").EnumerateArray());
            Assert.Equal(sessionId, session.GetProperty("sessionId").GetString());
            Assert.Equal(0, session.GetProperty("activeOperations").GetInt32());
            Assert.True(session.GetProperty("canClose").GetBoolean(),
                $"The reopened session cannot close: excelState={session.GetProperty("excelState")}; " +
                $"blockingReason={(session.TryGetProperty("blockingReason", out var reason) ? reason.ToString() : "none")}; " +
                $"operationPending={session.GetProperty("operationPending")}.");
        }
        catch (Exception ex)
        {
            failure = ex;
        }

        try
        {
            if (reopenedBatch is not null)
                reopenedBatch.BeforeRefreshStateReadHookForTests = null;
            Assert.True(statusReadFinished.Wait(TimeSpan.FromSeconds(5)),
                "The delayed status read did not finish before cleanup.");
            if (sessionId is not null)
            {
                var closeResponse = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.close",
                    SessionId = sessionId,
                    Args = JsonSerializer.Serialize(new { save = false }, ServiceProtocol.JsonOptions)
                });
                Assert.True(closeResponse.Success, closeResponse.ErrorMessage);
            }
        }
        catch (Exception ex)
        {
            failure = PersistentServiceCleanupFailures.Combine(failure, ex);
        }

        try
        {
            using var finalList = ParseSuccessfulResult(
                await service.ProcessAsync(new ServiceRequest { Command = "session.list" }));
            Assert.Empty(finalList.RootElement.GetProperty("sessions").EnumerateArray());
        }
        catch (Exception ex)
        {
            failure = PersistentServiceCleanupFailures.Combine(failure, ex);
        }

        try
        {
            service.Dispose();
        }
        catch (Exception ex)
        {
            failure = PersistentServiceCleanupFailures.Combine(failure, ex);
        }

        PersistentServiceCleanupFailures.ThrowIfFailed(failure);
    }

    private static Task<ServiceResponse> OpenAsync(ExcelMcpService service, string workbookPath)
    {
        return service.ProcessAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(
                new { filePath = workbookPath, show = false },
                ServiceProtocol.JsonOptions)
        });
    }

    private static JsonDocument ParseSuccessfulResult(ServiceResponse response)
    {
        Assert.True(response.Success, response.ErrorMessage);
        Assert.True(string.IsNullOrWhiteSpace(response.ErrorMessage), response.ErrorMessage);
        Assert.False(string.IsNullOrWhiteSpace(response.Result),
            $"Successful {response.Command} response contained no result.");
        return JsonDocument.Parse(response.Result);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDirectory))
        {
            Directory.Delete(_tempDirectory, recursive: true);
        }

        Assert.Empty(SessionManager.GetTrackedExcelProcesses());
    }
}
