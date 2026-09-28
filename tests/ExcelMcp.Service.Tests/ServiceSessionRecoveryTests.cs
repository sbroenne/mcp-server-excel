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

    [Fact(Timeout = 60000)]
    public async Task Open_LockedWorkbook_FailsWithoutSessionAndSucceedsAfterUnlock()
    {
        var workbookPath = Path.Combine(_tempDirectory, "LockedWorkbook.xlsx");
        File.Copy(
            Path.Combine(AppContext.BaseDirectory, "TestFiles", "batch-test-static.xlsx"),
            workbookPath);

        var service = new ExcelMcpService();
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

                Assert.False(lockedResponse.Success);
                Assert.Contains("already open", lockedResponse.ErrorMessage, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("close the file", lockedResponse.ErrorMessage, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("exclusive access", lockedResponse.ErrorMessage, StringComparison.OrdinalIgnoreCase);
                Assert.False(string.IsNullOrWhiteSpace(lockedResponse.ExceptionType));

                using var emptyList = ParseSuccessfulResult(
                    await service.ProcessAsync(new ServiceRequest { Command = "session.list" }));
                Assert.Empty(emptyList.RootElement.GetProperty("sessions").EnumerateArray());
            }

            var openResponse = await OpenAsync(service, workbookPath);
            using var openResult = ParseSuccessfulResult(openResponse);
            sessionId = openResult.RootElement.GetProperty("sessionId").GetString();
            Assert.False(string.IsNullOrWhiteSpace(sessionId));

            using var listResult = ParseSuccessfulResult(
                await service.ProcessAsync(new ServiceRequest { Command = "session.list" }));
            var session = Assert.Single(listResult.RootElement.GetProperty("sessions").EnumerateArray());
            Assert.Equal(sessionId, session.GetProperty("sessionId").GetString());
            Assert.Equal(0, session.GetProperty("activeOperations").GetInt32());
            Assert.True(session.GetProperty("canClose").GetBoolean());
        }
        catch (Exception ex)
        {
            failure = ex;
        }

        try
        {
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

        if (failure is not null)
        {
            throw failure;
        }
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
        Assert.True(string.IsNullOrWhiteSpace(response.ErrorMessage));
        Assert.False(string.IsNullOrWhiteSpace(response.Result));
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
