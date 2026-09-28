using System.Collections.Concurrent;
using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "Session")]
[Trait("Speed", "Medium")]
[Trait("RequiresExcel", "true")]
public sealed class ServiceWorkbookLifecycleTests
{
    [Theory]
    [InlineData("create", false)]
    [InlineData("create", true)]
    [InlineData("open", false)]
    [InlineData("open", true)]
    public async Task CreateAndOpen_ListReportsRequestedVisibilityAndCloseState(
        string action,
        bool show)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var workbookPath = Path.Join(directory, $"{action}-{show}.xlsx");
            if (action == "open")
            {
                File.Copy(
                    Path.Join(AppContext.BaseDirectory, "TestFiles", "batch-test-static.xlsx"),
                    workbookPath);
            }

            var sessionId = action == "create"
                ? await CreateSessionAsync(service, workbookPath, show)
                : await OpenSessionAsync(service, workbookPath, show);
            sessions[sessionId] = 0;

            var list = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.list"
            });

            Assert.True(list.Success, list.ErrorMessage);
            using var result = JsonDocument.Parse(list.Result!);
            var session = Assert.Single(
                result.RootElement.GetProperty("sessions").EnumerateArray(),
                item => item.GetProperty("sessionId").GetString() == sessionId);
            Assert.Equal(show, session.GetProperty("isExcelVisible").GetBoolean());
            Assert.Equal(0, session.GetProperty("activeOperations").GetInt32());
            Assert.True(session.GetProperty("canClose").GetBoolean());
        });
    }

    [Fact]
    public async Task SaveCloseReopen_PersistsMarkerInNewSession()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var workbookPath = Path.Join(directory, "persist.xlsx");
            var sessionId = await CreateSessionAsync(service, workbookPath);
            sessions[sessionId] = 0;
            await WriteMarkerAsync(service, sessionId, "Persisted");
            await CloseSessionAsync(service, sessionId, save: true);
            sessions.TryRemove(sessionId, out _);

            var reopenedSessionId = await OpenSessionAsync(service, workbookPath);
            sessions[reopenedSessionId] = 0;
            Assert.Equal(
                "Persisted",
                await ReadMarkerAsync(service, reopenedSessionId));
        });
    }

    [Fact]
    public async Task ConcurrentWorkbookWorkflows_StayIsolatedAndPersist()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            const int workflowCount = 4;
            var openedCount = 0;
            var allOpened = new TaskCompletionSource<bool>(
                TaskCreationOptions.RunContinuationsAsynchronously);
            var releaseWorkflows = new TaskCompletionSource<bool>(
                TaskCreationOptions.RunContinuationsAsynchronously);
            var startupFailure = new TaskCompletionSource<Exception>(
                TaskCreationOptions.RunContinuationsAsynchronously);
            var workflows = Enumerable.Range(0, workflowCount).Select(index => Task.Run(async () =>
            {
                var workbookPath = Path.Join(directory, $"concurrent-{index}.xlsx");
                var sheetName = $"Data{index}";
                var marker = $"Marker-{index}";
                try
                {
                    var sessionId = await CreateSessionAsync(service, workbookPath);
                    sessions[sessionId] = 0;
                    if (Interlocked.Increment(ref openedCount) == workflowCount)
                    {
                        allOpened.TrySetResult(true);
                    }

                    await releaseWorkflows.Task;
                    await CreateSheetAsync(service, sessionId, sheetName);
                    await WriteWorkflowValuesAsync(service, sessionId, sheetName, marker, index);
                    Assert.Equal(marker, await ReadMarkerAsync(service, sessionId, sheetName));
                    await FormatWorkflowValuesAsync(service, sessionId, sheetName);
                    await CloseSessionAsync(service, sessionId, save: true);
                    sessions.TryRemove(sessionId, out _);

                    Assert.True(File.Exists(workbookPath), $"Expected workbook to exist: {workbookPath}");
                    var reopenedSessionId = await OpenSessionAsync(service, workbookPath);
                    sessions[reopenedSessionId] = 0;
                    var persisted = await ReadMarkerAsync(service, reopenedSessionId, sheetName);
                    await CloseSessionAsync(service, reopenedSessionId, save: false);
                    sessions.TryRemove(reopenedSessionId, out _);
                    return new WorkflowResult(index, workbookPath, persisted);
                }
                catch (Exception ex)
                {
                    startupFailure.TrySetResult(ex);
                    throw;
                }
            })).ToArray();

            try
            {
                var readiness = await Task.WhenAny(allOpened.Task, startupFailure.Task)
                    .WaitAsync(TimeSpan.FromMinutes(2));
                if (readiness == startupFailure.Task)
                {
                    throw await startupFailure.Task;
                }

                Assert.Equal(workflowCount, service.SessionCount);
                var list = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
                Assert.True(list.Success, list.ErrorMessage);
                using (var result = JsonDocument.Parse(list.Result!))
                {
                    var openSessionIds = result.RootElement
                        .GetProperty("sessions")
                        .EnumerateArray()
                        .Select(item => item.GetProperty("sessionId").GetString())
                        .ToArray();
                    Assert.Equal(workflowCount, openSessionIds.Length);
                    Assert.All(sessions.Keys, sessionId => Assert.Contains(sessionId, openSessionIds));
                }
            }
            finally
            {
                releaseWorkflows.TrySetResult(true);
            }

            var results = await Task.WhenAll(workflows);
            Assert.Equal(workflowCount, results.Length);
            Assert.Equal(
                workflowCount,
                results.Select(result => result.FilePath).Distinct(StringComparer.OrdinalIgnoreCase).Count());
            Assert.All(
                results,
                result => Assert.Equal($"Marker-{result.Index}", result.PersistedValue));

            Assert.Equal(0, service.SessionCount);
            var finalList = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
            Assert.True(finalList.Success, finalList.ErrorMessage);
            using var finalResult = JsonDocument.Parse(finalList.Result!);
            Assert.Empty(finalResult.RootElement.GetProperty("sessions").EnumerateArray());
        });
    }

    private static async Task RunWithCleanupAsync(
        Func<ExcelMcpService, string, ConcurrentDictionary<string, byte>, Task> test)
    {
        var directory = Path.Join(
            Path.GetTempPath(),
            $"ServiceWorkbookLifecycleTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        var sessions = new ConcurrentDictionary<string, byte>();
        var service = new ExcelMcpService();
        Exception? failure = null;

        try
        {
            await test(service, directory, sessions);
        }
        catch (Exception ex)
        {
            failure = ex;
        }
        finally
        {
            foreach (var sessionId in sessions.Keys)
            {
                try
                {
                    await CloseSessionAsync(service, sessionId, save: false);
                }
                catch (Exception ex)
                {
                    failure = PersistentServiceCleanupFailures.Combine(failure, ex);
                }
            }

            try
            {
                service.Dispose();
            }
            catch (Exception ex)
            {
                failure = PersistentServiceCleanupFailures.Combine(failure, ex);
            }

            try
            {
                Directory.Delete(directory, recursive: true);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                failure = PersistentServiceCleanupFailures.Combine(failure, ex);
            }
        }

        if (failure is not null)
        {
            throw failure;
        }
    }

    private static async Task<string> CreateSessionAsync(
        ExcelMcpService service,
        string workbookPath,
        bool show = false)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.create",
            Args = JsonSerializer.Serialize(new
            {
                filePath = workbookPath,
                show,
                timeoutSeconds = 120
            }, ServiceProtocol.JsonOptions)
        });
        return GetSessionId(response);
    }

    private static async Task<string> OpenSessionAsync(
        ExcelMcpService service,
        string workbookPath,
        bool show = false)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(new
            {
                filePath = workbookPath,
                show,
                timeoutSeconds = 120
            }, ServiceProtocol.JsonOptions)
        });
        return GetSessionId(response);
    }

    private static async Task WriteMarkerAsync(
        ExcelMcpService service,
        string sessionId,
        string marker)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.set-values",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new
            {
                sheetName = "Sheet1",
                rangeAddress = "A1",
                values = new object?[][] { [marker] }
            }, ServiceProtocol.JsonOptions)
        });
        Assert.True(response.Success, response.ErrorMessage);
    }

    private static async Task CreateSheetAsync(
        ExcelMcpService service,
        string sessionId,
        string sheetName)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "sheet.create",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new { sheetName }, ServiceProtocol.JsonOptions)
        });
        Assert.True(response.Success, response.ErrorMessage);
    }

    private static async Task WriteWorkflowValuesAsync(
        ExcelMcpService service,
        string sessionId,
        string sheetName,
        string marker,
        int index)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.set-values",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new
            {
                sheetName,
                rangeAddress = "A1:A2",
                values = new object?[][] { [marker], [$"File-{index}"] }
            }, ServiceProtocol.JsonOptions)
        });
        Assert.True(response.Success, response.ErrorMessage);
    }

    private static async Task FormatWorkflowValuesAsync(
        ExcelMcpService service,
        string sessionId,
        string sheetName)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "rangeformat.format-range",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new
            {
                sheetName,
                rangeAddress = "A1:A2",
                bold = true
            }, ServiceProtocol.JsonOptions)
        });
        Assert.True(response.Success, response.ErrorMessage);
    }

    private static async Task<string?> ReadMarkerAsync(
        ExcelMcpService service,
        string sessionId,
        string sheetName = "Sheet1")
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.get-values",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new
            {
                sheetName,
                rangeAddress = "A1"
            }, ServiceProtocol.JsonOptions)
        });
        Assert.True(response.Success, response.ErrorMessage);
        using var result = JsonDocument.Parse(response.Result!);
        return result.RootElement.GetProperty("values")[0][0].GetString();
    }

    private static async Task CloseSessionAsync(
        ExcelMcpService service,
        string sessionId,
        bool save)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new { save }, ServiceProtocol.JsonOptions)
        });
        Assert.True(response.Success, response.ErrorMessage);
    }

    private static string GetSessionId(ServiceResponse response)
    {
        Assert.True(response.Success, response.ErrorMessage);
        using var result = JsonDocument.Parse(response.Result!);
        var sessionId = result.RootElement.GetProperty("sessionId").GetString();
        Assert.False(string.IsNullOrWhiteSpace(sessionId));
        return sessionId!;
    }

    private sealed record WorkflowResult(int Index, string FilePath, string? PersistedValue);
}
