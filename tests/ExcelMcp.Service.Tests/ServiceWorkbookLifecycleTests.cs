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
    [Fact]
    public async Task CreateHidden_ListReportsVisibilityAndCloseState()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var workbookPath = Path.Join(directory, "hidden.xlsx");
            var sessionId = await CreateSessionAsync(service, workbookPath);
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
            Assert.False(session.GetProperty("isExcelVisible").GetBoolean());
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
            await Task.WhenAll(Enumerable.Range(0, workflowCount).Select(async index =>
            {
                var workbookPath = Path.Join(directory, $"concurrent-{index}.xlsx");
                var marker = $"Marker-{index}";
                var sessionId = await CreateSessionAsync(service, workbookPath);
                sessions[sessionId] = 0;
                await WriteMarkerAsync(service, sessionId, marker);
                await CloseSessionAsync(service, sessionId, save: true);
                sessions.TryRemove(sessionId, out _);

                var reopenedSessionId = await OpenSessionAsync(service, workbookPath);
                sessions[reopenedSessionId] = 0;
                Assert.Equal(
                    marker,
                    await ReadMarkerAsync(service, reopenedSessionId));
                await CloseSessionAsync(service, reopenedSessionId, save: false);
                sessions.TryRemove(reopenedSessionId, out _);
            }));

            Assert.Equal(0, service.SessionCount);
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
        string workbookPath)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.create",
            Args = JsonSerializer.Serialize(new
            {
                filePath = workbookPath,
                show = false,
                timeoutSeconds = 120
            }, ServiceProtocol.JsonOptions)
        });
        return GetSessionId(response);
    }

    private static async Task<string> OpenSessionAsync(
        ExcelMcpService service,
        string workbookPath)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(new
            {
                filePath = workbookPath,
                show = false,
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

    private static async Task<string?> ReadMarkerAsync(
        ExcelMcpService service,
        string sessionId)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.get-values",
            SessionId = sessionId,
            Args = """{"sheetName":"Sheet1","rangeAddress":"A1"}"""
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
}
