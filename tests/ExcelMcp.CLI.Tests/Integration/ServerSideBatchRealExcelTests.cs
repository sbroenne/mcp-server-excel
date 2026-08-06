using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Batch")]
[Trait("Speed", "Medium")]
public sealed class ServerSideBatchRealExcelTests : IClassFixture<TempDirectoryFixture>
{
    private readonly TempDirectoryFixture _fixture;

    public ServerSideBatchRealExcelTests(TempDirectoryFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public async Task SessionBatch_ExecutesInOrderAndStopsAtExactFailureIndex()
    {
        var workbookPath = CoreTestHelper.CreateUniqueTestFile(
            nameof(ServerSideBatchRealExcelTests),
            nameof(SessionBatch_ExecutesInOrderAndStopsAtExactFailureIndex),
            _fixture.TempDir,
            ".xlsx");
        var stateRoot = Path.Combine(_fixture.TempDir, $"batch-state-{Guid.NewGuid():N}");

        using var service = new ExcelMcpService(stateRoot);
        var open = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(new { filePath = workbookPath, show = false }, ServiceProtocol.JsonOptions),
            Source = "test",
        });
        Assert.True(open.Success, open.ErrorMessage);
        using var openDocument = JsonDocument.Parse(open.Result!);
        var sessionId = openDocument.RootElement.GetProperty("sessionId").GetString()!;

        var success = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.batch",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new
            {
                operations = new object[]
                {
                    SetValue("A2", 10),
                    SetValue("A3", 20),
                    SetValue("A4", 30),
                },
                stopOnError = true,
            }, ServiceProtocol.JsonOptions),
            Source = "test",
        });

        Assert.True(success.Success, success.ErrorMessage);
        using (var successDocument = JsonDocument.Parse(success.Result!))
        {
            Assert.True(successDocument.RootElement.GetProperty("completed").GetBoolean());
            Assert.Equal(3, successDocument.RootElement.GetProperty("results").GetArrayLength());
            Assert.False(successDocument.RootElement.TryGetProperty("failedIndex", out _));
        }
        Assert.Equal([10d, 20d, 30d], await ReadColumnAsync(service, sessionId, "A2:A4"));

        var failure = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.batch",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new
            {
                operations = new object[]
                {
                    SetValue("A5", 40),
                    new
                    {
                        command = "range.set-values",
                        args = new { sheetName = "MissingSheet", rangeAddress = "A1", values = new object?[][] { [50] } },
                    },
                    SetValue("A6", 60),
                },
                stopOnError = true,
            }, ServiceProtocol.JsonOptions),
            Source = "test",
        });

        Assert.False(failure.Success);
        using (var failureDocument = JsonDocument.Parse(failure.Result!))
        {
            Assert.False(failureDocument.RootElement.GetProperty("completed").GetBoolean());
            Assert.Equal(1, failureDocument.RootElement.GetProperty("failedIndex").GetInt32());
            Assert.Equal(2, failureDocument.RootElement.GetProperty("results").GetArrayLength());
        }
        Assert.Equal([40d, null], await ReadColumnAsync(service, sessionId, "A5:A6"));

        var close = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = sessionId,
            Args = "{\"save\":false}",
            Source = "test",
        });
        Assert.True(close.Success, close.ErrorMessage);
    }

    private static object SetValue(string address, double value) => new
    {
        command = "range.set-values",
        args = new { sheetName = "Sheet1", rangeAddress = address, values = new object?[][] { [value] } },
    };

    private static async Task<double?[]> ReadColumnAsync(
        ExcelMcpService service,
        string sessionId,
        string rangeAddress)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.get-values",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new { sheetName = "Sheet1", rangeAddress }, ServiceProtocol.JsonOptions),
            Source = "test",
        });
        Assert.True(response.Success, response.ErrorMessage);
        using var document = JsonDocument.Parse(response.Result!);
        return document.RootElement.GetProperty("values")
            .EnumerateArray()
            .Select(row => row[0].ValueKind == JsonValueKind.Null ? (double?)null : row[0].GetDouble())
            .ToArray();
    }
}
