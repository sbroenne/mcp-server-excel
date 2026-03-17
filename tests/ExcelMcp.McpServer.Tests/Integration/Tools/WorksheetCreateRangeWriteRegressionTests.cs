using System.IO.Pipelines;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Regression coverage for worksheet-create followed immediately by non-A1 range writes.
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Worksheets")]
[Trait("RequiresExcel", "true")]
public class WorksheetCreateRangeWriteRegressionTests : IAsyncLifetime, IAsyncDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private readonly string _testExcelFile;

    private readonly Pipe _clientToServerPipe = new();
    private readonly Pipe _serverToClientPipe = new();
    private readonly CancellationTokenSource _cts = new();
    private McpClient? _client;
    private Task? _serverTask;
    private string? _sessionId;

    public WorksheetCreateRangeWriteRegressionTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Join(Path.GetTempPath(), $"WsCreateWriteRegression_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        _testExcelFile = Path.Join(_tempDir, "WorksheetCreateRangeWrite.xlsx");
    }

    public async Task InitializeAsync()
    {
        Program.ConfigureTestTransport(_clientToServerPipe, _serverToClientPipe);
        _serverTask = Program.Main([]);
        await Task.Delay(100);

        _client = await McpClient.CreateAsync(
            new StreamClientTransport(
                serverInput: _clientToServerPipe.Writer.AsStream(),
                serverOutput: _serverToClientPipe.Reader.AsStream()),
            clientOptions: new McpClientOptions
            {
                ClientInfo = new() { Name = "WorksheetCreateRangeWriteClient", Version = "1.0.0" },
                InitializationTimeout = TimeSpan.FromSeconds(30)
            },
            cancellationToken: _cts.Token);

        var createJson = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["path"] = _testExcelFile
        });

        using var createDoc = JsonDocument.Parse(createJson);
        Assert.True(createDoc.RootElement.GetProperty("success").GetBoolean(), $"Failed to create test file: {createJson}");

        _sessionId = createDoc.RootElement.GetProperty("session_id").GetString();
        Assert.False(string.IsNullOrWhiteSpace(_sessionId));
    }

    [Fact]
    public async Task CreateWorksheet_ThenSetValues_ToNonA1Range_SucceedsViaMcpProtocol()
    {
        var sheetName = "Bug2Data";
        var values = new List<List<object?>>
        {
            new() { "R1C1", "R1C2", "R1C3", "R1C4", "R1C5", "R1C6", "R1C7" },
            new() { "R2C1", "R2C2", "R2C3", "R2C4", "R2C5", "R2C6", "R2C7" },
            new() { "R3C1", "R3C2", "R3C3", "R3C4", "R3C5", "R3C6", "R3C7" },
            new() { "R4C1", "R4C2", "R4C3", "R4C4", "R4C5", "R4C6", "R4C7" },
            new() { "R5C1", "R5C2", "R5C3", "R5C4", "R5C5", "R5C6", "R5C7" },
            new() { "R6C1", "R6C2", "R6C3", "R6C4", "R6C5", "R6C6", "R6C7" },
            new() { "R7C1", "R7C2", "R7C3", "R7C4", "R7C5", "R7C6", "R7C7" },
            new() { "R8C1", "R8C2", "R8C3", "R8C4", "R8C5", "R8C6", "R8C7" }
        };

        var createSheetJson = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["session_id"] = _sessionId,
            ["sheet_name"] = sheetName
        });
        AssertSuccess(createSheetJson, "worksheet.create");

        var setValuesJson = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-values",
            ["path"] = _testExcelFile,
            ["session_id"] = _sessionId,
            ["sheet_name"] = sheetName,
            ["range_address"] = "A3:G10",
            ["values"] = values
        });
        AssertSuccess(setValuesJson, "range.set-values");

        var getValuesJson = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "get-values",
            ["path"] = _testExcelFile,
            ["session_id"] = _sessionId,
            ["sheet_name"] = sheetName,
            ["range_address"] = "A3:G10"
        });

        using var getValuesDoc = JsonDocument.Parse(getValuesJson);
        var root = getValuesDoc.RootElement;
        Assert.True(root.GetProperty("success").GetBoolean(), $"range.get-values failed: {getValuesJson}");
        Assert.Equal(8, root.GetProperty("rowCount").GetInt32());
        Assert.Equal(7, root.GetProperty("columnCount").GetInt32());

        var returnedValues = root.GetProperty("values");
        for (int rowIndex = 0; rowIndex < values.Count; rowIndex++)
        {
            for (int columnIndex = 0; columnIndex < values[rowIndex].Count; columnIndex++)
            {
                Assert.Equal(values[rowIndex][columnIndex]?.ToString(), returnedValues[rowIndex][columnIndex].GetString());
            }
        }

        var a1Json = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "get-values",
            ["path"] = _testExcelFile,
            ["session_id"] = _sessionId,
            ["sheet_name"] = sheetName,
            ["range_address"] = "A1"
        });

        using var a1Doc = JsonDocument.Parse(a1Json);
        Assert.True(a1Doc.RootElement.GetProperty("success").GetBoolean(), $"A1 read failed: {a1Json}");
        var a1Cell = a1Doc.RootElement.GetProperty("values")[0][0];
        Assert.True(a1Cell.ValueKind is JsonValueKind.Null || (a1Cell.ValueKind == JsonValueKind.String && string.IsNullOrEmpty(a1Cell.GetString())));
    }

    private async Task<string> CallToolAsync(string toolName, Dictionary<string, object?> arguments)
    {
        var result = await _client!.CallToolAsync(toolName, arguments, cancellationToken: _cts.Token);

        Assert.NotNull(result);
        Assert.NotNull(result.Content);
        Assert.NotEmpty(result.Content);

        var textBlock = result.Content.OfType<TextContentBlock>().FirstOrDefault();
        Assert.NotNull(textBlock);

        return textBlock.Text;
    }

    private static void AssertSuccess(string json, string operation)
    {
        using var doc = JsonDocument.Parse(json);
        Assert.True(doc.RootElement.GetProperty("success").GetBoolean(), $"{operation} failed: {json}");
    }

    public async Task DisposeAsync()
    {
        await CleanupAsync();
    }

    async ValueTask IAsyncDisposable.DisposeAsync()
    {
        await CleanupAsync();
        GC.SuppressFinalize(this);
    }

    private async Task CleanupAsync()
    {
        if (!string.IsNullOrEmpty(_sessionId) && _client != null)
        {
            try
            {
                await CallToolAsync("file", new Dictionary<string, object?>
                {
                    ["action"] = "close",
                    ["session_id"] = _sessionId,
                    ["save"] = false
                });
            }
            catch (Exception ex)
            {
                _output.WriteLine($"Warning: Failed to close session: {ex.Message}");
            }
        }

        if (_client != null)
        {
            await _client.DisposeAsync();
        }

        _clientToServerPipe.Writer.Complete();
        _serverToClientPipe.Writer.Complete();

        if (_serverTask != null)
        {
            var shutdownTimeout = Task.Delay(TimeSpan.FromSeconds(10));
            var completed = await Task.WhenAny(_serverTask, shutdownTimeout);

            if (completed == shutdownTimeout)
            {
                await _cts.CancelAsync();
                try
                {
                    await _serverTask;
                }
                catch (OperationCanceledException)
                {
                }
            }
        }

        Program.ResetTestTransport();
        _cts.Dispose();

        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, recursive: true);
            }
        }
        catch
        {
        }
    }
}