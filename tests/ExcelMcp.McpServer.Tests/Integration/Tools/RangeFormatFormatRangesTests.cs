using System.IO.Pipelines;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "true")]
public class RangeFormatFormatRangesTests : IAsyncLifetime, IAsyncDisposable
{
    private const int YellowFillColor = 65535;
    private const int CenterAlignment = -4108;
    private static readonly string[] SharedTargetRanges = ["A1:A2", "C1:C2"];
    private static readonly string[] InvalidTargetRanges = ["A1:A2", "NotARange"];

    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private readonly string _testExcelFile;
    private readonly Pipe _clientToServerPipe = new();
    private readonly Pipe _serverToClientPipe = new();
    private readonly CancellationTokenSource _cts = new();
    private McpClient? _client;
    private Task? _serverTask;
    private string? _sessionId;

    public RangeFormatFormatRangesTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Join(Path.GetTempPath(), $"RangeFormatRanges_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        _testExcelFile = Path.Join(_tempDir, "RangeFormatFormatRanges.xlsx");
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
                ClientInfo = new() { Name = "RangeFormatFormatRangesClient", Version = "1.0.0" },
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
    public async Task FormatRanges_AppliesSharedFormattingToEachTargetRange_ViaMcpProtocol()
    {
        const string sheetName = "Bug4Data";

        var createSheetJson = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["session_id"] = _sessionId,
            ["sheet_name"] = sheetName
        });
        AssertSuccess(createSheetJson, "worksheet.create");

        var formatJson = await CallToolAsync("range_format", new Dictionary<string, object?>
        {
            ["action"] = "format-ranges",
            ["session_id"] = _sessionId,
            ["sheet_name"] = sheetName,
            ["range_addresses"] = SharedTargetRanges,
            ["bold"] = true,
            ["fill_color"] = "#FFFF00",
            ["horizontal_alignment"] = "center"
        });
        AssertSuccess(formatJson, "range_format.format-ranges");

        await CloseSessionAsync(save: true);

        using var batch = ExcelSession.BeginBatch(_testExcelFile);
        Assert.Equal(new CellFormattingState(true, YellowFillColor, CenterAlignment), ReadCellFormattingState(batch, sheetName, "A1"));
        Assert.Equal(new CellFormattingState(true, YellowFillColor, CenterAlignment), ReadCellFormattingState(batch, sheetName, "A2"));
        Assert.Equal(new CellFormattingState(true, YellowFillColor, CenterAlignment), ReadCellFormattingState(batch, sheetName, "C1"));
        Assert.Equal(new CellFormattingState(true, YellowFillColor, CenterAlignment), ReadCellFormattingState(batch, sheetName, "C2"));
    }

    [Fact]
    public async Task FormatRanges_WithInvalidTarget_FailsWithoutApplyingEarlierRanges_ViaMcpProtocol()
    {
        const string sheetName = "Bug4Invalid";

        var createSheetJson = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["session_id"] = _sessionId,
            ["sheet_name"] = sheetName
        });
        AssertSuccess(createSheetJson, "worksheet.create");

        var formatJson = await CallToolAsync("range_format", new Dictionary<string, object?>
        {
            ["action"] = "format-ranges",
            ["session_id"] = _sessionId,
            ["sheet_name"] = sheetName,
            ["range_addresses"] = InvalidTargetRanges,
            ["bold"] = true,
            ["fill_color"] = "#FFFF00",
            ["horizontal_alignment"] = "center"
        });

        using (var formatDoc = JsonDocument.Parse(formatJson))
        {
            Assert.False(formatDoc.RootElement.GetProperty("success").GetBoolean(), $"range_format.format-ranges should have failed: {formatJson}");
        }

        await CloseSessionAsync(save: true);

        using var batch = ExcelSession.BeginBatch(_testExcelFile);
        Assert.Equal(ReadCellFormattingState(batch, sheetName, "B1"), ReadCellFormattingState(batch, sheetName, "A1"));
        Assert.Equal(ReadCellFormattingState(batch, sheetName, "B2"), ReadCellFormattingState(batch, sheetName, "A2"));
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
        Assert.True(
            json.TrimStart().StartsWith('{'),
            $"{operation} returned a non-JSON response. This indicates the action is not yet wired into the generated MCP contract or threw before returning an OperationResult. Response: {json}");

        using var doc = JsonDocument.Parse(json);
        Assert.True(doc.RootElement.GetProperty("success").GetBoolean(), $"{operation} failed: {json}");
    }

    private async Task CloseSessionAsync(bool save)
    {
        if (string.IsNullOrWhiteSpace(_sessionId) || _client is null)
        {
            return;
        }

        await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "close",
            ["session_id"] = _sessionId,
            ["save"] = save
        });

        _sessionId = null;
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
        if (!string.IsNullOrWhiteSpace(_sessionId) && _client != null)
        {
            try
            {
                await CloseSessionAsync(save: false);
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

    private static CellFormattingState ReadCellFormattingState(IExcelBatch batch, string sheetName, string cellAddress)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? font = null;
            dynamic? interior = null;

            try
            {
                sheet = ctx.Book.Worksheets[sheetName];
                range = sheet.Range[cellAddress];
                font = range.Font;
                interior = range.Interior;

                return new CellFormattingState(
                    Bold: Convert.ToBoolean(font.Bold),
                    FillColor: Convert.ToInt32(interior.Color),
                    HorizontalAlignment: Convert.ToInt32(range.HorizontalAlignment));
            }
            finally
            {
                ComUtilities.Release(ref interior!);
                ComUtilities.Release(ref font!);
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    private readonly record struct CellFormattingState(bool Bold, int FillColor, int HorizontalAlignment);
}