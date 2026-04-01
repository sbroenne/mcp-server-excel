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
public sealed class RangeFormatIssue585RegressionTests : IAsyncLifetime, IAsyncDisposable
{
    private const string IssueSheetName = "Toutes les transactions";
    private const int IssueFillColor = 7949855;
    private const int WhiteFontColor = 16777215;

    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private readonly string _testExcelFile;
    private readonly Pipe _clientToServerPipe = new();
    private readonly Pipe _serverToClientPipe = new();
    private readonly CancellationTokenSource _cts = new();
    private McpClient? _client;
    private Task? _serverTask;
    private string? _sessionId;

    public RangeFormatIssue585RegressionTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Join(Path.GetTempPath(), $"RangeFormatIssue585_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        _testExcelFile = Path.Join(_tempDir, "RangeFormatIssue585.xlsx");
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
                ClientInfo = new() { Name = "RangeFormatIssue585RegressionClient", Version = "1.0.0" },
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
    public async Task FormatRange_MinimalBoldPayload_AppliesViaMcpProtocol()
    {
        await CreateIssueSheetAsync();

        var formatJson = await CallToolAsync("range_format", new Dictionary<string, object?>
        {
            ["action"] = "format-range",
            ["session_id"] = _sessionId,
            ["sheet_name"] = IssueSheetName,
            ["range_address"] = "A1:J1",
            ["bold"] = true
        });
        AssertSuccess(formatJson, "range_format.format-range minimal");

        await CloseSessionAsync(save: true);

        using var batch = ExcelSession.BeginBatch(_testExcelFile);
        Assert.True(ReadCellFormatting(batch, IssueSheetName, "A1").Bold);
        Assert.True(ReadCellFormatting(batch, IssueSheetName, "J1").Bold);
    }

    [Fact]
    public async Task FormatRange_Issue585Payload_AppliesViaMcpProtocol()
    {
        await CreateIssueSheetAsync();

        var formatJson = await CallToolAsync("range_format", new Dictionary<string, object?>
        {
            ["action"] = "format-range",
            ["session_id"] = _sessionId,
            ["sheet_name"] = IssueSheetName,
            ["range_address"] = "A1:J1",
            ["bold"] = true,
            ["fill_color"] = "#1F4E79",
            ["font_color"] = "#FFFFFF"
        });
        AssertSuccess(formatJson, "range_format.format-range issue-585");

        await CloseSessionAsync(save: true);

        using var batch = ExcelSession.BeginBatch(_testExcelFile);
        var a1 = ReadCellFormatting(batch, IssueSheetName, "A1");
        var j1 = ReadCellFormatting(batch, IssueSheetName, "J1");

        Assert.Equal(new CellFormattingState(true, IssueFillColor, WhiteFontColor), a1);
        Assert.Equal(new CellFormattingState(true, IssueFillColor, WhiteFontColor), j1);
    }

    private async Task CreateIssueSheetAsync()
    {
        var createSheetJson = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["session_id"] = _sessionId,
            ["sheet_name"] = IssueSheetName
        });
        AssertSuccess(createSheetJson, "worksheet.create");
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
            $"{operation} returned a non-JSON response. Response: {json}");

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
        catch (Exception ex)
        {
            _output.WriteLine($"Warning: Failed to delete temp directory {_tempDir}: {ex.Message}");
        }
    }

    private static CellFormattingState ReadCellFormatting(IExcelBatch batch, string sheetName, string cellAddress)
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
                    FillColor: interior.Color == null ? null : Convert.ToInt32(interior.Color),
                    FontColor: font.Color == null ? null : Convert.ToInt32(font.Color));
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

    private sealed record CellFormattingState(bool Bold, int? FillColor, int? FontColor);
}
