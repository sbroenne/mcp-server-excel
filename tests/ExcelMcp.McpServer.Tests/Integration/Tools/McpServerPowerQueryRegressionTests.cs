// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.IO.Pipelines;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Black-box MCP regressions for Power Query operations that previously hung during
/// synchronous refresh/load paths. These tests exercise the real MCP transport, tool layer,
/// service bridge, and Excel COM automation with explicit per-call timeouts.
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
public class McpServerPowerQueryRegressionTests : IAsyncLifetime, IAsyncDisposable
{
    private static readonly TimeSpan ToolTimeout = TimeSpan.FromSeconds(90);

    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private readonly Pipe _clientToServerPipe = new();
    private readonly Pipe _serverToClientPipe = new();
    private readonly CancellationTokenSource _cts = new();
    private McpClient? _client;
    private Task? _serverTask;

    public McpServerPowerQueryRegressionTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Join(Path.GetTempPath(), $"McpPowerQueryRegression_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
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
                ClientInfo = new() { Name = "PowerQueryRegressionClient", Version = "1.0.0" },
                InitializationTimeout = TimeSpan.FromSeconds(30)
            },
            cancellationToken: _cts.Token);
    }

    public async Task DisposeAsync()
    {
        await DisposeAsyncCore();
    }

    async ValueTask IAsyncDisposable.DisposeAsync()
    {
        await DisposeAsyncCore();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public async Task PowerQuery_Evaluate_CompletesViaMcpProtocol()
    {
        var workbookPath = Path.Join(_tempDir, $"Evaluate_{Guid.NewGuid():N}.xlsx");
        var csvPath = await CreateCsvAsync("evaluate");
        var sessionId = await CreateSessionAsync(workbookPath);

        try
        {
            var mCode = BuildCsvMCode(csvPath);

            var evaluateResult = await CallToolAsync("powerquery", new Dictionary<string, object?>
            {
                ["action"] = "evaluate",
                ["session_id"] = sessionId,
                ["m_code"] = mCode
            }, ToolTimeout);

            AssertSuccess(evaluateResult, "powerquery.evaluate");
            Assert.Contains("Widget", evaluateResult);

            var listSheetsResult = await CallToolAsync("worksheet", new Dictionary<string, object?>
            {
                ["action"] = "list",
                ["session_id"] = sessionId
            }, ToolTimeout);

            AssertSuccess(listSheetsResult, "worksheet.list after powerquery.evaluate");
        }
        finally
        {
            await TryCloseSessionAsync(sessionId);
        }
    }

    [Fact]
    public async Task PowerQuery_LoadToDataModel_CompletesViaMcpProtocol()
    {
        var workbookPath = Path.Join(_tempDir, $"LoadToDataModel_{Guid.NewGuid():N}.xlsx");
        var csvPath = await CreateCsvAsync("loadtodm");
        var sessionId = await CreateSessionAsync(workbookPath);

        try
        {
            var mCode = BuildCsvMCode(csvPath);

            var createResult = await CallToolAsync("powerquery", new Dictionary<string, object?>
            {
                ["action"] = "create",
                ["session_id"] = sessionId,
                ["query_name"] = "CsvData",
                ["m_code"] = mCode,
                ["load_destination"] = "connection-only"
            }, ToolTimeout);

            AssertSuccess(createResult, "powerquery.create");

            var loadResult = await CallToolAsync("powerquery", new Dictionary<string, object?>
            {
                ["action"] = "load-to",
                ["session_id"] = sessionId,
                ["query_name"] = "CsvData",
                ["load_destination"] = "load-to-data-model"
            }, ToolTimeout);

            AssertSuccess(loadResult, "powerquery.load-to data-model");

            var listTablesResult = await CallToolAsync("datamodel", new Dictionary<string, object?>
            {
                ["action"] = "list-tables",
                ["session_id"] = sessionId
            }, ToolTimeout);

            AssertSuccess(listTablesResult, "datamodel.list-tables after powerquery.load-to");
            Assert.Contains("CsvData", listTablesResult);

            var listSessionsResult = await CallToolAsync("file", new Dictionary<string, object?>
            {
                ["action"] = "list"
            }, ToolTimeout);

            AssertSuccess(listSessionsResult, "file.list after powerquery.load-to");
        }
        finally
        {
            await TryCloseSessionAsync(sessionId);
        }
    }

    private async Task DisposeAsyncCore()
    {
        await _cts.CancelAsync();

        if (_client != null)
        {
            await _client.DisposeAsync();
        }

        await _clientToServerPipe.Writer.CompleteAsync();
        await _serverToClientPipe.Reader.CompleteAsync();

        if (_serverTask != null)
        {
            try
            {
                await _serverTask.WaitAsync(TimeSpan.FromSeconds(10));
            }
            catch (OperationCanceledException)
            {
            }
            catch (TimeoutException)
            {
                _output.WriteLine("Warning: MCP server did not stop within timeout.");
            }
        }

        _clientToServerPipe.Writer.Complete();
        _clientToServerPipe.Reader.Complete();
        _serverToClientPipe.Writer.Complete();
        _serverToClientPipe.Reader.Complete();

        Program.ResetTestTransport();

        if (Directory.Exists(_tempDir))
        {
            try
            {
                Directory.Delete(_tempDir, recursive: true);
            }
            catch
            {
            }
        }
    }

    private async Task<string> CreateSessionAsync(string workbookPath)
    {
        var createResult = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["path"] = workbookPath,
            ["show"] = false
        }, ToolTimeout);

        AssertSuccess(createResult, "file.create");
        var sessionId = GetJsonProperty(createResult, "session_id");
        Assert.False(string.IsNullOrWhiteSpace(sessionId));
        return sessionId!;
    }

    private async Task TryCloseSessionAsync(string? sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return;
        }

        try
        {
            await CallToolAsync("file", new Dictionary<string, object?>
            {
                ["action"] = "close",
                ["session_id"] = sessionId,
                ["save"] = false
            }, ToolTimeout);
        }
        catch (Exception ex)
        {
            _output.WriteLine($"Cleanup warning while closing session {sessionId}: {ex.Message}");
        }
    }

    private async Task<string> CallToolAsync(string toolName, Dictionary<string, object?> arguments, TimeSpan timeout)
    {
        using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(_cts.Token);
        timeoutCts.CancelAfter(timeout);

        try
        {
            var result = await _client!.CallToolAsync(toolName, arguments, cancellationToken: timeoutCts.Token);
            var textBlock = result.Content.OfType<TextContentBlock>().FirstOrDefault();

            if (textBlock?.Text == null)
            {
                throw new InvalidOperationException($"Unexpected response from {toolName}");
            }

            return textBlock.Text;
        }
        catch (OperationCanceledException ex) when (!_cts.IsCancellationRequested)
        {
            throw new TimeoutException($"{toolName} did not complete within {timeout.TotalSeconds} seconds.", ex);
        }
    }

    private static async Task<string> CreateCsvAsync(string prefix)
    {
        var csvPath = Path.Join(Path.GetTempPath(), $"{prefix}_{Guid.NewGuid():N}.csv");
        var csvContent = "Product,Quantity\nWidget,10\nGadget,20";
        await File.WriteAllTextAsync(csvPath, csvContent);
        return csvPath;
    }

    private static string BuildCsvMCode(string csvPath)
    {
        return $@"let
    Source = Csv.Document(File.Contents(""{csvPath.Replace("\\", "\\\\")}""), [Delimiter = "","", Columns = 2, Encoding = 1252, QuoteStyle = QuoteStyle.None]),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars = true])
in
    PromotedHeaders";
    }

    private static void AssertSuccess(string jsonResult, string operationName)
    {
        var json = JsonDocument.Parse(jsonResult);

        if (json.RootElement.TryGetProperty("error", out var error))
        {
            Assert.Fail($"{operationName} failed with error: {error.GetString()}");
        }

        if (json.RootElement.TryGetProperty("Success", out var successPascal) && !successPascal.GetBoolean())
        {
            var errorMessage = json.RootElement.TryGetProperty("ErrorMessage", out var errorPascal)
                ? errorPascal.GetString()
                : "Unknown error";
            Assert.Fail($"{operationName} returned Success=false: {errorMessage}");
        }

        if (json.RootElement.TryGetProperty("success", out var successCamel) && !successCamel.GetBoolean())
        {
            var errorMessage = json.RootElement.TryGetProperty("errorMessage", out var errorCamel)
                ? errorCamel.GetString()
                : "Unknown error";
            Assert.Fail($"{operationName} returned success=false: {errorMessage}");
        }
    }

    private static string? GetJsonProperty(string jsonResult, string propertyName)
    {
        var json = JsonDocument.Parse(jsonResult);
        return json.RootElement.TryGetProperty(propertyName, out var property) ? property.GetString() : null;
    }
}