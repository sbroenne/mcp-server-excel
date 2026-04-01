using System.IO.Pipelines;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Verifies the MCP worksheet tool exposes the intended rename/create parameter contract
/// through the discoverable tool schema.
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Worksheets")]
public sealed class WorksheetToolSchemaTests : IAsyncLifetime, IAsyncDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly Pipe _clientToServerPipe = new();
    private readonly Pipe _serverToClientPipe = new();
    private readonly CancellationTokenSource _cts = new();

    private McpClient? _client;
    private Task? _serverTask;

    public WorksheetToolSchemaTests(ITestOutputHelper output)
    {
        _output = output;
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
                ClientInfo = new() { Name = "WorksheetSchemaClient", Version = "1.0.0" },
                InitializationTimeout = TimeSpan.FromSeconds(30)
            },
            cancellationToken: _cts.Token);
    }

    [Fact]
    public async Task ListTools_WorksheetSchema_ExposesActionSpecificWorksheetDescriptions()
    {
        var tools = await _client!.ListToolsAsync(cancellationToken: _cts.Token);
        var worksheetTool = tools.SingleOrDefault(tool => tool.Name == "worksheet");

        Assert.NotNull(worksheetTool);

        var schema = worksheetTool!.JsonSchema;
        _output.WriteLine(schema.GetRawText());

        var properties = schema.GetProperty("properties");

        Assert.True(properties.TryGetProperty("sheet_name", out var sheetNameProperty), $"worksheet schema is missing sheet_name: {schema.GetRawText()}");
        Assert.True(properties.TryGetProperty("old_name", out var oldNameProperty), $"worksheet schema is missing old_name: {schema.GetRawText()}");
        Assert.True(properties.TryGetProperty("source_name", out var sourceNameProperty), $"worksheet schema is missing source_name: {schema.GetRawText()}");
        Assert.True(properties.TryGetProperty("target_name", out var targetNameProperty), $"worksheet schema is missing target_name: {schema.GetRawText()}");
        Assert.True(properties.TryGetProperty("new_name", out var newNameProperty), $"worksheet schema is missing new_name: {schema.GetRawText()}");
        Assert.True(properties.TryGetProperty("source_file", out var sourceFileProperty), $"worksheet schema is missing source_file: {schema.GetRawText()}");
        Assert.True(properties.TryGetProperty("source_sheet", out var sourceSheetProperty), $"worksheet schema is missing source_sheet: {schema.GetRawText()}");
        Assert.True(properties.TryGetProperty("target_file", out var targetFileProperty), $"worksheet schema is missing target_file: {schema.GetRawText()}");
        Assert.True(properties.TryGetProperty("target_sheet_name", out var targetSheetNameProperty), $"worksheet schema is missing target_sheet_name: {schema.GetRawText()}");

        Assert.Contains("create", GetDescription(sheetNameProperty), StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("rename", GetDescription(sheetNameProperty), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("rename", GetDescription(oldNameProperty), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("rename", GetDescription(newNameProperty), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("copy", GetDescription(sourceNameProperty), StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("rename", GetDescription(sourceNameProperty), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("copy", GetDescription(targetNameProperty), StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("rename", GetDescription(targetNameProperty), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("copy-to-file", GetDescription(sourceFileProperty), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("move-to-file", GetDescription(sourceFileProperty), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("copy-to-file", GetDescription(sourceSheetProperty), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("move-to-file", GetDescription(sourceSheetProperty), StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("rename", GetDescription(sourceSheetProperty), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("copy-to-file", GetDescription(targetFileProperty), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("move-to-file", GetDescription(targetFileProperty), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("copy-to-file", GetDescription(targetSheetNameProperty), StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("rename", GetDescription(targetSheetNameProperty), StringComparison.OrdinalIgnoreCase);
    }

    private static string GetDescription(JsonElement property)
    {
        Assert.True(property.TryGetProperty("description", out var description), $"Schema property is missing description: {property.GetRawText()}");
        return description.GetString() ?? string.Empty;
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
    }
}
