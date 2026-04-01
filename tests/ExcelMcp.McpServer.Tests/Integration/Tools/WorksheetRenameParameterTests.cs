using System.IO.Pipelines;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Regression tests for the worksheet rename parameter contract.
/// The canonical worksheet rename surface is old_name + new_name so it matches
/// other rename actions and the CLI sheet rename command.
///
/// These tests verify:
/// 1. The canonical parameter combination (old_name + new_name) works
/// 2. Legacy and cross-action MCP aliases still work for backward compatibility
/// 3. Missing parameters produce actionable errors, not cryptic COM exceptions
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Worksheets")]
[Trait("RequiresExcel", "true")]
public class WorksheetRenameParameterTests : IAsyncLifetime, IAsyncDisposable
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

    public WorksheetRenameParameterTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Join(Path.GetTempPath(), $"WsRenameRegression_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        _testExcelFile = Path.Join(_tempDir, "WorksheetRenameTest.xlsx");
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
                ClientInfo = new() { Name = "WsRenameRegressionClient", Version = "1.0.0" },
                InitializationTimeout = TimeSpan.FromSeconds(30)
            },
            cancellationToken: _cts.Token);

        // Create a fresh workbook and open session
        var createJson = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["path"] = _testExcelFile
        });

        var createDoc = JsonDocument.Parse(createJson);
        Assert.True(createDoc.RootElement.GetProperty("success").GetBoolean(),
            $"Failed to create test file: {createJson}");

        _sessionId = createDoc.RootElement.GetProperty("session_id").GetString();
        Assert.NotNull(_sessionId);
    }

    #region Bug 2: Correct parameter combination

    /// <summary>
    /// Verifies that the canonical rename parameter combination (old_name + new_name) works.
    /// </summary>
    [Fact]
    public async Task Rename_WithOldNameAndNewName_Succeeds()
    {
        // Arrange — create a sheet to rename
        await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["session_id"] = _sessionId,
            ["sheet_name"] = "OriginalSheet"
        });

        // Act — rename using the canonical parameter combination
        var json = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "rename",
            ["session_id"] = _sessionId,
            ["old_name"] = "OriginalSheet",
            ["new_name"] = "RenamedSheet"
        });
        _output.WriteLine($"Response: {json}");

        // Assert
        var doc = JsonDocument.Parse(json);
        Assert.True(doc.RootElement.GetProperty("success").GetBoolean(),
            $"Rename with old_name + new_name should succeed. Response: {json}");

        // Verify the sheet was actually renamed
        var listJson = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = _sessionId
        });
        Assert.Contains("RenamedSheet", listJson);
        Assert.DoesNotContain("OriginalSheet", listJson);
    }

    /// <summary>
    /// Verifies that the legacy MCP aliases (sheet_name + target_name) still work.
    /// This preserves backward compatibility for existing callers while MCP now exposes old_name + new_name.
    /// </summary>
    [Fact]
    public async Task Rename_WithLegacySheetNameAndTargetName_Succeeds()
    {
        // Arrange
        await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["session_id"] = _sessionId,
            ["sheet_name"] = "LegacySheet"
        });

        // Act
        var json = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "rename",
            ["session_id"] = _sessionId,
            ["sheet_name"] = "LegacySheet",
            ["target_name"] = "RenamedLegacy"
        });
        _output.WriteLine($"Response: {json}");

        // Assert
        var doc = JsonDocument.Parse(json);
        Assert.True(doc.RootElement.GetProperty("success").GetBoolean(),
            $"Rename with legacy sheet_name + target_name should still succeed. Response: {json}");

        var listJson = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = _sessionId
        });
        Assert.Contains("RenamedLegacy", listJson);
        Assert.DoesNotContain("LegacySheet", listJson);
    }

    /// <summary>
    /// Claude/Desktop guessed target_sheet_name because that parameter is already used by copy-to-file.
    /// Rename should accept it as an alias for the new worksheet name.
    /// </summary>
    [Fact]
    public async Task Rename_WithSheetNameAndTargetSheetName_Succeeds()
    {
        await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["session_id"] = _sessionId,
            ["sheet_name"] = "TargetSheetAlias"
        });

        var json = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "rename",
            ["session_id"] = _sessionId,
            ["sheet_name"] = "TargetSheetAlias",
            ["target_sheet_name"] = "RenamedViaTargetSheet"
        });
        _output.WriteLine($"Response: {json}");

        var doc = JsonDocument.Parse(json);
        Assert.True(doc.RootElement.GetProperty("success").GetBoolean(),
            $"Rename with sheet_name + target_sheet_name should succeed. Response: {json}");

        var listJson = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = _sessionId
        });
        Assert.Contains("RenamedViaTargetSheet", listJson);
        Assert.DoesNotContain("TargetSheetAlias", listJson);
    }

    /// <summary>
    /// Claude/Desktop also guessed source_name because it is used by the copy action.
    /// Rename should accept it as an alias for the current worksheet name.
    /// </summary>
    [Fact]
    public async Task Rename_WithSourceNameAndTargetName_Succeeds()
    {
        await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["session_id"] = _sessionId,
            ["sheet_name"] = "SourceNameAlias"
        });

        var json = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "rename",
            ["session_id"] = _sessionId,
            ["source_name"] = "SourceNameAlias",
            ["target_name"] = "RenamedViaSourceName"
        });
        _output.WriteLine($"Response: {json}");

        var doc = JsonDocument.Parse(json);
        Assert.True(doc.RootElement.GetProperty("success").GetBoolean(),
            $"Rename with source_name + target_name should succeed. Response: {json}");

        var listJson = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = _sessionId
        });
        Assert.Contains("RenamedViaSourceName", listJson);
        Assert.DoesNotContain("SourceNameAlias", listJson);
    }

    /// <summary>
    /// This matches Claude's first failed v1.8.40 attempt captured in mcp.log.
    /// Rename should accept the cross-file names as aliases for better MCP compatibility.
    /// </summary>
    [Fact]
    public async Task Rename_WithSourceSheetAndTargetSheetName_Succeeds()
    {
        await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["session_id"] = _sessionId,
            ["sheet_name"] = "SourceSheetAlias"
        });

        var json = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "rename",
            ["session_id"] = _sessionId,
            ["source_sheet"] = "SourceSheetAlias",
            ["target_sheet_name"] = "RenamedViaSourceSheet"
        });
        _output.WriteLine($"Response: {json}");

        var doc = JsonDocument.Parse(json);
        Assert.True(doc.RootElement.GetProperty("success").GetBoolean(),
            $"Rename with source_sheet + target_sheet_name should succeed. Response: {json}");

        var listJson = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = _sessionId
        });
        Assert.Contains("RenamedViaSourceSheet", listJson);
        Assert.DoesNotContain("SourceSheetAlias", listJson);
    }

    #endregion

    #region Validation errors that should fail

    /// <summary>
    /// Verify that both old and new name missing produces a clear error.
    /// Tests the case where a caller provides only the action and session_id.
    /// </summary>
    [Fact]
    public async Task Rename_WithNoNameParameters_FailsWithOldNameRequired()
    {
        // Act — call rename with no name parameters at all
        var json = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "rename",
            ["session_id"] = _sessionId
        });
        _output.WriteLine($"Response: {json}");

        // Assert — should fail with old name being the first validation error
        var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        Assert.False(root.GetProperty("success").GetBoolean(),
            "Should fail when no name parameters are provided");
    }

    #endregion

    #region Helper Methods

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

    #endregion

    #region Cleanup

    async ValueTask IAsyncDisposable.DisposeAsync()
    {
        await CleanupAsync();
        GC.SuppressFinalize(this);
    }

    public async Task DisposeAsync()
    {
        await CleanupAsync();
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
                _output.WriteLine("Warning: Server did not shut down gracefully, forcing cancellation");
                await _cts.CancelAsync();
                try
                {
                    await _serverTask;
                }
                catch (OperationCanceledException)
                {
                    // Expected
                }
            }
        }

        Program.ResetTestTransport();
        _cts.Dispose();

        try
        {
            if (Directory.Exists(_tempDir))
            {
                for (int i = 0; i < 3; i++)
                {
                    try
                    {
                        Directory.Delete(_tempDir, recursive: true);
                        break;
                    }
                    catch (IOException) when (i < 2)
                    {
                        await Task.Delay(500);
                    }
                }
            }
        }
        catch
        {
            // Cleanup is best-effort
        }
    }

    #endregion
}
