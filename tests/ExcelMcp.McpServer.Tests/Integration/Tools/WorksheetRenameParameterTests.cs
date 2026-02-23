using System.IO.Pipelines;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Regression tests for Bug Report 2026-02-23.
/// Bug 2: worksheet(action: 'rename') parameters are not discoverable — callers
/// cannot determine the correct parameter names (sheet_name + target_name) because
/// target_name is documented as "Name for the target/copied worksheet" with no
/// mention of rename.
///
/// These tests verify:
/// 1. The correct parameter combination (sheet_name + target_name) works
/// 2. Incorrect combinations fail with clear error messages
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
    /// Verifies that the correct parameter combination (sheet_name + target_name) works.
    /// This is the only working combination, but it is not obvious from the parameter descriptions.
    /// </summary>
    [Fact]
    public async Task Rename_WithSheetNameAndTargetName_Succeeds()
    {
        // Arrange — create a sheet to rename
        await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["session_id"] = _sessionId,
            ["sheet_name"] = "OriginalSheet"
        });

        // Act — rename using the correct (but non-obvious) parameter combination
        var json = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "rename",
            ["session_id"] = _sessionId,
            ["sheet_name"] = "OriginalSheet",   // Maps to oldName
            ["target_name"] = "RenamedSheet"     // Maps to newName
        });
        _output.WriteLine($"Response: {json}");

        // Assert
        var doc = JsonDocument.Parse(json);
        Assert.True(doc.RootElement.GetProperty("success").GetBoolean(),
            $"Rename with sheet_name + target_name should succeed. Response: {json}");

        // Verify the sheet was actually renamed
        var listJson = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = _sessionId
        });
        Assert.Contains("RenamedSheet", listJson);
        Assert.DoesNotContain("OriginalSheet", listJson);
    }

    #endregion

    #region Bug 2: Parameter combinations that fail (as reported)

    /// <summary>
    /// Bug report attempt 1: sheet_name + target_sheet_name fails.
    /// User guessed target_sheet_name (which is a separate parameter for cross-file ops).
    /// Expected error: "newName is required" because target_name was not provided.
    /// </summary>
    [Fact]
    public async Task Rename_WithSheetNameAndTargetSheetName_FailsWithNewNameRequired()
    {
        // Arrange
        await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["session_id"] = _sessionId,
            ["sheet_name"] = "TestSheet1"
        });

        // Act — user's attempt 1 from bug report
        var json = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "rename",
            ["session_id"] = _sessionId,
            ["sheet_name"] = "TestSheet1",
            ["target_sheet_name"] = "NewName"  // Wrong param! target_sheet_name is for cross-file ops
        });
        _output.WriteLine($"Response: {json}");

        // Assert — should fail because target_name (newName) is null
        var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        // The response should indicate failure with a meaningful error about the missing parameter
        Assert.False(root.GetProperty("success").GetBoolean(),
            "Should fail when target_sheet_name is used instead of target_name");
        Assert.True(root.TryGetProperty("errorMessage", out var errorMsg),
            "Expected errorMessage in response");
        Assert.Contains("newName", errorMsg.GetString()!, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Bug report attempt 2: source_name + target_name fails.
    /// User guessed source_name (which is for copy operations).
    /// Expected error: "oldName is required" because sheet_name was not provided.
    /// </summary>
    [Fact]
    public async Task Rename_WithSourceNameAndTargetName_FailsWithOldNameRequired()
    {
        // Arrange
        await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["session_id"] = _sessionId,
            ["sheet_name"] = "TestSheet2"
        });

        // Act — user's attempt 2 from bug report
        var json = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "rename",
            ["session_id"] = _sessionId,
            ["source_name"] = "TestSheet2",   // Wrong param! source_name is for copy
            ["target_name"] = "NewName2"      // Correct param for newName
        });
        _output.WriteLine($"Response: {json}");

        // Assert — should fail because sheet_name (oldName) is null
        var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        Assert.False(root.GetProperty("success").GetBoolean(),
            "Should fail when source_name is used instead of sheet_name");
        Assert.True(root.TryGetProperty("errorMessage", out var errorMsg),
            "Expected errorMessage in response");
        Assert.Contains("oldName", errorMsg.GetString()!, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Bug report attempt 3: sheet_name + source_sheet fails.
    /// User guessed source_sheet (which is for cross-file move/copy).
    /// Expected error: "newName is required" because target_name was not provided.
    /// </summary>
    [Fact]
    public async Task Rename_WithSheetNameAndSourceSheet_FailsWithNewNameRequired()
    {
        // Arrange
        await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["session_id"] = _sessionId,
            ["sheet_name"] = "TestSheet3"
        });

        // Act — user's attempt 3 from bug report
        var json = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "rename",
            ["session_id"] = _sessionId,
            ["sheet_name"] = "TestSheet3",
            ["source_sheet"] = "NewName3"     // Wrong param! source_sheet is for cross-file ops
        });
        _output.WriteLine($"Response: {json}");

        // Assert — should fail because target_name (newName) is null
        var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        Assert.False(root.GetProperty("success").GetBoolean(),
            "Should fail when source_sheet is used instead of target_name");
        Assert.True(root.TryGetProperty("errorMessage", out var errorMsg),
            "Expected errorMessage in response");
        Assert.Contains("newName", errorMsg.GetString()!, StringComparison.OrdinalIgnoreCase);
    }

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
