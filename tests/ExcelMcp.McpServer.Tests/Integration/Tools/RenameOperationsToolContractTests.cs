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
/// Contract tests for rename operations verifying deterministic MCP behavior.
/// These tests ensure rename operations return proper JSON responses (not exceptions)
/// for business logic errors, enabling LLM agents to handle outcomes predictably.
///
/// Key contracts verified:
/// - Missing object → JSON with success=false, "not found" in errorMessage
/// - Name conflict → JSON with success=false, "exists/conflict" in errorMessage
/// - Invalid name → JSON with success=false, validation error
/// - No-op (same name) → JSON with success=true
/// - Success → JSON with objectType, oldName, newName
/// - Excel limitation → JSON with success=false, clear explanation
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "RenameContract")]
[Trait("RequiresExcel", "true")]
public class RenameOperationsToolContractTests : IAsyncLifetime, IAsyncDisposable
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

    public RenameOperationsToolContractTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Join(Path.GetTempPath(), $"RenameContract_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        _testExcelFile = Path.Join(_tempDir, "RenameContractTest.xlsx");

        _output.WriteLine($"Test directory: {_tempDir}");
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
                ClientInfo = new() { Name = "RenameContractTestClient", Version = "1.0.0" },
                InitializationTimeout = TimeSpan.FromSeconds(30)
            },
            cancellationToken: _cts.Token);

        _output.WriteLine($"✓ Connected to server: {_client.ServerInfo?.Name} v{_client.ServerInfo?.Version}");

        // Create a fresh workbook and open session in one call (Create)
        var createJson = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["path"] = _testExcelFile
        });

        var createDoc = JsonDocument.Parse(createJson);
        Assert.True(createDoc.RootElement.GetProperty("success").GetBoolean(),
            $"Failed to create test file: {createJson}");

        _sessionId = createDoc.RootElement.GetProperty("sessionId").GetString();
        Assert.NotNull(_sessionId);

        _output.WriteLine($"✓ Created test file and opened session: {_sessionId}");
    }

    #region Power Query Rename Contract Tests

    /// <summary>
    /// Verifies that renaming a non-existent query returns JSON with success=false (not exception).
    /// Contract: Missing object → success=false with "not found" message.
    /// </summary>
    [Fact]
    public async Task PowerQueryRename_MissingQuery_ReturnsJsonWithSuccessFalse()
    {
        // Arrange - no query exists

        // Act - attempt to rename non-existent query
        var json = await CallToolAsync("powerquery", new Dictionary<string, object?>
        {
            ["action"] = "rename",
            ["session_id"] = _sessionId,
            ["old_name"] = "NonExistentQuery",
            ["new_name"] = "NewName"
        });
        _output.WriteLine($"Response: {json}");

        // Assert - should return JSON with success=false, NOT throw exception
        var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        Assert.False(root.GetProperty("success").GetBoolean(), "Expected success=false for missing query");
        Assert.True(root.TryGetProperty("errorMessage", out var errorMsg), "Expected errorMessage property");
        Assert.Contains("not found", errorMsg.GetString()!, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Verifies that renaming to a conflicting name returns JSON with success=false (not exception).
    /// Contract: Name conflict → success=false with "exists" or "conflict" message.
    /// </summary>
    [Fact]
    public async Task PowerQueryRename_NameConflict_ReturnsJsonWithSuccessFalse()
    {
        // Arrange - create two queries
        await CreatePowerQuery("QueryA", "let x = 1 in x");
        await CreatePowerQuery("QueryB", "let y = 2 in y");

        // Act - try to rename QueryA to QueryB (conflict)
        var json = await CallToolAsync("powerquery", new Dictionary<string, object?>
        {
            ["action"] = "rename",
            ["session_id"] = _sessionId,
            ["old_name"] = "QueryA",
            ["new_name"] = "QueryB" // Already exists!
        });
        _output.WriteLine($"Response: {json}");

        // Assert - should return JSON with success=false, NOT throw exception
        var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        Assert.False(root.GetProperty("success").GetBoolean(), "Expected success=false for name conflict");
        Assert.True(root.TryGetProperty("errorMessage", out var errorMsg), "Expected errorMessage property");
        var errorText = errorMsg.GetString()!;
        Assert.True(
            errorText.Contains("exists", StringComparison.OrdinalIgnoreCase) ||
            errorText.Contains("conflict", StringComparison.OrdinalIgnoreCase) ||
            errorText.Contains("already", StringComparison.OrdinalIgnoreCase),
            $"Expected conflict-related error message, got: {errorText}");
    }

    /// <summary>
    /// Verifies that renaming with empty new name returns JSON with success=false (not exception).
    /// Contract: Invalid name → success=false with validation message.
    /// </summary>
    [Fact]
    public async Task PowerQueryRename_EmptyNewName_ReturnsJsonWithSuccessFalse()
    {
        // Arrange - create a query
        await CreatePowerQuery("ValidQuery", "let x = 1 in x");

        // Act - try to rename with empty new name
        var json = await CallToolAsync("powerquery", new Dictionary<string, object?>
        {
            ["action"] = "rename",
            ["session_id"] = _sessionId,
            ["old_name"] = "ValidQuery",
            ["new_name"] = "   " // Empty after trim
        });
        _output.WriteLine($"Response: {json}");

        // Assert - should return JSON with success=false, NOT throw exception
        var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        Assert.False(root.GetProperty("success").GetBoolean(), "Expected success=false for empty name");
        Assert.True(root.TryGetProperty("errorMessage", out var errorMsg), "Expected errorMessage property");
        var errorText = errorMsg.GetString()!;
        Assert.True(
            errorText.Contains("empty", StringComparison.OrdinalIgnoreCase) ||
            errorText.Contains("blank", StringComparison.OrdinalIgnoreCase) ||
            errorText.Contains("invalid", StringComparison.OrdinalIgnoreCase),
            $"Expected empty/invalid name error message, got: {errorText}");
    }

    /// <summary>
    /// Verifies that no-op rename (same name after trim) returns success=true.
    /// Contract: No-op → success=true (no error, no change needed).
    /// </summary>
    [Fact]
    public async Task PowerQueryRename_SameNameAfterTrim_ReturnsSuccessTrue()
    {
        // Arrange - create a query
        await CreatePowerQuery("TestQuery", "let x = 1 in x");

        // Act - rename to same name with extra whitespace
        var json = await CallToolAsync("powerquery", new Dictionary<string, object?>
        {
            ["action"] = "rename",
            ["session_id"] = _sessionId,
            ["old_name"] = "TestQuery",
            ["new_name"] = "  TestQuery  " // Same after trim = no-op
        });
        _output.WriteLine($"Response: {json}");

        // Assert - should return success=true (no-op is valid)
        var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        Assert.True(root.GetProperty("success").GetBoolean(), $"Expected success=true for no-op rename. Response: {json}");
    }

    /// <summary>
    /// Verifies successful rename returns proper RenameResult structure.
    /// Contract: Success → success=true with objectType, oldName, newName populated.
    /// </summary>
    [Fact]
    public async Task PowerQueryRename_Success_ReturnsCompleteRenameResult()
    {
        // Arrange - create a query
        await CreatePowerQuery("OriginalName", "let x = 1 in x");

        // Act - perform valid rename
        var json = await CallToolAsync("powerquery", new Dictionary<string, object?>
        {
            ["action"] = "rename",
            ["session_id"] = _sessionId,
            ["old_name"] = "OriginalName",
            ["new_name"] = "NewName"
        });
        _output.WriteLine($"Response: {json}");

        // Assert - verify complete RenameResult structure
        var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        Assert.True(root.GetProperty("success").GetBoolean(), $"Expected success=true. Response: {json}");
        Assert.Equal("power-query", root.GetProperty("objectType").GetString());
        Assert.Equal("OriginalName", root.GetProperty("oldName").GetString());
        Assert.Equal("NewName", root.GetProperty("newName").GetString());
    }

    #endregion

    #region Data Model Rename Contract Tests

    /// <summary>
    /// Verifies that renaming a non-existent table returns JSON with success=false (not exception).
    /// Contract: Missing object → success=false with "not found" message.
    /// </summary>
    [Fact]
    public async Task DataModelRenameTable_MissingTable_ReturnsJsonWithSuccessFalse()
    {
        // Arrange - no data model table exists

        // Act - attempt to rename non-existent table
        var json = await CallToolAsync("datamodel", new Dictionary<string, object?>
        {
            ["action"] = "rename-table",
            ["session_id"] = _sessionId,
            ["old_name"] = "NonExistentTable",
            ["new_name"] = "NewTableName"
        });
        _output.WriteLine($"Response: {json}");

        // Assert - should return JSON with success=false, NOT throw exception
        var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        Assert.False(root.GetProperty("success").GetBoolean(), "Expected success=false for missing table");
        Assert.True(root.TryGetProperty("errorMessage", out var errorMsg), "Expected errorMessage property");
        // When Data Model is empty, error message explains there are no tables
        Assert.Contains("no tables", errorMsg.GetString()!, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Verifies that attempting rename on an existing table returns JSON explaining Excel limitation.
    /// Contract: Excel limitation → success=false with clear explanation about immutable names.
    /// </summary>
    [Fact]
    public async Task DataModelRenameTable_ExcelLimitation_ReturnsJsonWithClearError()
    {
        // Arrange - create a Power Query and load it to the Data Model
        await CreatePowerQuery("TestData", "let Source = #table({\"Col1\", \"Col2\"}, {{\"A\", 1}, {\"B\", 2}}) in Source");
        await LoadQueryToDataModel("TestData");

        // Act - attempt to rename the table in Data Model
        var json = await CallToolAsync("datamodel", new Dictionary<string, object?>
        {
            ["action"] = "rename-table",
            ["session_id"] = _sessionId,
            ["old_name"] = "TestData",
            ["new_name"] = "NewTableName"
        });
        _output.WriteLine($"Response: {json}");

        // Assert - should return JSON with success=false and clear explanation
        var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        Assert.False(root.GetProperty("success").GetBoolean(), "Expected success=false for Excel limitation");
        Assert.True(root.TryGetProperty("errorMessage", out var errorMsg), "Expected errorMessage property");
        var errorText = errorMsg.GetString()!;

        // Error should explain why rename cannot proceed
        Assert.True(
            errorText.Contains("immutable", StringComparison.OrdinalIgnoreCase) ||
            errorText.Contains("cannot rename", StringComparison.OrdinalIgnoreCase) ||
            errorText.Contains("read-only", StringComparison.OrdinalIgnoreCase) ||
            errorText.Contains("not supported", StringComparison.OrdinalIgnoreCase) ||
            errorText.Contains("not found", StringComparison.OrdinalIgnoreCase),
            $"Expected clear explanation of why rename cannot proceed, got: {errorText}");
    }

    /// <summary>
    /// Verifies that empty new name returns JSON with success=false (not exception).
    /// Contract: Invalid name → success=false with validation message.
    /// </summary>
    [Fact]
    public async Task DataModelRenameTable_EmptyNewName_ReturnsJsonWithSuccessFalse()
    {
        // Arrange - create table in Data Model
        await CreatePowerQuery("DataTable", "let Source = #table({\"Col1\", \"Col2\"}, {{\"A\", 1}, {\"B\", 2}}) in Source");
        await LoadQueryToDataModel("DataTable");

        // Act - try to rename with empty new name
        var json = await CallToolAsync("datamodel", new Dictionary<string, object?>
        {
            ["action"] = "rename-table",
            ["session_id"] = _sessionId,
            ["old_name"] = "DataTable",
            ["new_name"] = "   " // Empty after trim
        });
        _output.WriteLine($"Response: {json}");

        // Assert - should return JSON with success=false, NOT throw exception
        var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        // Check for "success" property
        bool isFailed = root.TryGetProperty("success", out var successProp) && !successProp.GetBoolean();
        Assert.True(isFailed, "Expected success=false for empty name");
        Assert.True(root.TryGetProperty("errorMessage", out _), "Expected errorMessage property");
    }

    #endregion

    #region Helper Methods

    /// <summary>
    /// Calls a tool via the MCP protocol and returns the text response.
    /// </summary>
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

    private async Task CreatePowerQuery(string name, string mCode)
    {
        var json = await CallToolAsync("powerquery", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["session_id"] = _sessionId,
            ["query_name"] = name,
            ["m_code"] = mCode
        });

        var doc = JsonDocument.Parse(json);
        Assert.True(doc.RootElement.GetProperty("success").GetBoolean(),
            $"Failed to create query {name}: {json}");
    }

    private async Task LoadQueryToDataModel(string queryName)
    {
        var json = await CallToolAsync("powerquery", new Dictionary<string, object?>
        {
            ["action"] = "load-to",
            ["session_id"] = _sessionId,
            ["query_name"] = queryName,
            ["load_destination"] = "load-to-data-model"
        });

        var doc = JsonDocument.Parse(json);
        Assert.True(doc.RootElement.GetProperty("success").GetBoolean(),
            $"Failed to load query {queryName} to data model: {json}");
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
        // Close the session first to release Excel COM resources
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
                _output.WriteLine("✓ Session closed during cleanup");
            }
            catch (Exception ex)
            {
                _output.WriteLine($"Warning: Failed to close session: {ex.Message}");
            }
        }

        // Dispose client first - signals we're done sending requests
        if (_client != null)
        {
            await _client.DisposeAsync();
        }

        // Complete BOTH pipes to signal EOF for graceful server shutdown
        _clientToServerPipe.Writer.Complete();
        _serverToClientPipe.Writer.Complete();

        // Wait for server graceful shutdown with timeout
        if (_serverTask != null)
        {
            var shutdownTimeout = Task.Delay(TimeSpan.FromSeconds(10));
            var completed = await Task.WhenAny(_serverTask, shutdownTimeout);

            if (completed == shutdownTimeout)
            {
                // Server didn't shut down in time - cancel as fallback
                _output.WriteLine("Warning: Server did not shut down gracefully, forcing cancellation");
                await _cts.CancelAsync();
                try
                {
                    await _serverTask;
                }
                catch (OperationCanceledException)
                {
                    // Expected when we had to force cancel
                }
            }
        }

        // Reset test transport for next test class
        Program.ResetTestTransport();

        _cts.Dispose();

        // Clean up temp files
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
        catch (Exception ex)
        {
            _output.WriteLine($"Warning: Failed to cleanup temp directory: {ex.Message}");
        }
    }

    #endregion
}




