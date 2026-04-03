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
/// 2. Alias-only rename payloads fail with clear canonical parameter errors
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
        (_client, _serverTask) = await ProgramTransportTestHost.StartAsync(
            _clientToServerPipe,
            _serverToClientPipe,
            _cts.Token,
            "WsRenameRegressionClient");

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

    #region Canonical parameter contract

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
    /// Verifies that alias-only rename payloads fail and point callers at the canonical old_name/new_name contract.
    /// </summary>
    [Theory]
    [MemberData(nameof(AliasOnlyRenamePayloads))]
    public async Task Rename_WithAliasOnlyPayload_FailsWithCanonicalParameterError(
        string originalSheetName,
        string renamedSheetName,
        string expectedErrorFragment,
        Dictionary<string, object?> aliasArguments)
    {
        // Arrange
        await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["session_id"] = _sessionId,
            ["sheet_name"] = originalSheetName
        });

        aliasArguments["action"] = "rename";
        aliasArguments["session_id"] = _sessionId;

        // Act
        var json = await CallToolAsync("worksheet", aliasArguments);
        _output.WriteLine($"Response: {json}");

        // Assert
        var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        Assert.False(root.GetProperty("success").GetBoolean(),
            $"Alias-only rename payload should fail. Response: {json}");
        Assert.Contains(expectedErrorFragment, root.GetProperty("errorMessage").GetString(), StringComparison.Ordinal);

        var listJson = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = _sessionId
        });
        Assert.Contains(originalSheetName, listJson);
        Assert.DoesNotContain(renamedSheetName, listJson);
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
        Assert.Contains("old_name is required for rename action", root.GetProperty("errorMessage").GetString(), StringComparison.Ordinal);
    }

    #endregion

    #region Helper Methods

    public static IEnumerable<object[]> AliasOnlyRenamePayloads()
    {
        yield return
        [
            "LegacySheet",
            "RenamedLegacy",
            "old_name is required for rename action",
            new Dictionary<string, object?>
            {
                ["sheet_name"] = "LegacySheet",
                ["target_name"] = "RenamedLegacy"
            }
        ];

        yield return
        [
            "TargetSheetAlias",
            "RenamedViaTargetSheet",
            "old_name is required for rename action",
            new Dictionary<string, object?>
            {
                ["sheet_name"] = "TargetSheetAlias",
                ["target_sheet_name"] = "RenamedViaTargetSheet"
            }
        ];

        yield return
        [
            "SourceNameAlias",
            "RenamedViaSourceName",
            "old_name is required for rename action",
            new Dictionary<string, object?>
            {
                ["source_name"] = "SourceNameAlias",
                ["target_name"] = "RenamedViaSourceName"
            }
        ];

        yield return
        [
            "SourceSheetAlias",
            "RenamedViaSourceSheet",
            "old_name is required for rename action",
            new Dictionary<string, object?>
            {
                ["source_sheet"] = "SourceSheetAlias",
                ["target_sheet_name"] = "RenamedViaSourceSheet"
            }
        ];

        yield return
        [
            "MixedCanonicalAlias",
            "MixedAliasTarget",
            "new_name is required for rename action",
            new Dictionary<string, object?>
            {
                ["old_name"] = "MixedCanonicalAlias",
                ["target_name"] = "MixedAliasTarget"
            }
        ];

        yield return
        [
            "MixedAliasCanonical",
            "MixedAliasCanonicalRenamed",
            "old_name is required for rename action",
            new Dictionary<string, object?>
            {
                ["sheet_name"] = "MixedAliasCanonical",
                ["new_name"] = "MixedAliasCanonicalRenamed"
            }
        ];
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

        await ProgramTransportTestHost.StopAsync(
            _client,
            _clientToServerPipe,
            _serverToClientPipe,
            _serverTask,
            _output);
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
