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
/// Tests for ExcelFileTool operation tracking functionality.
/// Verifies that LIST action returns operation counts and that
/// CLOSE is blocked when operations are running.
/// </summary>
[Collection("ProgramTransport")]  // Uses Program.ConfigureTestTransport() - must run sequentially
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "SessionManager")]
[Trait("RequiresExcel", "true")]
public class ExcelFileToolOperationTrackingTests : IAsyncLifetime, IAsyncDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private readonly string _testExcelFile;

    // MCP transport pipes
    private readonly Pipe _clientToServerPipe = new();
    private readonly Pipe _serverToClientPipe = new();
    private readonly CancellationTokenSource _cts = new();
    private McpClient? _client;
    private Task? _serverTask;

    public ExcelFileToolOperationTrackingTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Join(Path.GetTempPath(), $"ExcelFileToolOpTrackingTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        _testExcelFile = Path.Join(_tempDir, "TestWorkbook.xlsx");
        _output.WriteLine($"Test directory: {_tempDir}");
    }

    public async Task InitializeAsync()
    {
        // Configure the server to use our test pipes
        Program.ConfigureTestTransport(_clientToServerPipe, _serverToClientPipe);

        // Run the real server
        _serverTask = Program.Main([]);

        // Allow server to initialize before client connection
        // SDK 0.5.0+ has stricter initialization timing
        await Task.Delay(100);

        // Create client connected to the server via pipes
        _client = await McpClient.CreateAsync(
            new StreamClientTransport(
                serverInput: _clientToServerPipe.Writer.AsStream(),
                serverOutput: _serverToClientPipe.Reader.AsStream()),
            clientOptions: new McpClientOptions
            {
                ClientInfo = new() { Name = "OpTrackingTestClient", Version = "1.0.0" },
                InitializationTimeout = TimeSpan.FromSeconds(30)  // Increase timeout for test stability
            },
            cancellationToken: _cts.Token);

        _output.WriteLine($"✓ Connected to server: {_client.ServerInfo?.Name} v{_client.ServerInfo?.Version}");

        // Create a test Excel file and open session
        var createResult = await CallToolAsync("excel_file", new Dictionary<string, object?>
        {
            ["action"] = "CreateAndOpen",
            ["path"] = _testExcelFile
        });
        _output.WriteLine($"Created test file: {_testExcelFile}");

        // Close the session immediately - we only need the file to exist for the tests
        if (createResult.TryGetProperty("sessionId", out var sessionIdProp))
        {
            var sessionId = sessionIdProp.GetString();
            await CallToolAsync("excel_file", new Dictionary<string, object?>
            {
                ["action"] = "Close",
                ["sessionId"] = sessionId,
                ["save"] = false
            });
        }
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

    private async Task DisposeAsyncCore()
    {
        // Cancel any pending operations
        await _cts.CancelAsync();

        // Close client
        if (_client != null)
        {
            await _client.DisposeAsync();
        }

        // Complete the pipes to signal server to stop
        await _clientToServerPipe.Writer.CompleteAsync();
        await _serverToClientPipe.Reader.CompleteAsync();

        // Wait for server to finish
        if (_serverTask != null)
        {
            try
            {
                await _serverTask.WaitAsync(TimeSpan.FromSeconds(10));
            }
            catch (OperationCanceledException)
            {
                // Expected
            }
            catch (TimeoutException)
            {
                _output.WriteLine("Warning: Server did not stop within timeout");
            }
        }

        // Cleanup pipes
        _clientToServerPipe.Writer.Complete();
        _clientToServerPipe.Reader.Complete();
        _serverToClientPipe.Writer.Complete();
        _serverToClientPipe.Reader.Complete();

        // Reset test transport to avoid contaminating other tests
        Program.ResetTestTransport();

        // Delete test files
        if (Directory.Exists(_tempDir))
        {
#pragma warning disable CA1031 // Catch general exception - best effort cleanup in test disposal
            try
            {
                Directory.Delete(_tempDir, recursive: true);
            }
            catch
            {
                // Best effort cleanup - test files will be cleaned by OS temp cleanup
            }
#pragma warning restore CA1031
        }
    }

    private async Task<JsonElement> CallToolAsync(string toolName, Dictionary<string, object?> args)
    {
        var result = await _client!.CallToolAsync(toolName, args, cancellationToken: _cts.Token);

        // Parse response - use pattern from McpServerSmokeTests
        var textBlock = result.Content.OfType<TextContentBlock>().FirstOrDefault();
        if (textBlock?.Text != null)
        {
            return JsonDocument.Parse(textBlock.Text).RootElement;
        }

        throw new InvalidOperationException($"Unexpected response from {toolName}");
    }

    #region List Action with Operation Tracking

    [Fact]
    public async Task List_ReturnsSessionsWithOperationStatus()
    {
        // Open a session
        var openResult = await CallToolAsync("excel_file", new Dictionary<string, object?>
        {
            ["action"] = "Open",
            ["path"] = _testExcelFile,
            ["show"] = false
        });

        Assert.True(openResult.GetProperty("success").GetBoolean());
        var sessionId = openResult.GetProperty("sessionId").GetString();
        Assert.NotNull(sessionId);

        try
        {
            // List sessions
            var listResult = await CallToolAsync("excel_file", new Dictionary<string, object?>
            {
                ["action"] = "List"
            });

            Assert.True(listResult.GetProperty("success").GetBoolean());

            var sessions = listResult.GetProperty("sessions");
            Assert.Equal(1, sessions.GetArrayLength());

            var session = sessions[0];
            Assert.Equal(sessionId, session.GetProperty("sessionId").GetString());
            Assert.True(session.TryGetProperty("activeOperations", out var activeOps));
            Assert.Equal(0, activeOps.GetInt32());
            Assert.True(session.TryGetProperty("canClose", out var canClose));
            Assert.True(canClose.GetBoolean());
            Assert.True(session.TryGetProperty("isExcelVisible", out var isVisible));
            Assert.False(isVisible.GetBoolean());
        }
        finally
        {
            // Cleanup
            await CallToolAsync("excel_file", new Dictionary<string, object?>
            {
                ["action"] = "Close",
                ["sid"] = sessionId,
                ["save"] = false
            });
        }
    }

    [Fact]
    public async Task List_SessionWithShowExcelTrue_ReturnsIsExcelVisibleTrue()
    {
        // Open a session with show=true
        var openResult = await CallToolAsync("excel_file", new Dictionary<string, object?>
        {
            ["action"] = "Open",
            ["path"] = _testExcelFile,
            ["show"] = true
        });

        Assert.True(openResult.GetProperty("success").GetBoolean());
        var sessionId = openResult.GetProperty("sessionId").GetString();

        try
        {
            // List sessions
            var listResult = await CallToolAsync("excel_file", new Dictionary<string, object?>
            {
                ["action"] = "List"
            });

            var sessions = listResult.GetProperty("sessions");
            var session = sessions[0];
            Assert.True(session.GetProperty("isExcelVisible").GetBoolean());
        }
        finally
        {
            await CallToolAsync("excel_file", new Dictionary<string, object?>
            {
                ["action"] = "Close",
                ["sid"] = sessionId,
                ["save"] = false
            });
        }
    }

    #endregion

    #region Close Blocking (Note: Hard to test without simulating parallel calls)

    [Fact]
    public async Task Close_NoOperationsRunning_ClosesSuccessfully()
    {
        // Open a session
        var openResult = await CallToolAsync("excel_file", new Dictionary<string, object?>
        {
            ["action"] = "Open",
            ["path"] = _testExcelFile,
            ["show"] = false
        });

        var sessionId = openResult.GetProperty("sessionId").GetString();

        // Close should succeed (no operations running)
        var closeResult = await CallToolAsync("excel_file", new Dictionary<string, object?>
        {
            ["action"] = "Close",
            ["sid"] = sessionId,
            ["save"] = false
        });

        Assert.True(closeResult.GetProperty("success").GetBoolean());

        // Verify session is gone
        var listResult = await CallToolAsync("excel_file", new Dictionary<string, object?>
        {
            ["action"] = "List"
        });

        Assert.Equal(0, listResult.GetProperty("count").GetInt32());
    }

    [Fact]
    public async Task Close_NonExistentSession_ReturnsError()
    {
        var closeResult = await CallToolAsync("excel_file", new Dictionary<string, object?>
        {
            ["action"] = "Close",
            ["sid"] = "nonexistent-session-id",
            ["save"] = false
        });

        Assert.False(closeResult.GetProperty("success").GetBoolean());
        Assert.True(closeResult.TryGetProperty("errorMessage", out var errorMsg));
        Assert.Contains("not found", errorMsg.GetString());
    }

    #endregion
}
