// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.IO.Pipelines;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// End-to-end regressions for file tool behavior through the MCP protocol.
/// These tests use the real transport and server pipeline instead of calling tool methods directly.
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "File")]
[Trait("RequiresExcel", "true")]
#pragma warning disable CA1001 // _cts is disposed in IAsyncLifetime.DisposeAsync
public sealed class ExcelFileToolProtocolRegressionTests : IAsyncLifetime
#pragma warning restore CA1001
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private readonly Pipe _clientToServerPipe = new();
    private readonly Pipe _serverToClientPipe = new();
    private readonly CancellationTokenSource _cts = new();
    private McpClient? _client;
    private Task? _serverTask;

    public ExcelFileToolProtocolRegressionTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Join(Path.GetTempPath(), $"ExcelFileToolProtocolRegressionTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
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
                ClientInfo = new() { Name = "ExcelFileToolProtocolRegressionClient", Version = "1.0.0" },
                InitializationTimeout = TimeSpan.FromSeconds(30)
            },
            cancellationToken: _cts.Token);
    }

    public async Task DisposeAsync()
    {
        await DisposeAsyncCore();
        _cts.Dispose();
    }

    [Fact]
    public async Task FileOpen_FileLockedByAnotherProcess_ReturnsActionableError_AndNextOpenSucceeds()
    {
        var lockedFile = Path.Join(_tempDir, $"LockedOpen_{Guid.NewGuid():N}.xlsx");
        ExcelSession.CreateNew<bool>(lockedFile, false, (ctx, ct) => true);

        using (var fileLock = new FileStream(lockedFile, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
        {
            var lockedResult = await CallToolAsync("file", new Dictionary<string, object?>
            {
                ["action"] = "open",
                ["path"] = lockedFile
            });

            _output.WriteLine($"Locked file open result: {lockedResult}");

            using var lockedJson = JsonDocument.Parse(lockedResult);
            Assert.False(lockedJson.RootElement.GetProperty("success").GetBoolean());

            var errorMessage = lockedJson.RootElement.GetProperty("errorMessage").GetString();
            Assert.NotNull(errorMessage);
            Assert.Contains("already open", errorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("close the file", errorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("exclusive access", errorMessage, StringComparison.OrdinalIgnoreCase);
        }

        var listAfterFailure = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "list"
        });

        using (var listAfterFailureJson = JsonDocument.Parse(listAfterFailure))
        {
            Assert.True(listAfterFailureJson.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal(0, listAfterFailureJson.RootElement.GetProperty("sessions").GetArrayLength());
        }

        var openAfterRelease = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "open",
            ["path"] = lockedFile
        });
        AssertSuccess(openAfterRelease, "Open workbook after lock release");

        var sessionId = GetJsonProperty(openAfterRelease, "session_id");
        Assert.False(string.IsNullOrWhiteSpace(sessionId));

        await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "close",
            ["session_id"] = sessionId,
            ["save"] = false
        });
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
                _output.WriteLine("Warning: Server did not stop within timeout");
            }
        }

        Program.ResetTestTransport();

        if (Directory.Exists(_tempDir))
        {
#pragma warning disable CA1031
            try
            {
                Directory.Delete(_tempDir, recursive: true);
            }
            catch
            {
            }
#pragma warning restore CA1031
        }
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

    private static void AssertSuccess(string jsonResult, string operationName)
    {
        Assert.NotNull(jsonResult);

        try
        {
            var json = JsonDocument.Parse(jsonResult);

            if (json.RootElement.TryGetProperty("error", out var error))
            {
                Assert.Fail($"{operationName} failed with error: {error.GetString()}");
            }

            if (json.RootElement.TryGetProperty("Success", out var successPascal))
            {
                if (!successPascal.GetBoolean())
                {
                    var errorMsg = json.RootElement.TryGetProperty("ErrorMessage", out var errProp)
                        ? errProp.GetString()
                        : "Unknown error";
                    Assert.Fail($"{operationName} returned Success=false: {errorMsg}");
                }
            }
            else if (json.RootElement.TryGetProperty("success", out var successCamel))
            {
                if (!successCamel.GetBoolean())
                {
                    var errorMsg = json.RootElement.TryGetProperty("errorMessage", out var errProp)
                        ? errProp.GetString()
                        : "Unknown error";
                    Assert.Fail($"{operationName} returned success=false: {errorMsg}");
                }
            }
        }
        catch (JsonException ex)
        {
            Assert.Fail($"{operationName} returned invalid JSON: {ex.Message}\nResponse: {jsonResult}");
        }
    }

    private static string? GetJsonProperty(string jsonResult, string propertyName)
    {
        var json = JsonDocument.Parse(jsonResult);
        return json.RootElement.TryGetProperty(propertyName, out var prop) ? prop.GetString() : null;
    }
}