using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Integration tests for the batch CLI command.
/// Tests the full CLI pipeline: process launch → batch parsing → daemon dispatch → NDJSON output.
/// Uses diag commands (ping, echo) to validate batch infrastructure without requiring Excel.
/// </summary>
[Collection("Service")]
[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("Feature", "Batch")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class BatchCommandTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;

    public BatchCommandTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Combine(Path.GetTempPath(), $"BatchTest_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        try { Directory.Delete(_tempDir, recursive: true); }
        catch { /* best effort */ }
    }

    // ============================================
    // SINGLE COMMAND - Basic execution
    // ============================================

    [Fact]
    public async Task Batch_SinglePingCommand_ReturnsSuccess()
    {
        var inputFile = WriteBatchFile("""
            [
              {"command": "diag.ping"}
            ]
            """);

        var results = await RunBatchAsync($"batch --input \"{inputFile}\"");

        Assert.Single(results);
        Assert.Equal(0, results[0].GetProperty("index").GetInt32());
        Assert.Equal("diag.ping", results[0].GetProperty("command").GetString());
        Assert.True(results[0].GetProperty("success").GetBoolean());
    }

    // ============================================
    // MULTIPLE COMMANDS - Sequential execution
    // ============================================

    [Fact]
    public async Task Batch_MultipleCommands_ExecutesAllSequentially()
    {
        var inputFile = WriteBatchFile("""
            [
              {"command": "diag.ping"},
              {"command": "diag.echo", "args": {"message": "hello"}},
              {"command": "diag.echo", "args": {"message": "world"}}
            ]
            """);

        var results = await RunBatchAsync($"batch --input \"{inputFile}\"");

        Assert.Equal(3, results.Count);

        // All should succeed
        Assert.All(results, r => Assert.True(r.GetProperty("success").GetBoolean()));

        // Indices should be sequential
        Assert.Equal(0, results[0].GetProperty("index").GetInt32());
        Assert.Equal(1, results[1].GetProperty("index").GetInt32());
        Assert.Equal(2, results[2].GetProperty("index").GetInt32());

        // Verify echo results
        Assert.Equal("hello", results[1].GetProperty("result").GetProperty("message").GetString());
        Assert.Equal("world", results[2].GetProperty("result").GetProperty("message").GetString());
    }

    // ============================================
    // ERROR HANDLING - Continue on error (default)
    // ============================================

    [Fact]
    public async Task Batch_CommandFails_ContinuesRemainingByDefault()
    {
        var inputFile = WriteBatchFile("""
            [
              {"command": "diag.ping"},
              {"command": "nonexistent.action"},
              {"command": "diag.ping"}
            ]
            """);

        var results = await RunBatchAsync($"batch --input \"{inputFile}\"");

        Assert.Equal(3, results.Count);
        Assert.True(results[0].GetProperty("success").GetBoolean());
        Assert.False(results[1].GetProperty("success").GetBoolean()); // invalid command fails
        Assert.True(results[2].GetProperty("success").GetBoolean()); // should still execute
    }

    // ============================================
    // ERROR HANDLING - Stop on error
    // ============================================

    [Fact]
    public async Task Batch_StopOnError_StopsAfterFirstFailure()
    {
        var inputFile = WriteBatchFile("""
            [
              {"command": "diag.ping"},
              {"command": "nonexistent.action"},
              {"command": "diag.ping"}
            ]
            """);

        var results = await RunBatchAsync($"batch --input \"{inputFile}\" --stop-on-error");

        Assert.Equal(2, results.Count); // Only 2 results — stopped after failure
        Assert.True(results[0].GetProperty("success").GetBoolean());
        Assert.False(results[1].GetProperty("success").GetBoolean());
    }

    // ============================================
    // INPUT VALIDATION
    // ============================================

    [Fact]
    public async Task Batch_EmptyArray_ReturnsError()
    {
        var inputFile = WriteBatchFile("[]");

        var result = await CliProcessHelper.RunAsync($"batch --input \"{inputFile}\"");
        _output.WriteLine($"Exit: {result.ExitCode}, Stderr: {result.Stderr}");

        Assert.Equal(1, result.ExitCode);
    }

    [Fact]
    public async Task Batch_MissingCommandField_ReturnsError()
    {
        var inputFile = WriteBatchFile("""
            [
              {"args": {"message": "no command field"}}
            ]
            """);

        var result = await CliProcessHelper.RunAsync($"batch --input \"{inputFile}\"");
        _output.WriteLine($"Exit: {result.ExitCode}, Stdout: {result.Stdout}");

        Assert.Equal(1, result.ExitCode);
    }

    [Fact]
    public async Task Batch_FileNotFound_ReturnsError()
    {
        var result = await CliProcessHelper.RunAsync("batch --input \"nonexistent.json\"");
        _output.WriteLine($"Exit: {result.ExitCode}, Stderr: {result.Stderr}");

        Assert.Equal(1, result.ExitCode);
    }

    // ============================================
    // NDJSON FORMAT - Output structure
    // ============================================

    [Fact]
    public async Task Batch_OutputIsValidNdjson_OneLinePerCommand()
    {
        var inputFile = WriteBatchFile("""
            [
              {"command": "diag.ping"},
              {"command": "diag.echo", "args": {"message": "test"}}
            ]
            """);

        var result = await CliProcessHelper.RunAsync($"batch --input \"{inputFile}\"");
        _output.WriteLine(result.Stdout);

        Assert.Equal(0, result.ExitCode);

        // Each non-empty line should be valid JSON
        var lines = result.Stdout.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(2, lines.Length);

        foreach (var line in lines)
        {
            var json = JsonDocument.Parse(line.Trim());
            Assert.True(json.RootElement.TryGetProperty("index", out _));
            Assert.True(json.RootElement.TryGetProperty("command", out _));
            Assert.True(json.RootElement.TryGetProperty("success", out _));
        }
    }

    [Fact]
    public async Task Batch_SuccessResult_HasNoErrorField()
    {
        var inputFile = WriteBatchFile("""
            [
              {"command": "diag.ping"}
            ]
            """);

        var results = await RunBatchAsync($"batch --input \"{inputFile}\"");

        Assert.Single(results);
        Assert.True(results[0].GetProperty("success").GetBoolean());
        // error field should be absent (JsonIgnoreCondition.WhenWritingNull)
        Assert.False(results[0].TryGetProperty("error", out _),
            "Success result should not contain 'error' field");
    }

    [Fact]
    public async Task Batch_ErrorResult_HasErrorField()
    {
        var inputFile = WriteBatchFile("""
            [
              {"command": "nonexistent.action"}
            ]
            """);

        var results = await RunBatchAsync($"batch --input \"{inputFile}\"");

        Assert.Single(results);
        Assert.False(results[0].GetProperty("success").GetBoolean());
        Assert.True(results[0].TryGetProperty("error", out var error));
        Assert.False(string.IsNullOrWhiteSpace(error.GetString()));
    }

    // ============================================
    // NDJSON INPUT FORMAT
    // ============================================

    [Fact]
    public async Task Batch_NdjsonInputFormat_ParsesCorrectly()
    {
        // NDJSON: one command per line (no array wrapper)
        var inputFile = WriteBatchFile(
            """{"command": "diag.ping"}""" + "\n" +
            """{"command": "diag.echo", "args": {"message": "ndjson"}}""");

        var results = await RunBatchAsync($"batch --input \"{inputFile}\"");

        Assert.Equal(2, results.Count);
        Assert.True(results[0].GetProperty("success").GetBoolean());
        Assert.Equal("ndjson", results[1].GetProperty("result").GetProperty("message").GetString());
    }

    // ============================================
    // EXIT CODE
    // ============================================

    [Fact]
    public async Task Batch_AllSucceed_ExitCodeZero()
    {
        var inputFile = WriteBatchFile("""
            [
              {"command": "diag.ping"},
              {"command": "diag.echo", "args": {"message": "test"}}
            ]
            """);

        var result = await CliProcessHelper.RunAsync($"batch --input \"{inputFile}\"");
        Assert.Equal(0, result.ExitCode);
    }

    [Fact]
    public async Task Batch_AnyFails_ExitCodeOne()
    {
        var inputFile = WriteBatchFile("""
            [
              {"command": "diag.ping"},
              {"command": "nonexistent.action"}
            ]
            """);

        var result = await CliProcessHelper.RunAsync($"batch --input \"{inputFile}\"");
        Assert.Equal(1, result.ExitCode);
    }

    // ============================================
    // Helpers
    // ============================================

    private string WriteBatchFile(string content)
    {
        var path = Path.Combine(_tempDir, $"batch_{Guid.NewGuid():N}.json");
        File.WriteAllText(path, content);
        return path;
    }

    /// <summary>
    /// Runs a batch command and parses the NDJSON output into a list of JsonElements.
    /// </summary>
    private async Task<List<JsonElement>> RunBatchAsync(string args, int timeoutMs = 30000)
    {
        var result = await CliProcessHelper.RunAsync(args, timeoutMs);
        _output.WriteLine($"Exit: {result.ExitCode}");
        _output.WriteLine($"Stdout: {result.Stdout}");
        if (!string.IsNullOrEmpty(result.Stderr))
            _output.WriteLine($"Stderr: {result.Stderr}");

        var results = new List<JsonElement>();
        foreach (var line in result.Stdout.Split('\n', StringSplitOptions.RemoveEmptyEntries))
        {
            var trimmed = line.Trim();
            if (string.IsNullOrEmpty(trimmed)) continue;

            try
            {
                using var doc = JsonDocument.Parse(trimmed);
                results.Add(doc.RootElement.Clone());
            }
            catch (JsonException)
            {
                _output.WriteLine($"Non-JSON line: {trimmed}");
            }
        }

        return results;
    }
}
