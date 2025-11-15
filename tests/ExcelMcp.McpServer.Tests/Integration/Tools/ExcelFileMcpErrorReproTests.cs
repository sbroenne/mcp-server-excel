using System.Text.Json;
using Sbroenne.ExcelMcp.McpServer.Models;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Test to reproduce the exact MCP error scenario
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
public class ExcelFileMcpErrorReproTests
{
    private readonly ITestOutputHelper _output;
    /// <inheritdoc/>

    public ExcelFileMcpErrorReproTests(ITestOutputHelper output)
    {
        _output = output;
    }
    /// <inheritdoc/>

    [Fact]
    public async Task ExcelFile_ExactMcpTestScenario_ShouldWork()
    {
        // Arrange - Use exact path pattern from failing test
        var tempDir = Path.Join(Path.GetTempPath(), $"MCPClient_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDir);
        var testFile = Path.Join(tempDir, "roundtrip-test.xlsx");

        try
        {
            _output.WriteLine($"Testing exact MCP scenario:");
            _output.WriteLine($"Action: create-empty");
            _output.WriteLine($"ExcelPath: {testFile}");
            _output.WriteLine($"Directory exists: {Directory.Exists(tempDir)}");

            // Act - Call the tool with exact parameters from MCP test
            var result = ExcelFileTool.ExcelFile(FileAction.CreateEmpty, testFile);

            _output.WriteLine($"Tool result: {result}");

            // Parse the result to understand format
            var jsonDoc = JsonDocument.Parse(result);
            _output.WriteLine($"JSON structure: {jsonDoc.RootElement}");

            if (jsonDoc.RootElement.TryGetProperty("success", out var successElement))
            {
                var success = successElement.GetBoolean();
                if (success)
                {
                    _output.WriteLine("✅ SUCCESS: File creation worked");
                    Assert.True(File.Exists(testFile), "File should exist");
                }
                else
                {
                    _output.WriteLine("❌ FAILED: Tool returned success=false");
                    if (jsonDoc.RootElement.TryGetProperty("error", out var errorElement))
                    {
                        _output.WriteLine($"Error details: {errorElement.GetString()}");
                    }
                    Assert.Fail($"Tool returned failure: {result}");
                }
            }
            else if (jsonDoc.RootElement.TryGetProperty("error", out var errorElement))
            {
                var error = errorElement.GetString();
                _output.WriteLine($"❌ ERROR: {error}");
                Assert.Fail($"Tool returned error: {error}");
            }
            else
            {
                _output.WriteLine($"⚠️ UNKNOWN: Unexpected JSON format");
                Assert.Fail($"Unexpected response format: {result}");
            }
        }
        finally
        {
            // Cleanup
            try
            {
                if (Directory.Exists(tempDir))
                {
                    Directory.Delete(tempDir, true);
                }
            }
            catch { }
        }
    }
}

