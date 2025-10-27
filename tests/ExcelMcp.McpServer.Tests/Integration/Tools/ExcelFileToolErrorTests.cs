using System.Text.Json;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Simple test to diagnose the excel_file tool issue
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
public class ExcelFileToolErrorTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;

    public ExcelFileToolErrorTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelFile_Error_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, true);
            }
        }
        catch
        {
            // Cleanup failed - not critical for test results
        }
        GC.SuppressFinalize(this);
    }

    [Fact]
    public async Task ExcelFile_CreateEmpty_ShouldWork()
    {
        // Arrange
        var testFile = Path.Combine(_tempDir, "test-file.xlsx");

        _output.WriteLine($"Testing file creation at: {testFile}");

        // Act - Call the tool directly
        var result = await ExcelFileTool.ExcelFile("create-empty", testFile);

        _output.WriteLine($"Tool result: {result}");

        // Parse the result
        var jsonDoc = JsonDocument.Parse(result);
        var success = jsonDoc.RootElement.GetProperty("success").GetBoolean();

        // Assert
        Assert.True(success, $"File creation failed: {result}");
        Assert.True(File.Exists(testFile), "File was not actually created");
    }

    [Fact]
    public async Task ExcelFile_WithInvalidAction_ShouldReturnError()
    {
        // Arrange
        var testFile = Path.Combine(_tempDir, "test-file.xlsx");

        // Act & Assert - Should throw McpException for invalid action
        var exception = await Assert.ThrowsAsync<ModelContextProtocol.McpException>(async () =>
            await ExcelFileTool.ExcelFile("invalid-action", testFile));

        _output.WriteLine($"Exception message for invalid action: {exception.Message}");

        // Assert - Verify exception contains expected message
        Assert.Contains("Unknown action 'invalid-action'", exception.Message);
    }
}
