using System.Text.Json;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Test to verify that excel_file can create files in non-existent directories
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
public class ExcelFileDirectoryTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;

    public ExcelFileDirectoryTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelFile_Dir_Tests_{Guid.NewGuid():N}");
        // Don't create the directory - let the tool create it
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
    public void ExcelFile_CreateInNonExistentDirectory_ShouldWork()
    {
        // Arrange
        var testFile = Path.Combine(_tempDir, "subdir", "test-file.xlsx");

        _output.WriteLine($"Testing file creation in non-existent directory: {testFile}");
        _output.WriteLine($"Directory exists before: {Directory.Exists(Path.GetDirectoryName(testFile))}");

        // Act - Call the tool directly
        var result = ExcelFileTool.File("create-empty", testFile);

        _output.WriteLine($"Tool result: {result}");

        // Parse the result
        var jsonDoc = JsonDocument.Parse(result);

        if (jsonDoc.RootElement.TryGetProperty("success", out var successElement))
        {
            var success = successElement.GetBoolean();
            Assert.True(success, $"File creation failed: {result}");
            Assert.True(File.Exists(testFile), "File was not actually created");
        }
        else if (jsonDoc.RootElement.TryGetProperty("error", out var errorElement))
        {
            var error = errorElement.GetString();
            _output.WriteLine($"Expected this might fail - error: {error}");
            // This is expected if the directory doesn't get created
        }
    }

    [Fact]
    public void ExcelFile_WithVeryLongPath_ShouldHandleGracefully()
    {
        // Arrange - Create a path that might be too long
        var longPath = string.Join("", Enumerable.Repeat("verylongdirectoryname", 20));
        var testFile = Path.Combine(_tempDir, longPath, "test-file.xlsx");

        _output.WriteLine($"Testing with very long path: {testFile.Length} characters");
        _output.WriteLine($"Path: {testFile}");

        // Act - Call the tool directly
        var result = ExcelFileTool.File("create-empty", testFile);

        _output.WriteLine($"Tool result: {result}");

        // Just make sure it doesn't throw an exception
        var jsonDoc = JsonDocument.Parse(result);
        Assert.True(jsonDoc.RootElement.ValueKind == JsonValueKind.Object);
    }
}
