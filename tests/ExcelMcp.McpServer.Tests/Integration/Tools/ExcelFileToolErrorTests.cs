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

    [Fact]
    public async Task ExcelFile_TestAction_WithExistingFile_ShouldReturnSuccess()
    {
        // Arrange
        var testFile = Path.Combine(_tempDir, "test-validation.xlsx");

        // Create a dummy file (test action doesn't need a real Excel file, just checks existence and extension)
        File.WriteAllText(testFile, "dummy Excel content");

        _output.WriteLine($"Testing file validation at: {testFile}");

        // Act - Call the test action
        var result = await ExcelFileTool.ExcelFile("test", testFile);

        _output.WriteLine($"Test result: {result}");

        // Parse the result
        var jsonDoc = JsonDocument.Parse(result);
        var success = jsonDoc.RootElement.GetProperty("success").GetBoolean();
        var exists = jsonDoc.RootElement.GetProperty("exists").GetBoolean();
        var isValid = jsonDoc.RootElement.GetProperty("isValid").GetBoolean();
        var extension = jsonDoc.RootElement.GetProperty("extension").GetString();

        // Assert
        Assert.True(success, $"Test action failed: {result}");
        Assert.True(exists, "File should exist");
        Assert.True(isValid, "File should be valid");
        Assert.Equal(".xlsx", extension);
    }

    [Fact]
    public async Task ExcelFile_TestAction_WithNonExistentFile_ShouldReturnFailure()
    {
        // Arrange
        var testFile = Path.Combine(_tempDir, "nonexistent.xlsx");

        _output.WriteLine($"Testing non-existent file at: {testFile}");

        // Act - Call the test action on non-existent file
        var result = await ExcelFileTool.ExcelFile("test", testFile);

        _output.WriteLine($"Test result: {result}");

        // Parse the result
        var jsonDoc = JsonDocument.Parse(result);
        var success = jsonDoc.RootElement.GetProperty("success").GetBoolean();
        var exists = jsonDoc.RootElement.GetProperty("exists").GetBoolean();
        var isValid = jsonDoc.RootElement.GetProperty("isValid").GetBoolean();

        // Assert
        Assert.False(success, "Test action should fail for non-existent file");
        Assert.False(exists, "File should not exist");
        Assert.False(isValid, "File should not be valid");
    }

    [Fact]
    public async Task ExcelFile_TestAction_WithInvalidExtension_ShouldReturnFailure()
    {
        // Arrange
        var testFile = Path.Combine(_tempDir, "test-file.txt");

        // Create file with invalid extension
        File.WriteAllText(testFile, "test content");

        _output.WriteLine($"Testing invalid extension at: {testFile}");

        // Act - Call the test action
        var result = await ExcelFileTool.ExcelFile("test", testFile);

        _output.WriteLine($"Test result: {result}");

        // Parse the result
        var jsonDoc = JsonDocument.Parse(result);
        var success = jsonDoc.RootElement.GetProperty("success").GetBoolean();
        var exists = jsonDoc.RootElement.GetProperty("exists").GetBoolean();
        var isValid = jsonDoc.RootElement.GetProperty("isValid").GetBoolean();
        var extension = jsonDoc.RootElement.GetProperty("extension").GetString();

        // Assert
        Assert.False(success, "Test action should fail for invalid extension");
        Assert.True(exists, "File should exist");
        Assert.False(isValid, "File should not be valid");
        Assert.Equal(".txt", extension);
    }
}
