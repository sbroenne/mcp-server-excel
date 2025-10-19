using Xunit;
using Sbroenne.ExcelMcp.McpServer.Tools;
using System.IO;
using System.Text.Json;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Tools;

/// <summary>
/// Integration tests for ExcelCLI MCP Server using official MCP SDK
/// These tests validate the 6 resource-based tools for AI assistants
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "MCP")]
public class ExcelMcpServerTests : IDisposable
{
    private readonly string _testExcelFile;
    private readonly string _tempDir;

    public ExcelMcpServerTests()
    {
        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCLI_MCP_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        
        _testExcelFile = Path.Combine(_tempDir, "MCPTestWorkbook.xlsx");
    }

    public void Dispose()
    {
        // Cleanup test files
        if (Directory.Exists(_tempDir))
        {
            try
            {
                Directory.Delete(_tempDir, recursive: true);
            }
            catch
            {
                // Ignore cleanup errors in tests
            }
        }
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void ExcelFile_CreateEmpty_ShouldReturnSuccessJson()
    {
        // Act
        var result = ExcelTools.ExcelFile("create-empty", _testExcelFile);

        // Assert
        Assert.NotNull(result);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
        Assert.True(File.Exists(_testExcelFile));
    }

    [Fact]
    public void ExcelFile_ValidateExistingFile_ShouldReturnValidTrue()
    {
        // Arrange - Create a file first
        ExcelTools.ExcelFile("create-empty", _testExcelFile);

        // Act
        var result = ExcelTools.ExcelFile("validate", _testExcelFile);

        // Assert
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("valid").GetBoolean());
    }

    [Fact]
    public void ExcelFile_ValidateNonExistentFile_ShouldReturnValidFalse()
    {
        // Act
        var result = ExcelTools.ExcelFile("validate", "nonexistent.xlsx");

        // Assert
        var json = JsonDocument.Parse(result);
        Assert.False(json.RootElement.GetProperty("valid").GetBoolean());
        Assert.Equal("File does not exist", json.RootElement.GetProperty("error").GetString());
    }

    [Fact]
    public void ExcelFile_CheckExists_ShouldReturnExistsStatus()
    {
        // Act - Test non-existent file
        var result1 = ExcelTools.ExcelFile("check-exists", _testExcelFile);
        var json1 = JsonDocument.Parse(result1);
        Assert.False(json1.RootElement.GetProperty("exists").GetBoolean());

        // Create file and test again
        ExcelTools.ExcelFile("create-empty", _testExcelFile);
        var result2 = ExcelTools.ExcelFile("check-exists", _testExcelFile);
        var json2 = JsonDocument.Parse(result2);
        Assert.True(json2.RootElement.GetProperty("exists").GetBoolean());
    }

    [Fact]
    public void ExcelFile_UnknownAction_ShouldReturnError()
    {
        // Act
        var result = ExcelTools.ExcelFile("unknown", _testExcelFile);

        // Assert
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("error", out _));
    }

    [Fact]
    public void ExcelWorksheet_List_ShouldReturnSuccessAfterCreation()
    {
        // Arrange
        ExcelTools.ExcelFile("create-empty", _testExcelFile);

        // Act
        var result = ExcelTools.ExcelWorksheet("list", _testExcelFile);

        // Assert
        var json = JsonDocument.Parse(result);
        // Should succeed (return success: true) when file exists
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
    }

    [Fact]
    public void ExcelWorksheet_NonExistentFile_ShouldReturnError()
    {
        // Act
        var result = ExcelTools.ExcelWorksheet("list", "nonexistent.xlsx");

        // Assert
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("error", out _));
    }

    [Fact]
    public void ExcelParameter_List_ShouldReturnSuccessAfterCreation()
    {
        // Arrange
        ExcelTools.ExcelFile("create-empty", _testExcelFile);

        // Act
        var result = ExcelTools.ExcelParameter("list", _testExcelFile);

        // Assert
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
    }

    [Fact]
    public void ExcelCell_GetValue_RequiresExistingFile()
    {
        // Act - Try to get cell value from non-existent file
        var result = ExcelTools.ExcelCell("get-value", "nonexistent.xlsx", "Sheet1", "A1");

        // Assert
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("error", out _));
    }
}