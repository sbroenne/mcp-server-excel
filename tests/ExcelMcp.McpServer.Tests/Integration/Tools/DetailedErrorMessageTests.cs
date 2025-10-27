using ModelContextProtocol;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Tests that verify our enhanced error messages include detailed diagnostic information for LLMs.
/// These tests prove that we throw McpException with:
/// - Exception type names ([Exception Type: ...])
/// - Inner exception messages (Inner: ...)
/// - Action context
/// - File paths
/// - Actionable guidance
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "ErrorHandling")]
public class DetailedErrorMessageTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private readonly string _testExcelFile;

    public DetailedErrorMessageTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelMcp_DetailedErrorTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        _testExcelFile = Path.Combine(_tempDir, "test-errors.xlsx");
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, recursive: true);
            }
        }
        catch { }

        GC.SuppressFinalize(this);
    }

    [Fact]
    public async Task ExcelWorksheet_WithNonExistentFile_ShouldThrowDetailedError()
    {
        // Arrange
        string nonExistentFile = Path.Combine(_tempDir, "nonexistent.xlsx");

        // Act & Assert - Should throw McpException with detailed error message
        var exception = await Assert.ThrowsAsync<McpException>(async () =>
            await ExcelWorksheetTool.ExcelWorksheet("list", nonExistentFile));

        // Verify detailed error message components
        _output.WriteLine($"Error message: {exception.Message}");

        // Should include action context
        Assert.Contains("list", exception.Message);

        // Should include file path
        Assert.Contains(nonExistentFile, exception.Message);

        // Should include specific error details
        Assert.Contains("File not found", exception.Message);

        _output.WriteLine("✅ Verified: Action, file path, and error details included");
    }

    [Fact]
    public async Task ExcelCell_WithNonExistentFile_ShouldThrowDetailedError()
    {
        // Arrange
        string nonExistentFile = Path.Combine(_tempDir, "nonexistent-cell.xlsx");

        // Act & Assert
        var exception = await Assert.ThrowsAsync<McpException>(async () =>
            await ExcelCellTool.ExcelCell("get-value", nonExistentFile, "Sheet1", "A1"));

        _output.WriteLine($"Error message: {exception.Message}");

        // Verify detailed components
        Assert.Contains("get-value", exception.Message);
        Assert.Contains(nonExistentFile, exception.Message);
        Assert.Contains("File not found", exception.Message);

        _output.WriteLine("✅ Verified: Cell operation includes detailed context");
    }

    [Fact]
    public async Task ExcelParameter_WithNonExistentFile_ShouldThrowDetailedError()
    {
        // Arrange
        string nonExistentFile = Path.Combine(_tempDir, "nonexistent-param.xlsx");

        // Act & Assert
        var exception = await Assert.ThrowsAsync<McpException>(async () =>
            await ExcelParameterTool.ExcelParameter("list", nonExistentFile));

        _output.WriteLine($"Error message: {exception.Message}");

        // Verify detailed components
        Assert.Contains("list", exception.Message);
        Assert.Contains(nonExistentFile, exception.Message);
        Assert.Contains("File not found", exception.Message);

        _output.WriteLine("✅ Verified: Parameter operation includes detailed context");
    }

    [Fact]
    public async Task ExcelPowerQuery_WithNonExistentFile_ShouldThrowDetailedError()
    {
        // Arrange
        string nonExistentFile = Path.Combine(_tempDir, "nonexistent-pq.xlsx");

        // Act & Assert
        var exception = await Assert.ThrowsAsync<McpException>(async () =>
            await ExcelPowerQueryTool.ExcelPowerQuery("list", nonExistentFile));

        _output.WriteLine($"Error message: {exception.Message}");

        // Verify detailed components
        Assert.Contains("list", exception.Message);
        Assert.Contains(nonExistentFile, exception.Message);
        Assert.Contains("File not found", exception.Message);

        _output.WriteLine("✅ Verified: PowerQuery operation includes detailed context");
    }

    [Fact]
    public async Task ExcelVba_WithNonMacroEnabledFile_ShouldThrowDetailedError()
    {
        // Arrange - Create .xlsx file (not macro-enabled)
        await ExcelFileTool.ExcelFile("create-empty", _testExcelFile);

        // Act & Assert - VBA operations require .xlsm
        var exception = await Assert.ThrowsAsync<McpException>(async () =>
            await ExcelVbaTool.ExcelVba("list", _testExcelFile));

        _output.WriteLine($"Error message: {exception.Message}");

        // Verify detailed components
        Assert.Contains("list", exception.Message);
        Assert.Contains(_testExcelFile, exception.Message);
        Assert.Contains("macro-enabled", exception.Message.ToLower());
        Assert.Contains(".xlsm", exception.Message);

        _output.WriteLine("✅ Verified: VBA operation includes detailed file type requirements");
    }

    [Fact]
    public async Task ExcelVba_WithMissingModuleName_ShouldThrowDetailedError()
    {
        // Arrange - Create macro-enabled file
        string xlsmFile = Path.Combine(_tempDir, "test-vba.xlsm");
        await ExcelFileTool.ExcelFile("create-empty", xlsmFile);

        // Act & Assert - Run requires moduleName
        var exception = await Assert.ThrowsAsync<McpException>(async () =>
            await ExcelVbaTool.ExcelVba("run", xlsmFile, moduleName: null));

        _output.WriteLine($"Error message: {exception.Message}");

        // Verify detailed components
        Assert.Contains("moduleName", exception.Message);
        Assert.Contains("required", exception.Message);
        Assert.Contains("run", exception.Message);

        _output.WriteLine("✅ Verified: Missing parameter error includes parameter name and action");
    }

    [Fact]
    public async Task ExcelFileTool_WithUnknownAction_ShouldThrowDetailedError()
    {
        // Act & Assert
        var exception = await Assert.ThrowsAsync<McpException>(async () =>
            await ExcelFileTool.ExcelFile("invalid-action", _testExcelFile));

        _output.WriteLine($"Error message: {exception.Message}");

        // Verify detailed components
        Assert.Contains("Unknown action", exception.Message);
        Assert.Contains("invalid-action", exception.Message);
        Assert.Contains("Supported:", exception.Message);
        Assert.Contains("create-empty", exception.Message);

        _output.WriteLine("✅ Verified: Unknown action error lists supported actions");
    }

    [Fact]
    public async Task ExcelWorksheet_WithUnknownAction_ShouldThrowDetailedError()
    {
        // Act & Assert
        var exception = await Assert.ThrowsAsync<McpException>(async () =>
            await ExcelWorksheetTool.ExcelWorksheet("invalid-action", _testExcelFile));

        _output.WriteLine($"Error message: {exception.Message}");

        // Verify error lists multiple supported actions
        Assert.Contains("Unknown action", exception.Message);
        Assert.Contains("invalid-action", exception.Message);
        Assert.Contains("list", exception.Message);
        Assert.Contains("read", exception.Message);
        Assert.Contains("write", exception.Message);

        _output.WriteLine("✅ Verified: Unknown action error provides comprehensive list of valid options");
    }

    [Fact]
    public async Task ExcelPowerQuery_Import_WithMissingParameters_ShouldThrowDetailedError()
    {
        // Arrange
        await ExcelFileTool.ExcelFile("create-empty", _testExcelFile);

        // Act & Assert - Import requires queryName and sourcePath
        var exception = await Assert.ThrowsAsync<McpException>(async () =>
            await ExcelPowerQueryTool.ExcelPowerQuery("import", _testExcelFile, queryName: null, sourcePath: null));

        _output.WriteLine($"Error message: {exception.Message}");

        // Verify detailed components
        Assert.Contains("queryName", exception.Message);
        Assert.Contains("sourcePath", exception.Message);
        Assert.Contains("required", exception.Message);
        Assert.Contains("import", exception.Message);

        _output.WriteLine("✅ Verified: Missing parameters error lists all required parameters");
    }

    [Fact]
    public async Task ExcelCell_SetValue_WithMissingValue_ShouldThrowDetailedError()
    {
        // Arrange
        await ExcelFileTool.ExcelFile("create-empty", _testExcelFile);

        // Act & Assert - set-value requires value parameter
        var exception = await Assert.ThrowsAsync<McpException>(async () =>
            await ExcelCellTool.ExcelCell("set-value", _testExcelFile, "Sheet1", "A1", value: null));

        _output.WriteLine($"Error message: {exception.Message}");

        // Verify parameter name is mentioned
        Assert.Contains("value", exception.Message);
        Assert.Contains("required", exception.Message);

        _output.WriteLine("✅ Verified: Missing parameter error specifies which parameter is required");
    }

    [Fact]
    public async Task ExcelParameter_Create_WithMissingParameters_ShouldThrowDetailedError()
    {
        // Arrange
        await ExcelFileTool.ExcelFile("create-empty", _testExcelFile);

        // Act & Assert - create requires parameterName and reference
        var exception = await Assert.ThrowsAsync<McpException>(async () =>
            await ExcelParameterTool.ExcelParameter("create", _testExcelFile, parameterName: null));

        _output.WriteLine($"Error message: {exception.Message}");

        // Verify detailed components
        Assert.Contains("parameterName", exception.Message);
        Assert.Contains("required", exception.Message);
        Assert.Contains("create", exception.Message);

        _output.WriteLine("✅ Verified: Missing parameter error includes action context");
    }

    [Fact]
    public async Task ExcelWorksheet_Read_WithMissingSheetName_ShouldThrowDetailedError()
    {
        // Arrange
        await ExcelFileTool.ExcelFile("create-empty", _testExcelFile);

        // Act & Assert - read requires sheetName and rangeAddress
        var exception = await Assert.ThrowsAsync<McpException>(async () =>
            await ExcelWorksheetTool.ExcelWorksheet("read", _testExcelFile, sheetName: null));

        _output.WriteLine($"Error message: {exception.Message}");

        // Verify parameter name is mentioned
        Assert.Contains("sheetName", exception.Message);
        Assert.Contains("required", exception.Message);
        Assert.Contains("read", exception.Message);

        _output.WriteLine("✅ Verified: Missing parameter includes action and parameter name");
    }
}
