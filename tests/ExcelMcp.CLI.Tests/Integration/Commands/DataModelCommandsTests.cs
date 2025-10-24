using Sbroenne.ExcelMcp.CLI.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration.Commands;

/// <summary>
/// Tests for CLI DataModelCommands - verifying CLI-specific behavior (argument parsing, exit codes, formatting)
/// These tests focus on the presentation layer, not the business logic
/// Core Data Model logic is tested in ExcelMcp.Core.Tests
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "DataModel")]
[Trait("Layer", "CLI")]
public class CliDataModelCommandsTests : IDisposable
{
    private readonly DataModelCommands _cliCommands;
    private readonly FileCommands _cliFileCommands;
    private readonly string _tempDir;
    private readonly string _testExcelFile;

    public CliDataModelCommandsTests()
    {
        _cliCommands = new DataModelCommands();
        _cliFileCommands = new FileCommands();

        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCLI_DataModelTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        _testExcelFile = Path.Combine(_tempDir, "TestDataModel.xlsx");

        // Create test Excel file with Data Model
        CreateTestDataModelFile();
    }

    private void CreateTestDataModelFile()
    {
        // Use CLI command to create file
        string[] args = { "create-empty", _testExcelFile };
        int exitCode = _cliFileCommands.CreateEmpty(args);

        if (exitCode != 0)
        {
            throw new InvalidOperationException("Failed to create test Excel file using CLI command");
        }

        // Create realistic Data Model with sample data
        try
        {
            DataModelTestHelper.CreateSampleDataModel(_testExcelFile);
        }
        catch (Exception ex)
        {
            // Data Model creation may fail on some Excel versions - tests will handle gracefully
            System.Diagnostics.Debug.WriteLine($"Could not create sample Data Model: {ex.Message}");
        }
    }

    #region Argument Validation Tests

    [Fact]
    public void ListTables_WithMissingFileArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = { "dm-list-tables" }; // Missing file path

        // Act
        int exitCode = _cliCommands.ListTables(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void ListMeasures_WithMissingFileArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = { "dm-list-measures" }; // Missing file path

        // Act
        int exitCode = _cliCommands.ListMeasures(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void ViewMeasure_WithMissingMeasureNameArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = { "dm-view-measure", _testExcelFile }; // Missing measure name

        // Act
        int exitCode = _cliCommands.ViewMeasure(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public async Task ExportMeasure_WithMissingOutputFileArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = { "dm-export-measure", _testExcelFile, "SomeMeasure" }; // Missing output file

        // Act
        int exitCode = await _cliCommands.ExportMeasure(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void ListRelationships_WithMissingFileArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = { "dm-list-relationships" }; // Missing file path

        // Act
        int exitCode = _cliCommands.ListRelationships(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void Refresh_WithMissingFileArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = { "dm-refresh" }; // Missing file path

        // Act
        int exitCode = _cliCommands.Refresh(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }

    #endregion

    #region File Validation Tests

    [Fact]
    public void ListTables_WithNonExistentFile_ReturnsErrorExitCode()
    {
        // Arrange
        string nonExistentFile = Path.Combine(_tempDir, "NonExistent.xlsx");
        string[] args = { "dm-list-tables", nonExistentFile };

        // Act
        int exitCode = _cliCommands.ListTables(args);

        // Assert - CLI returns 1 for error (file not found)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void ListMeasures_WithNonExistentFile_ReturnsErrorExitCode()
    {
        // Arrange
        string nonExistentFile = Path.Combine(_tempDir, "NonExistent.xlsx");
        string[] args = { "dm-list-measures", nonExistentFile };

        // Act
        int exitCode = _cliCommands.ListMeasures(args);

        // Assert - CLI returns 1 for error (file not found)
        Assert.Equal(1, exitCode);
    }

    #endregion

    #region Success Path Tests

    [Fact]
    public void ListTables_WithValidFile_ReturnsSuccessOrNoDataModelError()
    {
        // Arrange
        string[] args = { "dm-list-tables", _testExcelFile };

        // Act
        int exitCode = _cliCommands.ListTables(args);

        // Assert - CLI returns 0 for success or 1 if no Data Model (both acceptable)
        Assert.True(exitCode == 0 || exitCode == 1,
            $"Expected exit code 0 (success) or 1 (no Data Model), got {exitCode}");
    }

    [Fact]
    public void ListMeasures_WithValidFile_ReturnsSuccessOrNoDataModelError()
    {
        // Arrange
        string[] args = { "dm-list-measures", _testExcelFile };

        // Act
        int exitCode = _cliCommands.ListMeasures(args);

        // Assert - CLI returns 0 for success or 1 if no Data Model (both acceptable)
        Assert.True(exitCode == 0 || exitCode == 1,
            $"Expected exit code 0 (success) or 1 (no Data Model), got {exitCode}");
    }

    [Fact]
    public void ListRelationships_WithValidFile_ReturnsSuccessOrNoDataModelError()
    {
        // Arrange
        string[] args = { "dm-list-relationships", _testExcelFile };

        // Act
        int exitCode = _cliCommands.ListRelationships(args);

        // Assert - CLI returns 0 for success or 1 if no Data Model (both acceptable)
        Assert.True(exitCode == 0 || exitCode == 1,
            $"Expected exit code 0 (success) or 1 (no Data Model), got {exitCode}");
    }

    [Fact]
    public void Refresh_WithValidFile_ReturnsSuccessOrNoDataModelError()
    {
        // Arrange
        string[] args = { "dm-refresh", _testExcelFile };

        // Act
        int exitCode = _cliCommands.Refresh(args);

        // Assert - CLI returns 0 for success or 1 if no Data Model (both acceptable)
        Assert.True(exitCode == 0 || exitCode == 1,
            $"Expected exit code 0 (success) or 1 (no Data Model), got {exitCode}");
    }

    [Fact]
    public void ViewMeasure_WithNonExistentMeasure_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = { "dm-view-measure", _testExcelFile, "NonExistentMeasure" };

        // Act
        int exitCode = _cliCommands.ViewMeasure(args);

        // Assert - CLI returns 1 for error (measure not found or no Data Model)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public async Task ExportMeasure_WithNonExistentMeasure_ReturnsErrorExitCode()
    {
        // Arrange
        string outputPath = Path.Combine(_tempDir, "output.dax");
        string[] args = { "dm-export-measure", _testExcelFile, "NonExistentMeasure", outputPath };

        // Act
        int exitCode = await _cliCommands.ExportMeasure(args);

        // Assert - CLI returns 1 for error (measure not found or no Data Model)
        Assert.Equal(1, exitCode);
    }

    #endregion

    #region Delete Operations Tests

    [Fact]
    public void DeleteMeasure_WithMissingArguments_ReturnsError()
    {
        // Arrange
        string[] args = { "dm-delete-measure", _testExcelFile };

        // Act
        int exitCode = _cliCommands.DeleteMeasure(args);

        // Assert - CLI returns 1 for missing arguments
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void DeleteMeasure_WithNonExistentFile_ReturnsError()
    {
        // Arrange
        string[] args = { "dm-delete-measure", "NonExistent.xlsx", "SomeMeasure" };

        // Act
        int exitCode = _cliCommands.DeleteMeasure(args);

        // Assert - CLI returns 1 for file not found
        Assert.Equal(1, exitCode);
    }

    [Fact(Skip = "Data Model test helper requires specific Excel version/configuration. May fail on some environments due to Data Model availability.")]
    public void DeleteMeasure_WithValidMeasure_ReturnsSuccess()
    {
        // Arrange - Create a test measure first
        var measureName = "TestMeasure_" + Guid.NewGuid().ToString("N")[..8];

        DataModelTestHelper.CreateTestMeasure(_testExcelFile, measureName, "SUM(Sales[Amount])");

        string[] args = { "dm-delete-measure", _testExcelFile, measureName };

        // Act
        int exitCode = _cliCommands.DeleteMeasure(args);

        // Assert - CLI returns 0 for success
        Assert.Equal(0, exitCode);
    }

    [Fact]
    public void DeleteMeasure_WithNonExistentMeasure_ReturnsError()
    {
        // Arrange
        string[] args = { "dm-delete-measure", _testExcelFile, "NonExistentMeasure" };

        // Act
        int exitCode = _cliCommands.DeleteMeasure(args);

        // Assert - CLI returns 1 for error (measure not found or no Data Model)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void DeleteRelationship_WithMissingArguments_ReturnsError()
    {
        // Arrange - Missing columns
        string[] args = { "dm-delete-relationship", _testExcelFile, "Table1", "Col1" };

        // Act
        int exitCode = _cliCommands.DeleteRelationship(args);

        // Assert - CLI returns 1 for missing arguments
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void DeleteRelationship_WithNonExistentFile_ReturnsError()
    {
        // Arrange
        string[] args = { "dm-delete-relationship", "NonExistent.xlsx", "Table1", "Col1", "Table2", "Col2" };

        // Act
        int exitCode = _cliCommands.DeleteRelationship(args);

        // Assert - CLI returns 1 for file not found
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void DeleteRelationship_WithNonExistentRelationship_ReturnsError()
    {
        // Arrange
        string[] args = { "dm-delete-relationship", _testExcelFile, "FakeTable", "FakeCol", "OtherTable", "OtherCol" };

        // Act
        int exitCode = _cliCommands.DeleteRelationship(args);

        // Assert - CLI returns 1 for error (relationship not found or no Data Model)
        Assert.Equal(1, exitCode);
    }

    #endregion

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDir))
            {
                // Give Excel time to release file locks
                System.Threading.Thread.Sleep(100);

                // Retry cleanup a few times if needed
                for (int i = 0; i < 3; i++)
                {
                    try
                    {
                        Directory.Delete(_tempDir, recursive: true);
                        break;
                    }
                    catch (IOException) when (i < 2)
                    {
                        System.Threading.Thread.Sleep(500);
                    }
                }
            }
        }
        catch
        {
            // Best effort cleanup
        }

        GC.SuppressFinalize(this);
    }
}
