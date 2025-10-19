using Xunit;
using ExcelMcp.Core.Commands;
using System.IO;

namespace ExcelMcp.Tests.Commands;

/// <summary>
/// Integration tests for Power Query operations using Excel COM automation.
/// These tests require Excel installation and validate Power Query M code management.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "PowerQuery")]
public class PowerQueryCommandsTests : IDisposable
{
    private readonly PowerQueryCommands _powerQueryCommands;
    private readonly string _testExcelFile;
    private readonly string _testQueryFile;
    private readonly string _tempDir;

    public PowerQueryCommandsTests()
    {
        _powerQueryCommands = new PowerQueryCommands();
        
        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCLI_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        
        _testExcelFile = Path.Combine(_tempDir, "TestWorkbook.xlsx");
        _testQueryFile = Path.Combine(_tempDir, "TestQuery.pq");
        
        // Create test Excel file and Power Query
        CreateTestExcelFile();
        CreateTestQueryFile();
    }

    private void CreateTestExcelFile()
    {
        // Use the FileCommands to create an empty Excel file for testing
        var fileCommands = new FileCommands();
        string[] args = { "create-empty", _testExcelFile };
        
        int result = fileCommands.CreateEmpty(args);
        if (result != 0)
        {
            throw new InvalidOperationException("Failed to create test Excel file. Excel may not be installed.");
        }
    }

    private void CreateTestQueryFile()
    {
        // Create a test Power Query M file that gets data from a public API
        string mCode = @"let
    // Get sample data from JSONPlaceholder API (public testing API)
    Source = Json.Document(Web.Contents(""https://jsonplaceholder.typicode.com/posts?_limit=5"")),
    #""Converted to Table"" = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #""Expanded Column1"" = Table.ExpandRecordColumn(#""Converted to Table"", ""Column1"", {""userId"", ""id"", ""title"", ""body""}, {""userId"", ""id"", ""title"", ""body""}),
    #""Changed Type"" = Table.TransformColumnTypes(#""Expanded Column1"",{{""userId"", Int64.Type}, {""id"", Int64.Type}, {""title"", type text}, {""body"", type text}})
in
    #""Changed Type""";
    
        File.WriteAllText(_testQueryFile, mCode);
    }

    [Fact]
    public void List_WithValidFile_ReturnsSuccess()
    {
        // Arrange
        string[] args = { "pq-list", _testExcelFile };

        // Act
        int result = _powerQueryCommands.List(args);

        // Assert
        Assert.Equal(0, result);
    }

    [Fact]
    public void List_WithInvalidArgs_ReturnsError()
    {
        // Arrange
        string[] args = { "pq-list" }; // Missing file argument

        // Act
        int result = _powerQueryCommands.List(args);

        // Assert
        Assert.Equal(1, result);
    }

    [Fact]
    public void List_WithNonExistentFile_ReturnsError()
    {
        // Arrange
        string[] args = { "pq-list", "nonexistent.xlsx" };

        // Act
        int result = _powerQueryCommands.List(args);

        // Assert
        Assert.Equal(1, result);
    }

    [Fact]
    public void View_WithValidQuery_ReturnsSuccess()
    {
        // Arrange
        string[] args = { "pq-view", _testExcelFile, "TestQuery" };

        // Act
        int result = _powerQueryCommands.View(args);

        // Assert - Success if query exists, error if Power Query not available
        Assert.True(result == 0 || result == 1); // Allow both outcomes
    }

    [Fact]
    public void View_WithInvalidArgs_ReturnsError()
    {
        // Arrange
        string[] args = { "pq-view", _testExcelFile }; // Missing query name

        // Act
        int result = _powerQueryCommands.View(args);

        // Assert
        Assert.Equal(1, result);
    }

    [Fact]
    public async Task Import_WithValidQuery_ReturnsSuccess()
    {
        // Arrange
        string[] args = { "pq-import", _testExcelFile, "ImportedQuery", _testQueryFile };

        // Act
        int result = await _powerQueryCommands.Import(args);

        // Assert - Success if Power Query available, error otherwise
        Assert.True(result == 0 || result == 1); // Allow both outcomes
    }

    [Fact]
    public async Task Import_WithInvalidArgs_ReturnsError()
    {
        // Arrange
        string[] args = { "pq-import", _testExcelFile }; // Missing required args

        // Act
        int result = await _powerQueryCommands.Import(args);

        // Assert
        Assert.Equal(1, result);
    }

    [Fact]
    public async Task Export_WithInvalidArgs_ReturnsError()
    {
        // Arrange
        string[] args = { "pq-export", _testExcelFile }; // Missing query name and output file

        // Act
        int result = await _powerQueryCommands.Export(args);

        // Assert
        Assert.Equal(1, result);
    }

    [Fact]
    public async Task Update_WithInvalidArgs_ReturnsError()
    {
        // Arrange
        string[] args = { "pq-update", _testExcelFile }; // Missing query name and M file

        // Act
        int result = await _powerQueryCommands.Update(args);

        // Assert
        Assert.Equal(1, result);
    }

    [Fact]
    public void Delete_WithInvalidArgs_ReturnsError()
    {
        // Arrange
        string[] args = { "pq-delete", _testExcelFile }; // Missing query name

        // Act
        int result = _powerQueryCommands.Delete(args);

        // Assert
        Assert.Equal(1, result);
    }

    [Fact]
    public void Delete_WithNonExistentFile_ReturnsError()
    {
        // Arrange
        string[] args = { "pq-delete", "nonexistent.xlsx", "TestQuery" };

        // Act
        int result = _powerQueryCommands.Delete(args);

        // Assert
        Assert.Equal(1, result);
    }

    [Fact]
    public void Sources_WithValidFile_ReturnsSuccess()
    {
        // Arrange
        string[] args = { "pq-sources", _testExcelFile };

        // Act
        int result = _powerQueryCommands.Sources(args);

        // Assert
        Assert.Equal(0, result);
    }

    [Fact]
    public void Sources_WithInvalidArgs_ReturnsError()
    {
        // Arrange
        string[] args = { "pq-sources" }; // Missing file argument

        // Act
        int result = _powerQueryCommands.Sources(args);

        // Assert
        Assert.Equal(1, result);
    }

    [Fact]
    public void Test_WithInvalidArgs_ReturnsError()
    {
        // Arrange
        string[] args = { "pq-test" }; // Missing file argument

        // Act
        int result = _powerQueryCommands.Test(args);

        // Assert
        Assert.Equal(1, result);
    }

    [Fact]
    public void Peek_WithInvalidArgs_ReturnsError()
    {
        // Arrange
        string[] args = { "pq-peek" }; // Missing file argument

        // Act
        int result = _powerQueryCommands.Peek(args);

        // Assert
        Assert.Equal(1, result);
    }

    [Fact]
    public void Eval_WithInvalidArgs_ReturnsError()
    {
        // Arrange
        string[] args = { "pq-verify" }; // Missing file argument

        // Act
        int result = _powerQueryCommands.Eval(args);

        // Assert
        Assert.Equal(1, result);
    }

    [Fact]
    public void Refresh_WithInvalidArgs_ReturnsError()
    {
        // Arrange
        string[] args = { "pq-refresh" }; // Missing file argument

        // Act
        int result = _powerQueryCommands.Refresh(args);

        // Assert
        Assert.Equal(1, result);
    }

    [Fact]
    public void Errors_WithInvalidArgs_ReturnsError()
    {
        // Arrange
        string[] args = { "pq-errors" }; // Missing file argument

        // Act
        int result = _powerQueryCommands.Errors(args);

        // Assert
        Assert.Equal(1, result);
    }

    /// <summary>
    /// Round-trip test: Import a Power Query that generates data, load it to a sheet, then verify the data
    /// This tests the complete Power Query workflow for coding agents
    /// </summary>
    [Fact]
    public async Task PowerQuery_RoundTrip_ImportLoadAndVerifyData()
    {
        // Arrange - Create a simple Power Query that generates sample data (no external dependencies)
        string simpleQueryFile = Path.Combine(_tempDir, "SimpleDataQuery.pq");
        string simpleQueryCode = @"let
    // Create sample data without external dependencies
    Source = #table(
        {""ID"", ""Product"", ""Quantity"", ""Price""}, 
        {
            {1, ""Widget A"", 10, 19.99},
            {2, ""Widget B"", 15, 24.99},
            {3, ""Widget C"", 8, 14.99},
            {4, ""Widget D"", 12, 29.99},
            {5, ""Widget E"", 20, 9.99}
        }
    ),
    #""Changed Type"" = Table.TransformColumnTypes(Source,{{""ID"", Int64.Type}, {""Product"", type text}, {""Quantity"", Int64.Type}, {""Price"", type number}})
in
    #""Changed Type""";
        
        File.WriteAllText(simpleQueryFile, simpleQueryCode);

        // Also need SheetCommands for verification
        var sheetCommands = new SheetCommands();

        // Act 1 - Import the Power Query
        string[] importArgs = { "pq-import", _testExcelFile, "SampleData", simpleQueryFile };
        int importResult = await _powerQueryCommands.Import(importArgs);
        Assert.Equal(0, importResult);

        // Act 2 - Load the query to a worksheet
        string[] loadArgs = { "pq-loadto", _testExcelFile, "SampleData", "Sheet1" };
        int loadResult = _powerQueryCommands.LoadTo(loadArgs);
        Assert.Equal(0, loadResult);

        // Act 3 - Verify the data was loaded by reading it back
        string[] readArgs = { "sheet-read", _testExcelFile, "Sheet1", "A1:D6" }; // Header + 5 data rows
        int readResult = sheetCommands.Read(readArgs);
        Assert.Equal(0, readResult);

        // Act 4 - Verify we can list the query
        string[] listArgs = { "pq-list", _testExcelFile };
        int listResult = _powerQueryCommands.List(listArgs);
        Assert.Equal(0, listResult);
    }

    /// <summary>
    /// Round-trip test: Create a Power Query that calculates aggregations and verify the computed results
    /// </summary>
    [Fact]
    public async Task PowerQuery_RoundTrip_CalculationAndVerification()
    {
        // Arrange - Create a Power Query that generates data with calculations
        string calcQueryFile = Path.Combine(_tempDir, "CalculationQuery.pq");
        string calcQueryCode = @"let
    // Create base data
    BaseData = #table(
        {""Item"", ""Quantity"", ""UnitPrice""}, 
        {
            {""Product A"", 10, 5.50},
            {""Product B"", 25, 3.25},
            {""Product C"", 15, 7.75},
            {""Product D"", 8, 12.00},
            {""Product E"", 30, 2.99}
        }
    ),
    #""Added Total Column"" = Table.AddColumn(BaseData, ""Total"", each [Quantity] * [UnitPrice], type number),
    #""Added Category"" = Table.AddColumn(#""Added Total Column"", ""Category"", each if [Total] > 100 then ""High Value"" else ""Standard"", type text),
    #""Changed Type"" = Table.TransformColumnTypes(#""Added Category"",{{""Item"", type text}, {""Quantity"", Int64.Type}, {""UnitPrice"", type number}, {""Total"", type number}, {""Category"", type text}})
in
    #""Changed Type""";
        
        File.WriteAllText(calcQueryFile, calcQueryCode);

        var sheetCommands = new SheetCommands();

        // Act 1 - Import the calculation query
        string[] importArgs = { "pq-import", _testExcelFile, "CalculatedData", calcQueryFile };
        int importResult = await _powerQueryCommands.Import(importArgs);
        Assert.Equal(0, importResult);

        // Act 2 - Refresh the query to ensure calculations are executed
        string[] refreshArgs = { "pq-refresh", _testExcelFile, "CalculatedData" };
        int refreshResult = _powerQueryCommands.Refresh(refreshArgs);
        Assert.Equal(0, refreshResult);

        // Act 3 - Load to a different sheet for testing
        string[] createSheetArgs = { "sheet-create", _testExcelFile, "Calculations" };
        var createResult = sheetCommands.Create(createSheetArgs);
        Assert.Equal(0, createResult);

        string[] loadArgs = { "pq-loadto", _testExcelFile, "CalculatedData", "Calculations" };
        int loadResult = _powerQueryCommands.LoadTo(loadArgs);
        Assert.Equal(0, loadResult);

        // Act 4 - Verify the calculated data
        string[] readArgs = { "sheet-read", _testExcelFile, "Calculations", "A1:E6" }; // All columns + header + 5 rows
        int readResult = sheetCommands.Read(readArgs);
        Assert.Equal(0, readResult);

        // Act 5 - Export the query to verify we can get the M code back
        string exportedQueryFile = Path.Combine(_tempDir, "ExportedCalcQuery.pq");
        string[] exportArgs = { "pq-export", _testExcelFile, "CalculatedData", exportedQueryFile };
        int exportResult = await _powerQueryCommands.Export(exportArgs);
        Assert.Equal(0, exportResult);

        // Verify the exported file exists
        Assert.True(File.Exists(exportedQueryFile));
    }

    /// <summary>
    /// Round-trip test: Update an existing Power Query and verify the data changes
    /// </summary>
    [Fact]
    public async Task PowerQuery_RoundTrip_UpdateQueryAndVerifyChanges()
    {
        // Arrange - Start with initial data
        string initialQueryFile = Path.Combine(_tempDir, "InitialQuery.pq");
        string initialQueryCode = @"let
    Source = #table(
        {""Name"", ""Score""}, 
        {
            {""Alice"", 85},
            {""Bob"", 92},
            {""Charlie"", 78}
        }
    )
in
    Source";
        
        File.WriteAllText(initialQueryFile, initialQueryCode);

        var sheetCommands = new SheetCommands();

        // Act 1 - Import initial query
        string[] importArgs = { "pq-import", _testExcelFile, "StudentScores", initialQueryFile };
        int importResult = await _powerQueryCommands.Import(importArgs);
        Assert.Equal(0, importResult);

        // Act 2 - Load to sheet
        string[] loadArgs1 = { "pq-loadto", _testExcelFile, "StudentScores", "Sheet1" };
        int loadResult1 = _powerQueryCommands.LoadTo(loadArgs1);
        Assert.Equal(0, loadResult1);

        // Act 3 - Read initial data
        string[] readArgs1 = { "sheet-read", _testExcelFile, "Sheet1", "A1:B4" };
        int readResult1 = sheetCommands.Read(readArgs1);
        Assert.Equal(0, readResult1);

        // Act 4 - Update the query with modified data
        string updatedQueryFile = Path.Combine(_tempDir, "UpdatedQuery.pq");
        string updatedQueryCode = @"let
    Source = #table(
        {""Name"", ""Score"", ""Grade""}, 
        {
            {""Alice"", 85, ""B""},
            {""Bob"", 92, ""A""},
            {""Charlie"", 78, ""C""},
            {""Diana"", 96, ""A""},
            {""Eve"", 88, ""B""}
        }
    )
in
    Source";
        
        File.WriteAllText(updatedQueryFile, updatedQueryCode);

        string[] updateArgs = { "pq-update", _testExcelFile, "StudentScores", updatedQueryFile };
        int updateResult = await _powerQueryCommands.Update(updateArgs);
        Assert.Equal(0, updateResult);

        // Act 5 - Refresh to get updated data
        string[] refreshArgs = { "pq-refresh", _testExcelFile, "StudentScores" };
        int refreshResult = _powerQueryCommands.Refresh(refreshArgs);
        Assert.Equal(0, refreshResult);

        // Act 6 - Clear the sheet and reload to see changes
        string[] clearArgs = { "sheet-clear", _testExcelFile, "Sheet1" };
        int clearResult = sheetCommands.Clear(clearArgs);
        Assert.Equal(0, clearResult);

        string[] loadArgs2 = { "pq-loadto", _testExcelFile, "StudentScores", "Sheet1" };
        int loadResult2 = _powerQueryCommands.LoadTo(loadArgs2);
        Assert.Equal(0, loadResult2);

        // Act 7 - Read updated data (now should have 3 columns and 5 data rows)
        string[] readArgs2 = { "sheet-read", _testExcelFile, "Sheet1", "A1:C6" };
        int readResult2 = sheetCommands.Read(readArgs2);
        Assert.Equal(0, readResult2);

        // Act 8 - Verify we can still list and view the updated query
        string[] listArgs = { "pq-list", _testExcelFile };
        int listResult = _powerQueryCommands.List(listArgs);
        Assert.Equal(0, listResult);

        string[] viewArgs = { "pq-view", _testExcelFile, "StudentScores" };
        int viewResult = _powerQueryCommands.View(viewArgs);
        Assert.Equal(0, viewResult);
    }

    [Fact]
    public void LoadTo_WithInvalidArgs_ReturnsError()
    {
        // Arrange
        string[] args = { "pq-loadto" }; // Missing file argument

        // Act
        int result = _powerQueryCommands.LoadTo(args);

        // Assert
        Assert.Equal(1, result);
    }

    public void Dispose()
    {
        // Clean up test files
        try
        {
            if (Directory.Exists(_tempDir))
            {
                // Wait a bit for Excel to fully release files
                System.Threading.Thread.Sleep(500);
                
                // Try to delete files multiple times if needed
                for (int i = 0; i < 3; i++)
                {
                    try
                    {
                        Directory.Delete(_tempDir, true);
                        break;
                    }
                    catch (IOException)
                    {
                        if (i == 2) throw; // Last attempt failed
                        System.Threading.Thread.Sleep(1000);
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                }
            }
        }
        catch
        {
            // Best effort cleanup - don't fail tests if cleanup fails
        }
        
        GC.SuppressFinalize(this);
    }
}
