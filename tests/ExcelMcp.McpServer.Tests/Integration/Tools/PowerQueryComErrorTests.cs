using Xunit;
using Xunit.Abstractions;
using Sbroenne.ExcelMcp.Core.Commands;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Focused tests to diagnose COM error 0x800A03EC in Power Query operations
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
public class PowerQueryComErrorTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private readonly FileCommands _fileCommands;
    private readonly PowerQueryCommands _powerQueryCommands;
    private readonly SheetCommands _sheetCommands;

    public PowerQueryComErrorTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Combine(Path.GetTempPath(), $"PowerQueryComError_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        _fileCommands = new FileCommands();
        _powerQueryCommands = new PowerQueryCommands();
        _sheetCommands = new SheetCommands();
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
    public async Task SetLoadToTable_WithSimpleQuery_ShouldWork()
    {
        // Arrange
        var testFile = Path.Combine(_tempDir, "simple-test.xlsx");
        var queryName = "SimpleTestQuery";
        var targetSheet = "DataSheet";
        
        var simpleMCode = @"
let
    Source = #table(
        {""Name"", ""Value""}, 
        {{""Item1"", 10}, {""Item2"", 20}, {""Item3"", 30}}
    )
in
    Source";
        
        var mCodeFile = Path.Combine(_tempDir, "simple-query.pq");
        File.WriteAllText(mCodeFile, simpleMCode);

        // Act & Assert Step by Step
        _output.WriteLine("Step 1: Creating Excel file...");
        var createResult = _fileCommands.CreateEmpty(testFile);
        Assert.True(createResult.Success, $"Failed to create Excel file: {createResult.ErrorMessage}");

        _output.WriteLine("Step 2: Importing Power Query...");
        var importResult = await _powerQueryCommands.Import(testFile, queryName, mCodeFile);
        Assert.True(importResult.Success, $"Failed to import Power Query: {importResult.ErrorMessage}");

        _output.WriteLine("Step 3: Listing queries to verify import...");
        var listResult = _powerQueryCommands.List(testFile);
        Assert.True(listResult.Success, $"Failed to list queries: {listResult.ErrorMessage}");
        Assert.Contains(listResult.Queries, q => q.Name == queryName);

        _output.WriteLine("Step 4: Attempting to set load to table (critical step)...");
        var setLoadResult = _powerQueryCommands.SetLoadToTable(testFile, queryName, targetSheet);
        
        if (!setLoadResult.Success)
        {
            _output.WriteLine($"ERROR: {setLoadResult.ErrorMessage}");
            _output.WriteLine("This error will help us understand the COM issue");
        }
        
        Assert.True(setLoadResult.Success, $"Failed to set load to table: {setLoadResult.ErrorMessage}");
    }

    [Fact]
    public async Task SetLoadToTable_WithExistingSheet_ShouldWork()
    {
        // Arrange
        var testFile = Path.Combine(_tempDir, "existing-sheet-test.xlsx");
        var queryName = "ExistingSheetQuery";
        var targetSheet = "PreExistingSheet";
        
        var simpleMCode = @"
let
    Source = #table(
        {""Column1"", ""Column2""}, 
        {{""A"", 1}, {""B"", 2}}
    )
in
    Source";
        
        var mCodeFile = Path.Combine(_tempDir, "existing-sheet-query.pq");
        File.WriteAllText(mCodeFile, simpleMCode);

        // Act & Assert
        _output.WriteLine("Step 1: Creating Excel file...");
        var createResult = _fileCommands.CreateEmpty(testFile);
        Assert.True(createResult.Success, $"Failed to create Excel file: {createResult.ErrorMessage}");

        _output.WriteLine("Step 2: Creating target sheet first...");
        var createSheetResult = _sheetCommands.Create(testFile, targetSheet);
        Assert.True(createSheetResult.Success, $"Failed to create sheet: {createSheetResult.ErrorMessage}");

        _output.WriteLine("Step 3: Importing Power Query...");
        var importResult = await _powerQueryCommands.Import(testFile, queryName, mCodeFile);
        Assert.True(importResult.Success, $"Failed to import Power Query: {importResult.ErrorMessage}");

        _output.WriteLine("Step 4: Setting load to existing sheet...");
        var setLoadResult = _powerQueryCommands.SetLoadToTable(testFile, queryName, targetSheet);
        
        if (!setLoadResult.Success)
        {
            _output.WriteLine($"ERROR WITH EXISTING SHEET: {setLoadResult.ErrorMessage}");
        }
        
        Assert.True(setLoadResult.Success, $"Failed to set load to existing sheet: {setLoadResult.ErrorMessage}");
    }
}
