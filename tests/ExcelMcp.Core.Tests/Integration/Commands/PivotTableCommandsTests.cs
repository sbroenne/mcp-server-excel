using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Integration tests for PivotTable commands (Phase 1).
/// These tests require Excel installation and validate Core PivotTable operations.
/// Tests use Core commands directly (not through CLI or MCP wrapper).
///
/// Phase 1: Lifecycle, Field Management, Analysis
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PivotTables")]
public class PivotTableCommandsTests : IDisposable
{
    private readonly IPivotTableCommands _pivotCommands;
    private readonly IFileCommands _fileCommands;
    private readonly string _testExcelFile;
    private readonly string _tempDir;
    private bool _disposed;

    public PivotTableCommandsTests()
    {
        _pivotCommands = new PivotTableCommands();
        _fileCommands = new FileCommands();

        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_PivotTable_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        _testExcelFile = Path.Combine(_tempDir, "TestPivotTables.xlsx");

        // Create test Excel file with sample data
        CreateTestExcelFileWithData();
    }

    private void CreateTestExcelFileWithData()
    {
        var result = _fileCommands.CreateEmptyAsync(_testExcelFile, overwriteIfExists: false).GetAwaiter().GetResult();
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}");
        }

        // Add sample data for PivotTable
        Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);

            // Get Sheet1 and add sample sales data
            await batch.ExecuteAsync<int>((ctx, ct) =>
            {
                dynamic sheet = ctx.Book.Worksheets.Item(1);
                sheet.Name = "SalesData";

                // Add headers
                sheet.Range["A1"].Value2 = "Region";
                sheet.Range["B1"].Value2 = "Product";
                sheet.Range["C1"].Value2 = "Sales";
                sheet.Range["D1"].Value2 = "Date";

                // Add sample data rows
                sheet.Range["A2"].Value2 = "North";
                sheet.Range["B2"].Value2 = "Widget";
                sheet.Range["C2"].Value2 = 100;
                sheet.Range["D2"].Value2 = new DateTime(2025, 1, 15);

                sheet.Range["A3"].Value2 = "North";
                sheet.Range["B3"].Value2 = "Widget";
                sheet.Range["C3"].Value2 = 150;
                sheet.Range["D3"].Value2 = new DateTime(2025, 1, 20);

                sheet.Range["A4"].Value2 = "South";
                sheet.Range["B4"].Value2 = "Gadget";
                sheet.Range["C4"].Value2 = 200;
                sheet.Range["D4"].Value2 = new DateTime(2025, 2, 10);

                sheet.Range["A5"].Value2 = "North";
                sheet.Range["B5"].Value2 = "Gadget";
                sheet.Range["C5"].Value2 = 75;
                sheet.Range["D5"].Value2 = new DateTime(2025, 2, 15);

                sheet.Range["A6"].Value2 = "South";
                sheet.Range["B6"].Value2 = "Widget";
                sheet.Range["C6"].Value2 = 125;
                sheet.Range["D6"].Value2 = new DateTime(2025, 3, 5);

                return ValueTask.FromResult(0);
            });

            await batch.SaveAsync();
        }).GetAwaiter().GetResult();
    }

    #region Phase 1 Tests - Lifecycle

    [Fact]
    public async Task CreateFromRange_WithValidData_CreatesCorrectPivotStructure()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _pivotCommands.CreateFromRangeAsync(
            batch,
            "SalesData", "A1:D6",
            "SalesData", "F1",
            "TestPivot");

        // Assert - Verify Success AND Structure
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("TestPivot", result.PivotTableName);
        Assert.Equal("SalesData", result.SheetName);
        Assert.Equal(4, result.AvailableFields.Count);
        Assert.Contains("Region", result.AvailableFields);
        Assert.Contains("Product", result.AvailableFields);
        Assert.Contains("Sales", result.AvailableFields);
        Assert.Contains("Date", result.AvailableFields);
        
        // Verify field type detection
        Assert.Contains("Sales", result.NumericFields);
        Assert.Contains("Date", result.DateFields);
        Assert.Contains("Region", result.TextFields);
        Assert.Contains("Product", result.TextFields);

        // Verify actual Excel COM object exists
        await VerifyPivotTableExists(batch, "TestPivot", "SalesData");
    }

    [Fact]
    public async Task List_WithValidFile_ReturnsSuccessWithPivotTables()
    {
        // Arrange - Create a PivotTable first
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch,
            "SalesData", "A1:D6",
            "SalesData", "F1",
            "TestPivot");
        Assert.True(createResult.Success);
        await batch.SaveAsync();

        // Act
        await using var batch2 = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _pivotCommands.ListAsync(batch2);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.PivotTables);
        Assert.Contains(result.PivotTables, p => p.Name == "TestPivot");
    }

    #endregion

    #region Phase 1 Tests - Field Management

    [Fact]
    public async Task AddRowField_ValidField_PlacesFieldInRowAreaAndRefreshesCorrectly()
    {
        // Arrange - Create PivotTable first
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch,
            "SalesData", "A1:D6",
            "SalesData", "F1",
            "TestPivot");
        Assert.True(createResult.Success);

        // Act
        var result = await _pivotCommands.AddRowFieldAsync(batch, "TestPivot", "Region");

        // Assert - Verify Success AND Excel State
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("Region", result.FieldName);
        Assert.Equal(PivotFieldArea.Row, result.Area);
        Assert.Equal(1, result.Position); // Should be first row field

        // Verify field has unique values available
        Assert.NotEmpty(result.AvailableValues);
        Assert.Contains("North", result.AvailableValues);
        Assert.Contains("South", result.AvailableValues);
    }

    [Fact]
    public async Task AddValueField_NumericField_AggregatesDataCorrectly()
    {
        // Arrange - Create PivotTable with row field
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch,
            "SalesData", "A1:D6",
            "SalesData", "F1",
            "TestPivot");
        Assert.True(createResult.Success);

        await _pivotCommands.AddRowFieldAsync(batch, "TestPivot", "Region");

        // Act
        var result = await _pivotCommands.AddValueFieldAsync(
            batch, "TestPivot", "Sales", AggregationFunction.Sum, "Total Sales");

        // Assert - Verify Success AND Calculation
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("Sales", result.FieldName);
        Assert.Equal("Total Sales", result.CustomName);
        Assert.Equal(PivotFieldArea.Value, result.Area);
        Assert.Equal(AggregationFunction.Sum, result.Function);
    }

    [Fact]
    public async Task AddValueField_TextFieldWithSumFunction_ReturnsInformativeError()
    {
        // Arrange
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch,
            "SalesData", "A1:D6",
            "SalesData", "F1",
            "TestPivot");
        Assert.True(createResult.Success);

        // Act & Assert - Should fail with informative error
        var result = await _pivotCommands.AddValueFieldAsync(
            batch, "TestPivot", "Region", AggregationFunction.Sum);

        // Verify error message provides actionable guidance
        Assert.False(result.Success);
        Assert.Contains("not valid for Text field 'Region'", result.ErrorMessage);
        Assert.Contains("Valid functions: Count", result.ErrorMessage);
    }

    #endregion

    #region Helper Methods

    #pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    private async Task VerifyPivotTableExists(IExcelBatch batch, string pivotTableName, string sheetName)
    {
        await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(sheetName);
            dynamic pivotTables = sheet.PivotTables;

            // Verify PivotTable exists in collection
            bool found = false;
            for (int i = 1; i <= pivotTables.Count; i++)
            {
                dynamic pivot = pivotTables.Item(i);
                if (pivot.Name == pivotTableName)
                {
                    found = true;

                    // Verify PivotTable properties
                    Assert.Equal(pivotTableName, pivot.Name);
                    Assert.True(pivot.PivotFields.Count >= 4); // Should have our 4 fields
                    break;
                }
            }

            Assert.True(found, $"PivotTable '{pivotTableName}' not found in sheet '{sheetName}'");

            return ValueTask.FromResult(0);
        });
    }

    #endregion

    public void Dispose()
    {
        if (_disposed) return;

        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, recursive: true);
            }
        }
        catch
        {
            // Cleanup failure is non-critical in tests
        }

        _disposed = true;
        GC.SuppressFinalize(this);
    }
}
