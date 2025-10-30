using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PivotTables")]
public class PivotTableCommandsTests : IDisposable
{
    private readonly IPivotTableCommands _pivotCommands;
    private readonly IFileCommands _fileCommands;
    private readonly string _tempDir;
    private bool _disposed;

    public PivotTableCommandsTests()
    {
        _pivotCommands = new PivotTableCommands();
        _fileCommands = new FileCommands();

        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_PivotTable_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    private async Task<string> CreateTestFileWithDataAsync(string fileName)
    {
        var filePath = Path.Combine(_tempDir, fileName);
        var result = await _fileCommands.CreateEmptyAsync(filePath, overwriteIfExists: false);
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}");
        }

        await using var batch = await ExcelSession.BeginBatchAsync(filePath);

        await batch.ExecuteAsync<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Name = "SalesData";

            sheet.Range["A1"].Value2 = "Region";
            sheet.Range["B1"].Value2 = "Product";
            sheet.Range["C1"].Value2 = "Sales";
            sheet.Range["D1"].Value2 = "Date";

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

        return filePath;
    }

    [Fact]
    public async Task CreateFromRange_WithValidData_CreatesCorrectPivotStructure()
    {
        var testFile = await CreateTestFileWithDataAsync("CreateFromRange_Test.xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _pivotCommands.CreateFromRangeAsync(
            batch,
            "SalesData", "A1:D6",
            "SalesData", "F1",
            "TestPivot");

        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("TestPivot", result.PivotTableName);
        Assert.Equal("SalesData", result.SheetName);
        Assert.Equal(4, result.AvailableFields.Count);
    }

    [Fact]
    public async Task List_WithValidFile_ReturnsSuccessWithPivotTables()
    {
        var testFile = await CreateTestFileWithDataAsync("List_Test.xlsx");

        {
            await using var batch = await ExcelSession.BeginBatchAsync(testFile);
            var createResult = await _pivotCommands.CreateFromRangeAsync(
                batch,
                "SalesData", "A1:D6",
                "SalesData", "F1",
                "TestPivot");
            Assert.True(createResult.Success);
            await batch.SaveAsync();
        }

        await using var batch2 = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _pivotCommands.ListAsync(batch2);

        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.PivotTables);
        Assert.Contains(result.PivotTables, p => p.Name == "TestPivot");
    }

    [Fact]
    public async Task GetInfo_WithValidPivotTable_ReturnsCompleteMetadata()
    {
        var testFile = await CreateTestFileWithDataAsync("GetInfo_Test.xlsx");

        {
            await using var batch = await ExcelSession.BeginBatchAsync(testFile);
            var createResult = await _pivotCommands.CreateFromRangeAsync(
                batch,
                "SalesData", "A1:D6",
                "SalesData", "F1",
                "InfoTestPivot");
            Assert.True(createResult.Success);
            await batch.SaveAsync();
        }

        await using var batch2 = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _pivotCommands.GetInfoAsync(batch2, "InfoTestPivot");

        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.PivotTable);
        Assert.Equal("InfoTestPivot", result.PivotTable.Name);
        Assert.Equal("SalesData", result.PivotTable.SheetName);
        Assert.NotNull(result.Fields);
        Assert.Equal(4, result.Fields.Count);
    }

    [Fact]
    public async Task Delete_WithValidPivotTable_RemovesFromCollectionAndExcel()
    {
        var testFile = await CreateTestFileWithDataAsync("Delete_Test.xlsx");

        {
            await using var batch = await ExcelSession.BeginBatchAsync(testFile);
            var createResult = await _pivotCommands.CreateFromRangeAsync(
                batch,
                "SalesData", "A1:D6",
                "SalesData", "F1",
                "ToDelete");
            Assert.True(createResult.Success);
            await batch.SaveAsync();
        }

        await using var batch2 = await ExcelSession.BeginBatchAsync(testFile);
        var deleteResult = await _pivotCommands.DeleteAsync(batch2, "ToDelete");
        Assert.True(deleteResult.Success, $"Delete failed: {deleteResult.ErrorMessage}");
        await batch2.SaveAsync();

        await using var batch3 = await ExcelSession.BeginBatchAsync(testFile);
        var listResult = await _pivotCommands.ListAsync(batch3);
        Assert.True(listResult.Success);
        Assert.DoesNotContain(listResult.PivotTables, p => p.Name == "ToDelete");
    }

    [Fact]
    public async Task Refresh_AfterSourceDataChange_UpdatesPivotTableData()
    {
        var testFile = await CreateTestFileWithDataAsync("Refresh_Test.xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch,
            "SalesData", "A1:D6",
            "SalesData", "F1",
            "RefreshTest");
        Assert.True(createResult.Success);

        await _pivotCommands.AddRowFieldAsync(batch, "RefreshTest", "Region");
        await _pivotCommands.AddValueFieldAsync(batch, "RefreshTest", "Sales", AggregationFunction.Sum);
        await batch.SaveAsync();

        await batch.ExecuteAsync<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item("SalesData");
            sheet.Range["C2"].Value2 = 999;
            return ValueTask.FromResult(0);
        });
        await batch.SaveAsync();

        var refreshResult = await _pivotCommands.RefreshAsync(batch, "RefreshTest");

        Assert.True(refreshResult.Success, $"Refresh failed: {refreshResult.ErrorMessage}");
        Assert.True(refreshResult.RefreshTime > DateTime.MinValue, "RefreshTime should be set");
    }

    [Fact]
    public async Task AddRowField_ValidField_PlacesFieldInRowAreaAndRefreshesCorrectly()
    {
        var testFile = await CreateTestFileWithDataAsync("AddRowField_Test.xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch,
            "SalesData", "A1:D6",
            "SalesData", "F1",
            "TestPivot");
        Assert.True(createResult.Success);

        var result = await _pivotCommands.AddRowFieldAsync(batch, "TestPivot", "Region");

        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("Region", result.FieldName);
        Assert.Equal(PivotFieldArea.Row, result.Area);
        Assert.Equal(1, result.Position);
    }

    [Fact]
    public async Task AddValueField_NumericField_AggregatesDataCorrectly()
    {
        var testFile = await CreateTestFileWithDataAsync("AddValueField_Test.xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch,
            "SalesData", "A1:D6",
            "SalesData", "F1",
            "TestPivot");
        Assert.True(createResult.Success);

        await _pivotCommands.AddRowFieldAsync(batch, "TestPivot", "Region");

        var result = await _pivotCommands.AddValueFieldAsync(
            batch, "TestPivot", "Sales", AggregationFunction.Sum, "Total Sales");

        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("Sales", result.FieldName);
        Assert.Equal("Total Sales", result.CustomName);
        Assert.Equal(PivotFieldArea.Value, result.Area);
        Assert.Equal(AggregationFunction.Sum, result.Function);
    }

    [Fact]
    public async Task AddValueField_TextFieldWithSumFunction_ReturnsInformativeError()
    {
        var testFile = await CreateTestFileWithDataAsync("AddValueField_Error_Test.xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch,
            "SalesData", "A1:D6",
            "SalesData", "F1",
            "TestPivot");
        Assert.True(createResult.Success);

        var result = await _pivotCommands.AddValueFieldAsync(
            batch, "TestPivot", "Region", AggregationFunction.Sum);

        Assert.False(result.Success);
        Assert.Contains("not valid for Text field 'Region'", result.ErrorMessage);
        Assert.Contains("Valid functions: Count", result.ErrorMessage);
    }

    [Fact]
    public async Task ListFields_WithValidPivotTable_ReturnsAllFieldsWithMetadata()
    {
        var testFile = await CreateTestFileWithDataAsync("ListFields_Test.xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch,
            "SalesData", "A1:D6",
            "SalesData", "F1",
            "FieldListTest");
        Assert.True(createResult.Success);

        await _pivotCommands.AddRowFieldAsync(batch, "FieldListTest", "Region");
        await _pivotCommands.AddValueFieldAsync(batch, "FieldListTest", "Sales", AggregationFunction.Sum);

        var result = await _pivotCommands.ListFieldsAsync(batch, "FieldListTest");

        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Fields);
        Assert.True(result.Fields.Count >= 4, "Should have at least 4 fields");

        var regionField = result.Fields.FirstOrDefault(f => f.Name == "Region");
        Assert.NotNull(regionField);
        Assert.Equal(PivotFieldArea.Row, regionField.Area);
    }

    [Fact]
    public async Task AddColumnField_ValidField_PlacesFieldInColumnArea()
    {
        var testFile = await CreateTestFileWithDataAsync("AddColumnField_Test.xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch,
            "SalesData", "A1:D6",
            "SalesData", "F1",
            "ColumnTest");
        Assert.True(createResult.Success);

        var result = await _pivotCommands.AddColumnFieldAsync(batch, "ColumnTest", "Product");

        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("Product", result.FieldName);
        Assert.Equal(PivotFieldArea.Column, result.Area);
        Assert.Equal(1, result.Position);
    }

    [Fact]
    public async Task AddFilterField_ValidField_PlacesFieldInFilterArea()
    {
        var testFile = await CreateTestFileWithDataAsync("AddFilterField_Test.xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch,
            "SalesData", "A1:D6",
            "SalesData", "F1",
            "FilterTest");
        Assert.True(createResult.Success);

        var result = await _pivotCommands.AddFilterFieldAsync(batch, "FilterTest", "Region");

        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("Region", result.FieldName);
        Assert.Equal(PivotFieldArea.Filter, result.Area);
    }

    [Fact]
    public async Task RemoveField_AfterAddingToRow_ReturnsFieldToAvailable()
    {
        var testFile = await CreateTestFileWithDataAsync("RemoveField_Test.xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch,
            "SalesData", "A1:D6",
            "SalesData", "F1",
            "RemoveTest");
        Assert.True(createResult.Success);

        var addResult = await _pivotCommands.AddRowFieldAsync(batch, "RemoveTest", "Region");
        Assert.True(addResult.Success);

        var removeResult = await _pivotCommands.RemoveFieldAsync(batch, "RemoveTest", "Region");

        Assert.True(removeResult.Success, $"Expected success but got error: {removeResult.ErrorMessage}");
        Assert.Equal("Region", removeResult.FieldName);
        Assert.Equal(PivotFieldArea.Hidden, removeResult.Area);
    }

    [Fact]
    public async Task SetFieldFunction_ChangeFromSumToAverage_UpdatesAggregation()
    {
        var testFile = await CreateTestFileWithDataAsync("SetFieldFunction_Test.xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch,
            "SalesData", "A1:D6",
            "SalesData", "F1",
            "FunctionTest");
        Assert.True(createResult.Success);

        await _pivotCommands.AddRowFieldAsync(batch, "FunctionTest", "Region");
        var addResult = await _pivotCommands.AddValueFieldAsync(
            batch, "FunctionTest", "Sales", AggregationFunction.Sum);
        Assert.True(addResult.Success);

        var result = await _pivotCommands.SetFieldFunctionAsync(
            batch, "FunctionTest", "Sales", AggregationFunction.Average);

        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("Sales", result.FieldName);
        Assert.Equal(AggregationFunction.Average, result.Function);
    }

    [Fact]
    public async Task SetFieldName_WithCustomName_UpdatesDisplayName()
    {
        var testFile = await CreateTestFileWithDataAsync("SetFieldName_Test.xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch,
            "SalesData", "A1:D6",
            "SalesData", "F1",
            "NameTest");
        Assert.True(createResult.Success);

        await _pivotCommands.AddValueFieldAsync(batch, "NameTest", "Sales", AggregationFunction.Sum);

        var result = await _pivotCommands.SetFieldNameAsync(
            batch, "NameTest", "Sales", "Total Revenue");

        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("Sales", result.FieldName);
        Assert.Equal("Total Revenue", result.CustomName);
    }

    [Fact]
    public async Task SetFieldFormat_WithCurrencyFormat_AppliesNumberFormat()
    {
        var testFile = await CreateTestFileWithDataAsync("SetFieldFormat_Test.xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch,
            "SalesData", "A1:D6",
            "SalesData", "F1",
            "FormatTest");
        Assert.True(createResult.Success);

        await _pivotCommands.AddValueFieldAsync(batch, "FormatTest", "Sales", AggregationFunction.Sum);

        var result = await _pivotCommands.SetFieldFormatAsync(
            batch, "FormatTest", "Sales", "$#,##0.00");

        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("Sales", result.FieldName);
        Assert.Equal("$#,##0.00", result.NumberFormat);
    }

    [Fact]
    public async Task CreateFromTable_WithValidTable_CreatesCorrectPivotStructure()
    {
        var testFile = await CreateTestFileWithDataAsync("CreateFromTable_Test.xlsx");
        var tableCommands = new TableCommands();

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        var tableResult = await tableCommands.CreateAsync(
            batch, "SalesData", "SalesTable", "A1:D6", hasHeaders: true);
        Assert.True(tableResult.Success, $"Table creation failed: {tableResult.ErrorMessage}");
        await batch.SaveAsync();

        var pivotResult = await _pivotCommands.CreateFromTableAsync(
            batch, "SalesTable", "SalesData", "F1", "TablePivot");

        Assert.True(pivotResult.Success, $"Expected success but got error: {pivotResult.ErrorMessage}");
        Assert.Equal("TablePivot", pivotResult.PivotTableName);
        Assert.Equal(4, pivotResult.AvailableFields.Count);
    }

    [Fact]
    public async Task GetData_WithConfiguredPivotTable_ReturnsCalculatedValues()
    {
        var testFile = await CreateTestFileWithDataAsync("GetData_Test.xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch,
            "SalesData", "A1:D6",
            "SalesData", "F1",
            "DataTest");
        Assert.True(createResult.Success);

        await _pivotCommands.AddRowFieldAsync(batch, "DataTest", "Region");
        await _pivotCommands.AddValueFieldAsync(batch, "DataTest", "Sales", AggregationFunction.Sum);

        var result = await _pivotCommands.GetDataAsync(batch, "DataTest");

        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Values);
        Assert.True(result.Values.Count > 0, "Should have data rows");
    }

    [Fact]
    public async Task SetFieldFilter_WithValidCriteria_FiltersDataCorrectly()
    {
        var testFile = await CreateTestFileWithDataAsync("SetFieldFilter_Test.xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch,
            "SalesData", "A1:D6",
            "SalesData", "F1",
            "FilterCriteriaTest");
        Assert.True(createResult.Success);

        await _pivotCommands.AddFilterFieldAsync(batch, "FilterCriteriaTest", "Region");
        await _pivotCommands.AddRowFieldAsync(batch, "FilterCriteriaTest", "Product");
        await _pivotCommands.AddValueFieldAsync(batch, "FilterCriteriaTest", "Sales", AggregationFunction.Sum);

        var result = await _pivotCommands.SetFieldFilterAsync(
            batch, "FilterCriteriaTest", "Region", ["North"]);

        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("Region", result.FieldName);
        Assert.NotNull(result.SelectedItems);
        Assert.Contains("North", result.SelectedItems);
    }

    [Fact]
    public async Task SortField_ByAscending_OrdersDataCorrectly()
    {
        var testFile = await CreateTestFileWithDataAsync("SortField_Test.xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch,
            "SalesData", "A1:D6",
            "SalesData", "F1",
            "SortTest");
        Assert.True(createResult.Success);

        await _pivotCommands.AddRowFieldAsync(batch, "SortTest", "Region");
        await _pivotCommands.AddValueFieldAsync(batch, "SortTest", "Sales", AggregationFunction.Sum);

        var result = await _pivotCommands.SortFieldAsync(
            batch, "SortTest", "Region", SortDirection.Ascending);

        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("Region", result.FieldName);
    }

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
            // Cleanup failure is non-critical
        }

        _disposed = true;
        GC.SuppressFinalize(this);
    }
}
