// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Core.Tests.Diagnostics;

/// <summary>
/// Diagnostic tests to understand Excel's native behavior regarding RefreshTable().
/// These tests use RAW Excel COM API without our abstraction layer to determine:
/// 1. Which operations require RefreshTable() to take effect
/// 2. Which operations work immediately without RefreshTable()
/// 
/// Purpose: Inform optimization decisions by understanding Excel's true behavior.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Layer", "Diagnostics")]
[Trait("Speed", "Slow")]
[Trait("RequiresExcel", "true")]
[Trait("RunType", "OnDemand")]
public class PivotTableRefreshBehaviorTests : IClassFixture<TempDirectoryFixture>, IDisposable
{
    private readonly string _tempDir;
    private readonly ITestOutputHelper _output;
    private dynamic? _excel;
    private dynamic? _workbook;
    private readonly string _testFile;

    // Excel constants
    private const int xlRowField = 1;
    private const int xlColumnField = 2;
    private const int xlDataField = 4;
    private const int xlPageField = 3;
    private const int xlHidden = 0;
    private const int xlSum = -4157;
    private const int xlCount = -4112;
    private const int xlAverage = -4106;
    private const int xlDatabase = 1;

    public PivotTableRefreshBehaviorTests(TempDirectoryFixture fixture, ITestOutputHelper output)
    {
        _tempDir = fixture.TempDir;
        _output = output;
        _testFile = Path.Combine(_tempDir, $"RefreshBehavior_{Guid.NewGuid():N}.xlsx");

        // Create Excel instance directly (no abstraction)
        var excelType = Type.GetTypeFromProgID("Excel.Application");
        _excel = Activator.CreateInstance(excelType!);
        _excel.Visible = false;
        _excel.DisplayAlerts = false;

        // Create workbook with test data
        _workbook = _excel.Workbooks.Add();
        SetupTestData();
    }

    private void SetupTestData()
    {
        dynamic sheet = _workbook.Worksheets.Item(1);
        sheet.Name = "Data";

        // Create simple sales data
        sheet.Range["A1"].Value2 = "Region";
        sheet.Range["B1"].Value2 = "Product";
        sheet.Range["C1"].Value2 = "Sales";
        sheet.Range["D1"].Value2 = "Quantity";

        sheet.Range["A2"].Value2 = "North";
        sheet.Range["B2"].Value2 = "Widget";
        sheet.Range["C2"].Value2 = 100;
        sheet.Range["D2"].Value2 = 10;

        sheet.Range["A3"].Value2 = "South";
        sheet.Range["B3"].Value2 = "Gadget";
        sheet.Range["C3"].Value2 = 200;
        sheet.Range["D3"].Value2 = 20;

        sheet.Range["A4"].Value2 = "North";
        sheet.Range["B4"].Value2 = "Gadget";
        sheet.Range["C4"].Value2 = 150;
        sheet.Range["D4"].Value2 = 15;

        sheet.Range["A5"].Value2 = "South";
        sheet.Range["B5"].Value2 = "Widget";
        sheet.Range["C5"].Value2 = 175;
        sheet.Range["D5"].Value2 = 17;
    }

    private dynamic CreatePivotTable(string name)
    {
        dynamic dataSheet = _workbook.Worksheets.Item("Data");
        dynamic sourceRange = dataSheet.Range["A1:D5"];

        // Add a new sheet for the PivotTable
        dynamic pivotSheet = _workbook.Worksheets.Add();
        pivotSheet.Name = $"Pivot_{name}";

        // Create PivotCache and PivotTable
        dynamic pivotCache = _workbook.PivotCaches().Create(xlDatabase, sourceRange);
        dynamic pivotTable = pivotCache.CreatePivotTable(pivotSheet.Range["A3"], name);

        return pivotTable;
    }

    [Fact]
    public void AddRowField_WithoutRefresh_VerifyOrientationTakesEffect()
    {
        // Arrange
        var pivot = CreatePivotTable("TestPivot1");
        dynamic field = pivot.PivotFields("Region");

        // Act - Set orientation WITHOUT calling RefreshTable()
        field.Orientation = xlRowField;
        // NO pivot.RefreshTable();

        // Assert - Check if orientation took effect
        int actualOrientation = Convert.ToInt32(field.Orientation);
        _output.WriteLine($"AddRowField WITHOUT RefreshTable:");
        _output.WriteLine($"  Expected Orientation: {xlRowField} (xlRowField)");
        _output.WriteLine($"  Actual Orientation:   {actualOrientation}");
        _output.WriteLine($"  Match: {actualOrientation == xlRowField}");

        // Check if data appears in the PivotTable
        dynamic pivotSheet = _workbook.Worksheets.Item($"Pivot_TestPivot1");
        string cellA4Value = pivotSheet.Range["A4"].Value2?.ToString() ?? "(empty)";
        _output.WriteLine($"  Cell A4 value: {cellA4Value}");

        Assert.Equal(xlRowField, actualOrientation);
    }

    [Fact]
    public void AddRowField_WithRefresh_VerifyOrientationTakesEffect()
    {
        // Arrange
        var pivot = CreatePivotTable("TestPivot2");
        dynamic field = pivot.PivotFields("Region");

        // Act - Set orientation WITH RefreshTable()
        field.Orientation = xlRowField;
        pivot.RefreshTable();

        // Assert
        int actualOrientation = Convert.ToInt32(field.Orientation);
        _output.WriteLine($"AddRowField WITH RefreshTable:");
        _output.WriteLine($"  Expected Orientation: {xlRowField} (xlRowField)");
        _output.WriteLine($"  Actual Orientation:   {actualOrientation}");
        _output.WriteLine($"  Match: {actualOrientation == xlRowField}");

        dynamic pivotSheet = _workbook.Worksheets.Item($"Pivot_TestPivot2");
        string cellA4Value = pivotSheet.Range["A4"].Value2?.ToString() ?? "(empty)";
        _output.WriteLine($"  Cell A4 value: {cellA4Value}");

        Assert.Equal(xlRowField, actualOrientation);
    }

    [Fact]
    public void AddValueField_WithoutRefresh_VerifyDataAppears()
    {
        // Arrange
        var pivot = CreatePivotTable("TestPivot3");
        dynamic rowField = pivot.PivotFields("Region");
        rowField.Orientation = xlRowField;

        dynamic valueField = pivot.PivotFields("Sales");

        // Act - Add value field WITHOUT RefreshTable()
        valueField.Orientation = xlDataField;
        valueField.Function = xlSum;
        // NO pivot.RefreshTable();

        // Assert
        int actualOrientation = Convert.ToInt32(valueField.Orientation);
        _output.WriteLine($"AddValueField WITHOUT RefreshTable:");
        _output.WriteLine($"  Expected Orientation: {xlDataField} (xlDataField)");
        _output.WriteLine($"  Actual Orientation:   {actualOrientation}");

        // Check if sum values appear
        dynamic pivotSheet = _workbook.Worksheets.Item($"Pivot_TestPivot3");
        var cellB4 = pivotSheet.Range["B4"].Value2;
        _output.WriteLine($"  Cell B4 (should have sum): {cellB4}");

        Assert.Equal(xlDataField, actualOrientation);
    }

    [Fact]
    public void RemoveField_WithoutRefresh_VerifyFieldHidden()
    {
        // Arrange
        var pivot = CreatePivotTable("TestPivot4");
        dynamic field = pivot.PivotFields("Region");
        field.Orientation = xlRowField;
        pivot.RefreshTable(); // Initial setup with refresh

        // Act - Remove field WITHOUT RefreshTable()
        field.Orientation = xlHidden;
        // NO pivot.RefreshTable();

        // Assert
        int actualOrientation = Convert.ToInt32(field.Orientation);
        _output.WriteLine($"RemoveField WITHOUT RefreshTable:");
        _output.WriteLine($"  Expected Orientation: {xlHidden} (xlHidden)");
        _output.WriteLine($"  Actual Orientation:   {actualOrientation}");
        _output.WriteLine($"  Match: {actualOrientation == xlHidden}");

        Assert.Equal(xlHidden, actualOrientation);
    }

    [Fact]
    public void SetNumberFormat_WithoutRefresh_VerifyFormatApplied()
    {
        // Arrange
        var pivot = CreatePivotTable("TestPivot5");
        dynamic rowField = pivot.PivotFields("Region");
        rowField.Orientation = xlRowField;

        dynamic valueField = pivot.PivotFields("Sales");
        valueField.Orientation = xlDataField;
        valueField.Function = xlSum;
        pivot.RefreshTable(); // Initial setup

        // Act - Set number format WITHOUT RefreshTable()
        string formatBefore = valueField.NumberFormat?.ToString() ?? "(none)";
        valueField.NumberFormat = "$#,##0.00";
        // NO pivot.RefreshTable();

        // Assert
        string formatAfter = valueField.NumberFormat?.ToString() ?? "(none)";
        _output.WriteLine($"SetNumberFormat WITHOUT RefreshTable:");
        _output.WriteLine($"  Format Before: {formatBefore}");
        _output.WriteLine($"  Format After:  {formatAfter}");
        _output.WriteLine($"  Match Expected: {formatAfter == "$#,##0.00"}");

        Assert.Equal("$#,##0.00", formatAfter);
    }

    [Fact]
    public void ChangeFunction_WithoutRefresh_VerifyFunctionChanged()
    {
        // Arrange
        var pivot = CreatePivotTable("TestPivot6");
        dynamic rowField = pivot.PivotFields("Region");
        rowField.Orientation = xlRowField;

        dynamic valueField = pivot.PivotFields("Sales");
        valueField.Orientation = xlDataField;
        valueField.Function = xlSum;
        pivot.RefreshTable(); // Initial setup

        // Capture value with SUM
        dynamic pivotSheet = _workbook.Worksheets.Item($"Pivot_TestPivot6");
        var sumValue = pivotSheet.Range["B4"].Value2;
        _output.WriteLine($"Value with SUM: {sumValue}");

        // Act - Change function WITHOUT RefreshTable()
        valueField.Function = xlCount;
        // NO pivot.RefreshTable();

        // Assert
        int actualFunction = Convert.ToInt32(valueField.Function);
        var countValue = pivotSheet.Range["B4"].Value2;

        _output.WriteLine($"ChangeFunction WITHOUT RefreshTable:");
        _output.WriteLine($"  Expected Function: {xlCount} (xlCount)");
        _output.WriteLine($"  Actual Function:   {actualFunction}");
        _output.WriteLine($"  Value after COUNT: {countValue}");
        _output.WriteLine($"  Values different:  {!Equals(sumValue, countValue)}");

        Assert.Equal(xlCount, actualFunction);
    }

    [Fact]
    public void MultipleOperations_WithoutRefresh_VerifyAllApplied()
    {
        // Arrange
        var pivot = CreatePivotTable("TestPivot7");

        // Act - Multiple operations WITHOUT RefreshTable() between them
        dynamic regionField = pivot.PivotFields("Region");
        regionField.Orientation = xlRowField;

        dynamic productField = pivot.PivotFields("Product");
        productField.Orientation = xlColumnField;

        dynamic salesField = pivot.PivotFields("Sales");
        salesField.Orientation = xlDataField;
        salesField.Function = xlSum;
        salesField.NumberFormat = "$#,##0";

        // NO RefreshTable() at all

        // Assert
        _output.WriteLine($"Multiple Operations WITHOUT any RefreshTable:");
        _output.WriteLine($"  Region Orientation:  {regionField.Orientation} (expected {xlRowField})");
        _output.WriteLine($"  Product Orientation: {productField.Orientation} (expected {xlColumnField})");
        _output.WriteLine($"  Sales Orientation:   {salesField.Orientation} (expected {xlDataField})");
        _output.WriteLine($"  Sales NumberFormat:  {salesField.NumberFormat}");

        Assert.Equal(xlRowField, Convert.ToInt32(regionField.Orientation));
        Assert.Equal(xlColumnField, Convert.ToInt32(productField.Orientation));
        Assert.Equal(xlDataField, Convert.ToInt32(salesField.Orientation));
    }

    [Fact]
    public void Persistence_WithoutRefresh_VerifySaveAndReopen()
    {
        // Arrange
        var pivot = CreatePivotTable("TestPivot8");

        dynamic regionField = pivot.PivotFields("Region");
        regionField.Orientation = xlRowField;

        dynamic salesField = pivot.PivotFields("Sales");
        salesField.Orientation = xlDataField;
        salesField.NumberFormat = "0.00%";

        // NO RefreshTable()

        // Save and close
        _workbook.SaveAs(_testFile);
        _workbook.Close(false);
        Marshal.ReleaseComObject(_workbook);
        _workbook = null;

        // Reopen
        _workbook = _excel.Workbooks.Open(_testFile);
        dynamic reopenedPivot = _workbook.Worksheets.Item("Pivot_TestPivot8").PivotTables("TestPivot8");

        // Assert - Check if settings persisted
        dynamic reopenedRegion = reopenedPivot.PivotFields("Region");
        dynamic reopenedSales = reopenedPivot.PivotFields("Sales");

        _output.WriteLine($"Persistence WITHOUT RefreshTable:");
        _output.WriteLine($"  Region Orientation after reopen: {reopenedRegion.Orientation}");
        _output.WriteLine($"  Sales Orientation after reopen:  {reopenedSales.Orientation}");
        _output.WriteLine($"  Sales NumberFormat after reopen: {reopenedSales.NumberFormat}");

        // The question: do these persist without RefreshTable()?
        // FINDING: Row field orientation DOES persist, but value field orientation does NOT
        // This proves RefreshTable() IS needed for value fields to persist properly
        Assert.Equal(xlRowField, Convert.ToInt32(reopenedRegion.Orientation));
        // Value field did NOT persist - this is expected without RefreshTable()
        Assert.Equal(xlHidden, Convert.ToInt32(reopenedSales.Orientation)); // 0 = Hidden, not 4 = DataField
    }

    [Fact]
    public void Persistence_WithRefresh_VerifySaveAndReopen()
    {
        // Arrange
        var pivot = CreatePivotTable("TestPivot9");

        dynamic regionField = pivot.PivotFields("Region");
        regionField.Orientation = xlRowField;

        dynamic salesField = pivot.PivotFields("Sales");
        salesField.Orientation = xlDataField;

        // CRITICAL: RefreshTable after structure changes, BEFORE setting visual properties
        pivot.RefreshTable();

        // Now set visual properties
        salesField.NumberFormat = "0.00%";

        // Save and close
        _workbook.SaveAs(_testFile);
        _workbook.Close(false);
        Marshal.ReleaseComObject(_workbook);
        _workbook = null;

        // Reopen
        _workbook = _excel.Workbooks.Open(_testFile);
        dynamic reopenedPivot = _workbook.Worksheets.Item("Pivot_TestPivot9").PivotTables("TestPivot9");

        // Assert - Check if settings persisted
        dynamic reopenedRegion = reopenedPivot.PivotFields("Region");
        dynamic reopenedSales = reopenedPivot.PivotFields("Sum of Sales");  // Name changes after refresh!

        _output.WriteLine($"Persistence WITH RefreshTable (before visual props):");
        _output.WriteLine($"  Region Orientation after reopen: {reopenedRegion.Orientation}");
        _output.WriteLine($"  Sum of Sales Orientation after reopen:  {reopenedSales.Orientation}");
        _output.WriteLine($"  Sum of Sales NumberFormat after reopen: {reopenedSales.NumberFormat}");

        Assert.Equal(xlRowField, Convert.ToInt32(reopenedRegion.Orientation));
        Assert.Equal(xlDataField, Convert.ToInt32(reopenedSales.Orientation));
        Assert.Equal("0.00%", reopenedSales.NumberFormat?.ToString());
    }

    [Fact]
    public void FunctionChange_WithoutRefresh_VerifyPersistence()
    {
        // Arrange - Create PivotTable with value field using SUM
        var pivot = CreatePivotTable("TestPivot10");
        dynamic rowField = pivot.PivotFields("Region");
        rowField.Orientation = xlRowField;

        dynamic valueField = pivot.PivotFields("Sales");
        valueField.Orientation = xlDataField;
        valueField.Function = xlSum;
        pivot.RefreshTable(); // Initial setup - structure must be refreshed

        // Act - Change function WITHOUT RefreshTable()
        dynamic dataField = pivot.DataFields.Item(1);  // Get from DataFields collection
        dataField.Function = xlAverage;
        // NO pivot.RefreshTable();

        // Save and close
        _workbook.SaveAs(_testFile);
        _workbook.Close(false);
        Marshal.ReleaseComObject(_workbook);
        _workbook = null;

        // Reopen
        _workbook = _excel.Workbooks.Open(_testFile);
        dynamic reopenedPivot = _workbook.Worksheets.Item("Pivot_TestPivot10").PivotTables("TestPivot10");
        dynamic reopenedDataField = reopenedPivot.DataFields.Item(1);
        int reopenedFunction = Convert.ToInt32(reopenedDataField.Function);

        _output.WriteLine($"Function Change WITHOUT RefreshTable - Persistence Test:");
        _output.WriteLine($"  Expected Function: {xlAverage} (xlAverage)");
        _output.WriteLine($"  Actual Function:   {reopenedFunction}");
        _output.WriteLine($"  Persisted: {reopenedFunction == xlAverage}");

        // Question: Does function change persist without RefreshTable()?
        // If this fails, RefreshTable() IS required after SetFieldFunction
        Assert.Equal(xlAverage, reopenedFunction);
    }

    [Fact]
    public void Filter_WithoutRefresh_VerifyPersistence()
    {
        // Arrange - Create PivotTable with row field
        var pivot = CreatePivotTable("TestPivot11");
        dynamic rowField = pivot.PivotFields("Region");
        rowField.Orientation = xlRowField;
        pivot.RefreshTable(); // Initial setup

        // Act - Apply filter WITHOUT RefreshTable()
        // Set only "North" visible
        dynamic items = rowField.PivotItems();
        for (int i = 1; i <= items.Count; i++)
        {
            dynamic item = items.Item(i);
            string itemName = item.Name?.ToString() ?? "";
            item.Visible = (itemName == "North");
        }
        // NO pivot.RefreshTable();

        // Save and close
        _workbook.SaveAs(_testFile);
        _workbook.Close(false);
        Marshal.ReleaseComObject(_workbook);
        _workbook = null;

        // Reopen
        _workbook = _excel.Workbooks.Open(_testFile);
        dynamic reopenedPivot = _workbook.Worksheets.Item("Pivot_TestPivot11").PivotTables("TestPivot11");
        dynamic reopenedField = reopenedPivot.PivotFields("Region");
        dynamic reopenedItems = reopenedField.PivotItems();

        int visibleCount = 0;
        string visibleItemName = "";
        for (int i = 1; i <= reopenedItems.Count; i++)
        {
            dynamic item = reopenedItems.Item(i);
            if (item.Visible)
            {
                visibleCount++;
                visibleItemName = item.Name?.ToString() ?? "";
            }
        }

        _output.WriteLine($"Filter WITHOUT RefreshTable - Persistence Test:");
        _output.WriteLine($"  Expected: Only 'North' visible");
        _output.WriteLine($"  Visible count: {visibleCount}");
        _output.WriteLine($"  Visible item: {visibleItemName}");

        // Question: Does filter persist without RefreshTable()?
        // If this fails, RefreshTable() IS required after SetFieldFilter
        Assert.Equal(1, visibleCount);
        Assert.Equal("North", visibleItemName);
    }

    public void Dispose()
    {
        GC.SuppressFinalize(this);
        try
        {
            if (_workbook != null)
            {
                _workbook.Close(false);
                Marshal.ReleaseComObject(_workbook);
            }
        }
#pragma warning disable CA1031 // Intentional: cleanup code must not throw
        catch (Exception) { /* Ignore cleanup errors */ }
#pragma warning restore CA1031

        try
        {
            if (_excel != null)
            {
                _excel.Quit();
                Marshal.ReleaseComObject(_excel);
            }
        }
#pragma warning disable CA1031 // Intentional: cleanup code must not throw
        catch (Exception) { /* Ignore cleanup errors */ }
#pragma warning restore CA1031

        // Clean up test file
        try
        {
            if (File.Exists(_testFile))
                File.Delete(_testFile);
        }
#pragma warning disable CA1031 // Intentional: cleanup code must not throw
        catch (Exception) { /* Ignore file cleanup errors */ }
#pragma warning restore CA1031

        GC.Collect();
        GC.WaitForPendingFinalizers();
    }
}
