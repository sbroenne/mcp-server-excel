// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Diagnostics.Tests.Diagnostics;

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
[Collection("Sequential")]
[Trait("RequiresExcel", "true")]
[Trait("RunType", "OnDemand")]
public class PivotTableRefreshBehaviorTests : IClassFixture<TempDirectoryFixture>, IDisposable
{
    private readonly string _tempDir;
    private readonly ITestOutputHelper _output;
    private Excel.Application? _excel;
    private Excel.Workbook? _workbook;
    private readonly string _testFile;

    // Excel constants
    private const int xlRowField = 1;
    private const int xlColumnField = 2;
    private const int xlDataField = 4;
    private const int xlHidden = 0;
    private const int xlSum = -4157;
    private const int xlCount = -4112;
    private const int xlAverage = -4106;

    public PivotTableRefreshBehaviorTests(TempDirectoryFixture fixture, ITestOutputHelper output)
    {
        _tempDir = fixture.TempDir;
        _output = output;
        _testFile = Path.Combine(_tempDir, $"RefreshBehavior_{Guid.NewGuid():N}.xlsx");

        // Create Excel instance directly (no abstraction)
        var excelType = Type.GetTypeFromProgID("Excel.Application");
        Assert.NotNull(excelType);
        try
        {
            _excel = (Excel.Application?)Activator.CreateInstance(excelType);
            Assert.NotNull(_excel);
            _excel.Visible = false;
            _excel.DisplayAlerts = false;
            _workbook = DiagnosticExcelObjects.AddWorkbook(_excel);
            SetupTestData();
        }
        catch (Exception failure)
        {
            try
            {
                Dispose();
            }
            catch (Exception cleanup)
            {
                throw new AggregateException(failure, cleanup);
            }
            throw;
        }
    }

    private void SetupTestData()
    {
        Assert.NotNull(_workbook);
        Excel.Worksheet? sheet = null;
        Excel.Range? data = null;
        try
        {
            sheet = DiagnosticExcelObjects.GetWorksheet(_workbook, 1);
            sheet.Name = "Data";
            data = sheet.Range["A1:D5"];
            data.Value2 = new object[,]
            {
                { "Region", "Product", "Sales", "Quantity" },
                { "North", "Widget", 100, 10 },
                { "South", "Gadget", 200, 20 },
                { "North", "Gadget", 150, 15 },
                { "South", "Widget", 175, 17 }
            };
        }
        finally
        {
            ComUtilities.Release(ref data);
            ComUtilities.Release(ref sheet);
        }
    }

    private Excel.PivotTable CreatePivotTable(string name)
    {
        Assert.NotNull(_workbook);
        Excel.Worksheet? dataSheet = null;
        Excel.Range? sourceRange = null;
        Excel.Worksheet? pivotSheet = null;
        Excel.Range? destination = null;
        Excel.PivotCaches? caches = null;
        Excel.PivotCache? cache = null;
        try
        {
            dataSheet = DiagnosticExcelObjects.GetWorksheet(_workbook, "Data");
            sourceRange = dataSheet.Range["A1:D5"];
            pivotSheet = DiagnosticExcelObjects.AddWorksheet(_workbook);
            pivotSheet.Name = $"Pivot_{name}";
            destination = pivotSheet.Range["A3"];
            caches = _workbook.PivotCaches();
            cache = caches.Create(Excel.XlPivotTableSourceType.xlDatabase, sourceRange);
            return cache.CreatePivotTable(destination, name);
        }
        finally
        {
            ComUtilities.Release(ref cache);
            ComUtilities.Release(ref caches);
            ComUtilities.Release(ref destination);
            ComUtilities.Release(ref pivotSheet);
            ComUtilities.Release(ref sourceRange);
            ComUtilities.Release(ref dataSheet);
        }
    }

    [Fact]
    public void RemoveField_WithoutRefresh_VerifyFieldHidden()
    {
        Excel.PivotTable? pivot = null;
        Excel.PivotField? field = null;
        try
        {
            // Arrange
            pivot = CreatePivotTable("TestPivot4");
            field = (Excel.PivotField)pivot.PivotFields("Region");
            field.Orientation = (Excel.XlPivotFieldOrientation)xlRowField;
            pivot.RefreshTable(); // Initial setup with refresh

            // Act - Remove field WITHOUT RefreshTable()
            field.Orientation = (Excel.XlPivotFieldOrientation)xlHidden;
            // NO pivot.RefreshTable();

            // Assert
            int actualOrientation = (int)field.Orientation;
            _output.WriteLine($"RemoveField WITHOUT RefreshTable:");
            _output.WriteLine($"  Expected Orientation: {xlHidden} (xlHidden)");
            _output.WriteLine($"  Actual Orientation:   {actualOrientation}");
            _output.WriteLine($"  Match: {actualOrientation == xlHidden}");

            Assert.Equal(xlHidden, actualOrientation);
        }
        finally
        {
            ComUtilities.Release(ref field);
            ComUtilities.Release(ref pivot);
        }
    }

    [Fact]
    public void SetNumberFormat_WithoutRefresh_VerifyFormatApplied()
    {
        Excel.PivotTable? pivot = null;
        Excel.PivotField? rowField = null;
        Excel.PivotField? valueField = null;
        try
        {
            // Arrange
            pivot = CreatePivotTable("TestPivot5");
            rowField = (Excel.PivotField)pivot.PivotFields("Region");
            rowField.Orientation = (Excel.XlPivotFieldOrientation)xlRowField;

            valueField = (Excel.PivotField)pivot.PivotFields("Sales");
            valueField.Orientation = (Excel.XlPivotFieldOrientation)xlDataField;
            valueField.Function = (Excel.XlConsolidationFunction)xlSum;
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
        finally
        {
            ComUtilities.Release(ref valueField);
            ComUtilities.Release(ref rowField);
            ComUtilities.Release(ref pivot);
        }
    }

    [Fact]
    public void ChangeFunction_WithoutRefresh_VerifyFunctionChanged()
    {
        Excel.PivotTable? pivot = null;
        Excel.PivotField? rowField = null;
        Excel.PivotField? valueField = null;
        Excel.Worksheet? pivotSheet = null;
        try
        {
            Assert.NotNull(_workbook);
            // Arrange
            pivot = CreatePivotTable("TestPivot6");
            rowField = (Excel.PivotField)pivot.PivotFields("Region");
            rowField.Orientation = (Excel.XlPivotFieldOrientation)xlRowField;

            valueField = (Excel.PivotField)pivot.PivotFields("Sales");
            valueField.Orientation = (Excel.XlPivotFieldOrientation)xlDataField;
            valueField.Function = (Excel.XlConsolidationFunction)xlSum;
            pivot.RefreshTable(); // Initial setup

            // Capture value with SUM
            pivotSheet = DiagnosticExcelObjects.GetWorksheet(_workbook, $"Pivot_TestPivot6");
            var sumValue = DiagnosticExcelObjects.ReadCell(pivotSheet, "B4");
            _output.WriteLine($"Value with SUM: {sumValue}");

            // Act - Change function WITHOUT RefreshTable()
            valueField.Function = (Excel.XlConsolidationFunction)xlCount;
            // NO pivot.RefreshTable();

            // Assert
            int actualFunction = (int)valueField.Function;
            var countValue = DiagnosticExcelObjects.ReadCell(pivotSheet, "B4");

            _output.WriteLine($"ChangeFunction WITHOUT RefreshTable:");
            _output.WriteLine($"  Expected Function: {xlCount} (xlCount)");
            _output.WriteLine($"  Actual Function:   {actualFunction}");
            _output.WriteLine($"  Value after COUNT: {countValue}");
            _output.WriteLine($"  Values different:  {!Equals(sumValue, countValue)}");

            Assert.Equal(xlCount, actualFunction);
        }
        finally
        {
            ComUtilities.Release(ref pivotSheet);
            ComUtilities.Release(ref valueField);
            ComUtilities.Release(ref rowField);
            ComUtilities.Release(ref pivot);
        }
    }

    [Fact]
    public void MultipleOperations_WithoutRefresh_VerifyAllApplied()
    {
        Excel.PivotTable? pivot = null;
        Excel.PivotField? regionField = null;
        Excel.PivotField? productField = null;
        Excel.PivotField? salesField = null;
        try
        {
            // Arrange
            pivot = CreatePivotTable("TestPivot7");

            // Act - Multiple operations WITHOUT RefreshTable() between them
            regionField = (Excel.PivotField)pivot.PivotFields("Region");
            regionField.Orientation = (Excel.XlPivotFieldOrientation)xlRowField;

            productField = (Excel.PivotField)pivot.PivotFields("Product");
            productField.Orientation = (Excel.XlPivotFieldOrientation)xlColumnField;

            salesField = (Excel.PivotField)pivot.PivotFields("Sales");
            salesField.Orientation = (Excel.XlPivotFieldOrientation)xlDataField;
            salesField.Function = (Excel.XlConsolidationFunction)xlSum;
            salesField.NumberFormat = "$#,##0";

            // NO RefreshTable() at all

            // Assert
            _output.WriteLine($"Multiple Operations WITHOUT any RefreshTable:");
            _output.WriteLine($"  Region Orientation:  {regionField.Orientation} (expected {xlRowField})");
            _output.WriteLine($"  Product Orientation: {productField.Orientation} (expected {xlColumnField})");
            _output.WriteLine($"  Sales Orientation:   {salesField.Orientation} (expected {xlDataField})");
            _output.WriteLine($"  Sales NumberFormat:  {salesField.NumberFormat}");

            Assert.Equal(xlRowField, (int)regionField.Orientation);
            Assert.Equal(xlColumnField, (int)productField.Orientation);
            Assert.Equal(xlDataField, (int)salesField.Orientation);
        }
        finally
        {
            ComUtilities.Release(ref salesField);
            ComUtilities.Release(ref productField);
            ComUtilities.Release(ref regionField);
            ComUtilities.Release(ref pivot);
        }
    }

    [Fact]
    public void Persistence_WithoutRefresh_VerifySaveAndReopen()
    {
        Excel.PivotTable? pivot = null;
        Excel.PivotField? regionField = null;
        Excel.PivotField? salesField = null;
        Excel.PivotTable? reopenedPivot = null;
        Excel.PivotField? reopenedRegion = null;
        Excel.PivotField? reopenedSales = null;
        try
        {
            Assert.NotNull(_excel);
            Assert.NotNull(_workbook);
            // Arrange
            pivot = CreatePivotTable("TestPivot8");

            regionField = (Excel.PivotField)pivot.PivotFields("Region");
            regionField.Orientation = (Excel.XlPivotFieldOrientation)xlRowField;

            salesField = (Excel.PivotField)pivot.PivotFields("Sales");
            salesField.Orientation = (Excel.XlPivotFieldOrientation)xlDataField;
            salesField.NumberFormat = "0.00%";

            // NO RefreshTable()

            // Save and close
            _workbook.SaveAs(_testFile);
            _workbook.Close(false);
            Marshal.ReleaseComObject(_workbook);
            _workbook = null;

            // Reopen
            _workbook = DiagnosticExcelObjects.OpenWorkbook(_excel, _testFile);
            reopenedPivot = DiagnosticExcelObjects.GetPivotTable(_workbook, "Pivot_TestPivot8", "TestPivot8");

            // Assert - Check if settings persisted
            reopenedRegion = (Excel.PivotField)reopenedPivot.PivotFields("Region");
            reopenedSales = (Excel.PivotField)reopenedPivot.PivotFields("Sales");

            _output.WriteLine($"Persistence WITHOUT RefreshTable:");
            _output.WriteLine($"  Region Orientation after reopen: {reopenedRegion.Orientation}");
            _output.WriteLine($"  Sales Orientation after reopen:  {reopenedSales.Orientation}");
            _output.WriteLine($"  Sales NumberFormat after reopen: {reopenedSales.NumberFormat}");

            // The question: do these persist without RefreshTable()?
            // FINDING: Row field orientation DOES persist, but value field orientation does NOT
            // This proves RefreshTable() IS needed for value fields to persist properly
            Assert.Equal(xlRowField, (int)reopenedRegion.Orientation);
            // Value field did NOT persist - this is expected without RefreshTable()
            Assert.Equal(xlHidden, (int)reopenedSales.Orientation); // 0 = Hidden, not 4 = DataField
        }
        finally
        {
            ComUtilities.Release(ref reopenedSales);
            ComUtilities.Release(ref reopenedRegion);
            ComUtilities.Release(ref reopenedPivot);
            ComUtilities.Release(ref salesField);
            ComUtilities.Release(ref regionField);
            ComUtilities.Release(ref pivot);
        }
    }

    [Fact]
    public void Persistence_WithRefresh_VerifySaveAndReopen()
    {
        Excel.PivotTable? pivot = null;
        Excel.PivotField? regionField = null;
        Excel.PivotField? salesField = null;
        Excel.PivotTable? reopenedPivot = null;
        Excel.PivotField? reopenedRegion = null;
        try
        {
            Assert.NotNull(_excel);
            Assert.NotNull(_workbook);
            // Arrange
            pivot = CreatePivotTable("TestPivot9");

            regionField = (Excel.PivotField)pivot.PivotFields("Region");
            regionField.Orientation = (Excel.XlPivotFieldOrientation)xlRowField;

            salesField = (Excel.PivotField)pivot.PivotFields("Sales");
            salesField.Orientation = (Excel.XlPivotFieldOrientation)xlDataField;

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
            _workbook = DiagnosticExcelObjects.OpenWorkbook(_excel, _testFile);
            reopenedPivot = DiagnosticExcelObjects.GetPivotTable(_workbook, "Pivot_TestPivot9", "TestPivot9");

            // Assert - Check if settings persisted
            reopenedRegion = (Excel.PivotField)reopenedPivot.PivotFields("Region");
            Excel.PivotField reopenedSales = (Excel.PivotField)reopenedPivot.PivotFields("Sum of Sales");  // Name changes after refresh!

            _output.WriteLine($"Persistence WITH RefreshTable (before visual props):");
            _output.WriteLine($"  Region Orientation after reopen: {reopenedRegion.Orientation}");
            _output.WriteLine($"  Sum of Sales Orientation after reopen:  {reopenedSales.Orientation}");
            _output.WriteLine($"  Sum of Sales NumberFormat after reopen: {reopenedSales.NumberFormat}");

            Assert.Equal(xlRowField, (int)reopenedRegion.Orientation);
            Assert.Equal(xlDataField, (int)reopenedSales.Orientation);
            Assert.Equal("0.00%", reopenedSales.NumberFormat?.ToString());
        }
        finally
        {
            ComUtilities.Release(ref reopenedRegion);
            ComUtilities.Release(ref reopenedPivot);
            ComUtilities.Release(ref salesField);
            ComUtilities.Release(ref regionField);
            ComUtilities.Release(ref pivot);
        }
    }

    [Fact]
    public void FunctionChange_WithoutRefresh_VerifyPersistence()
    {
        Excel.PivotTable? pivot = null;
        Excel.PivotField? rowField = null;
        Excel.PivotField? valueField = null;
        Excel.PivotTable? reopenedPivot = null;
        Excel.PivotField? reopenedDataField = null;
        try
        {
            Assert.NotNull(_excel);
            Assert.NotNull(_workbook);
            // Arrange - Create PivotTable with value field using SUM
            pivot = CreatePivotTable("TestPivot10");
            rowField = (Excel.PivotField)pivot.PivotFields("Region");
            rowField.Orientation = (Excel.XlPivotFieldOrientation)xlRowField;

            valueField = (Excel.PivotField)pivot.PivotFields("Sales");
            valueField.Orientation = (Excel.XlPivotFieldOrientation)xlDataField;
            valueField.Function = (Excel.XlConsolidationFunction)xlSum;
            pivot.RefreshTable(); // Initial setup - structure must be refreshed

            // Act - Change function WITHOUT RefreshTable()
            Excel.PivotField dataField = DiagnosticExcelObjects.GetDataField(pivot, 1);  // Get from DataFields collection
            dataField.Function = (Excel.XlConsolidationFunction)xlAverage;
            // NO pivot.RefreshTable();

            // Save and close
            _workbook.SaveAs(_testFile);
            _workbook.Close(false);
            Marshal.ReleaseComObject(_workbook);
            _workbook = null;

            // Reopen
            _workbook = DiagnosticExcelObjects.OpenWorkbook(_excel, _testFile);
            reopenedPivot = DiagnosticExcelObjects.GetPivotTable(_workbook, "Pivot_TestPivot10", "TestPivot10");
            reopenedDataField = DiagnosticExcelObjects.GetDataField(reopenedPivot, 1);
            int reopenedFunction = (int)reopenedDataField.Function;

            _output.WriteLine($"Function Change WITHOUT RefreshTable - Persistence Test:");
            _output.WriteLine($"  Expected Function: {xlAverage} (xlAverage)");
            _output.WriteLine($"  Actual Function:   {reopenedFunction}");
            _output.WriteLine($"  Persisted: {reopenedFunction == xlAverage}");

            // Question: Does function change persist without RefreshTable()?
            // If this fails, RefreshTable() IS required after SetFieldFunction
            Assert.Equal(xlAverage, reopenedFunction);
        }
        finally
        {
            ComUtilities.Release(ref reopenedDataField);
            ComUtilities.Release(ref reopenedPivot);
            ComUtilities.Release(ref valueField);
            ComUtilities.Release(ref rowField);
            ComUtilities.Release(ref pivot);
        }
    }

    [Fact]
    public void Filter_WithoutRefresh_VerifyPersistence()
    {
        Excel.PivotTable? pivot = null;
        Excel.PivotField? rowField = null;
        Excel.PivotItems? items = null;
        Excel.PivotTable? reopenedPivot = null;
        Excel.PivotField? reopenedField = null;
        Excel.PivotItems? reopenedItems = null;
        try
        {
            Assert.NotNull(_excel);
            Assert.NotNull(_workbook);
            // Arrange - Create PivotTable with row field
            pivot = CreatePivotTable("TestPivot11");
            rowField = (Excel.PivotField)pivot.PivotFields("Region");
            rowField.Orientation = (Excel.XlPivotFieldOrientation)xlRowField;
            pivot.RefreshTable(); // Initial setup

            // Act - Apply filter WITHOUT RefreshTable()
            // Set only "North" visible
            items = (Excel.PivotItems)rowField.PivotItems();
            for (int i = 1; i <= items.Count; i++)
            {
                Excel.PivotItem? item = null;
                try
                {
                    item = items.Item(i);
                    string itemName = item.Name?.ToString() ?? "";
                    item.Visible = (itemName == "North");
                }
                finally
                {
                    ComUtilities.Release(ref item);
                }
            }
            // NO pivot.RefreshTable();

            // Save and close
            _workbook.SaveAs(_testFile);
            _workbook.Close(false);
            Marshal.ReleaseComObject(_workbook);
            _workbook = null;

            // Reopen
            _workbook = DiagnosticExcelObjects.OpenWorkbook(_excel, _testFile);
            reopenedPivot = DiagnosticExcelObjects.GetPivotTable(_workbook, "Pivot_TestPivot11", "TestPivot11");
            reopenedField = (Excel.PivotField)reopenedPivot.PivotFields("Region");
            reopenedItems = (Excel.PivotItems)reopenedField.PivotItems();

            int visibleCount = 0;
            string visibleItemName = "";
            for (int i = 1; i <= reopenedItems.Count; i++)
            {
                Excel.PivotItem? item = null;
                try
                {
                    item = reopenedItems.Item(i);
                    if (item.Visible)
                    {
                        visibleCount++;
                        visibleItemName = item.Name?.ToString() ?? "";
                    }
                }
                finally
                {
                    ComUtilities.Release(ref item);
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
        finally
        {
            ComUtilities.Release(ref reopenedItems);
            ComUtilities.Release(ref reopenedField);
            ComUtilities.Release(ref reopenedPivot);
            ComUtilities.Release(ref items);
            ComUtilities.Release(ref rowField);
            ComUtilities.Release(ref pivot);
        }
    }

    public void Dispose()
    {
        GC.SuppressFinalize(this);
        ExcelShutdownService.CloseAndQuit(_workbook, _excel, save: false);
        _workbook = null;
        _excel = null;
        if (File.Exists(_testFile))
            File.Delete(_testFile);
    }
}
