using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// Fixture that creates ONE realistic PivotTable workbook per test CLASS.
/// Creates PivotTables from multiple source types to test real-world scenarios:
/// - PivotTables from Excel Tables (range-based)
/// - PivotTables from Data Model (Power Pivot)
/// - PivotTables with various field configurations
/// The fixture initialization IS the test for PivotTable creation.
/// - Created ONCE before any tests run
/// - Shared READ-ONLY by all tests in the class
/// - Each test gets its own batch (isolation at batch level)
/// </summary>
public class PivotTableRealisticFixture : IAsyncLifetime
{
    private readonly string _tempDir;

    /// <summary>
    /// Path to the test PivotTable file
    /// </summary>
    public string TestFilePath { get; private set; } = null!;

    /// <summary>
    /// Results of PivotTable creation (exposed for validation)
    /// </summary>
    public PivotTableRealisticCreationResult CreationResult { get; private set; } = null!;

    public PivotTableRealisticFixture()
    {
        _tempDir = Path.Join(Path.GetTempPath(), $"PivotTableRealisticTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    /// <summary>
    /// Called ONCE before any tests in the class run.
    /// Creates realistic PivotTables from multiple source types.
    /// </summary>
    public async Task InitializeAsync()
    {
        var sw = Stopwatch.StartNew();

        TestFilePath = Path.Join(_tempDir, "RealisticPivotTables.xlsx");
        CreationResult = new PivotTableRealisticCreationResult();

        try
        {
            var fileCommands = new FileCommands();
            var createFileResult = fileCommands.CreateEmpty(TestFilePath);
            if (!createFileResult.Success)
                throw new InvalidOperationException(
                    $"CREATION TEST FAILED: File creation failed: {createFileResult.ErrorMessage}");

            CreationResult.FileCreated = true;

            using var batch = ExcelSession.BeginBatch(TestFilePath);

            // Create source data tables
            CreateSalesTable(batch);
            CreateRegionalSalesTable(batch);
            CreationResult.TablesCreated = 2;

            var sheetCommands = new SheetCommands();
            var tableCommands = new TableCommands();
            var pivotCommands = new PivotTableCommands();
            var dataModelCommands = new DataModelCommands();

            // 1. Create PivotTable from Range (simple scenario)
            var createPivotData = sheetCommands.Create(batch, "PivotData");

            var rangePivot = pivotCommands.CreateFromRange(
                batch,
                "SalesData",
                "A1:D11",
                "PivotData",
                "A1",
                "SalesByRegion");

            if (!rangePivot.Success)
                throw new InvalidOperationException(
                    $"CREATION TEST FAILED: Range PivotTable creation failed: {rangePivot.ErrorMessage}");

            // Add fields to range pivot
            var addRowField1 = pivotCommands.AddRowField(batch, "SalesByRegion", "Region", null);
            var addValueField1 = pivotCommands.AddValueField(batch, "SalesByRegion", "Revenue", AggregationFunction.Sum, "Total Revenue");

            CreationResult.RangePivotTablesCreated = 1;

            // 2. Create PivotTable from Excel Table
            var tablePivot = pivotCommands.CreateFromTable(
                batch,
                "RegionalSalesTable",
                "RegionalData",
                "F1",
                "RegionalSummary");

            if (!tablePivot.Success)
                throw new InvalidOperationException(
                    $"CREATION TEST FAILED: Table PivotTable creation failed: {tablePivot.ErrorMessage}");

            // Add fields to table pivot
            var addRowField2 = pivotCommands.AddRowField(batch, "RegionalSummary", "Quarter", null);
            var addColField1 = pivotCommands.AddColumnField(batch, "RegionalSummary", "Region", null);
            var addValueField2 = pivotCommands.AddValueField(batch, "RegionalSummary", "Sales", AggregationFunction.Sum, "Total Sales");

            CreationResult.TablePivotTablesCreated = 1;

            // 3. Create Data Model PivotTable
            // First add table to Data Model
            var addToModel = tableCommands.AddToDataModel(batch, "RegionalSalesTable");
            if (!addToModel.Success)
                throw new InvalidOperationException(
                    $"CREATION TEST FAILED: AddToDataModel failed: {addToModel.ErrorMessage}");

            // Create measure in Data Model
            var measure = dataModelCommands.CreateMeasure(
                batch,
                "RegionalSalesTable",
                "TotalRevenue",
                "SUM([Sales])",
                "Total Revenue from all regions",
                "#,##0");

            if (!measure.Success)
                throw new InvalidOperationException(
                    $"CREATION TEST FAILED: Measure creation failed: {measure.ErrorMessage}");

            // Create PivotTable from Data Model
            var createModelData = sheetCommands.Create(batch, "ModelData");
            var dataModelPivot = pivotCommands.CreateFromDataModel(
                batch,
                "RegionalSalesTable",
                "ModelData",
                "A1",
                "DataModelPivot");

            if (!dataModelPivot.Success)
                throw new InvalidOperationException(
                    $"CREATION TEST FAILED: Data Model PivotTable creation failed: {dataModelPivot.ErrorMessage}");

            // Add fields from Data Model
            var addRowField3 = pivotCommands.AddRowField(batch, "DataModelPivot", "Region", null);
            // Data Model measures don't use aggregation function - they have their own DAX formula
            var addValueField3 = pivotCommands.AddValueField(batch, "DataModelPivot", "[Measures].[TotalRevenue]", AggregationFunction.Sum, "Revenue");

            CreationResult.DataModelPivotTablesCreated = 1;
            CreationResult.MeasuresCreated = 1;

            // Save the workbook
            batch.Save();

            sw.Stop();
            CreationResult.Success = true;
            CreationResult.CreationTimeMs = sw.ElapsedMilliseconds;

            Console.WriteLine($"✅ PivotTable fixture created in {sw.ElapsedMilliseconds}ms:");
            Console.WriteLine($"   - {CreationResult.TablesCreated} source tables");
            Console.WriteLine($"   - {CreationResult.RangePivotTablesCreated} range-based PivotTables");
            Console.WriteLine($"   - {CreationResult.TablePivotTablesCreated} table-based PivotTables");
            Console.WriteLine($"   - {CreationResult.DataModelPivotTablesCreated} Data Model PivotTables");
        }
        catch (Exception ex)
        {
            CreationResult.Success = false;
            CreationResult.ErrorMessage = ex.Message;
            sw.Stop();
            Console.WriteLine($"❌ PivotTable fixture creation FAILED after {sw.ElapsedMilliseconds}ms: {ex.Message}");
            throw;
        }
    }

    private static void CreateSalesTable(IExcelBatch batch)
    {
        var sheetCommands = new SheetCommands();
        var rangeCommands = new RangeCommands();

        sheetCommands.Create(batch, "SalesData");

        // Write headers AND data together (like the working Range tests do)
        var allData = new List<List<object?>>
        {
            // Headers
            new() { "Date", "Region", "Product", "Revenue" },
            // Data rows
            new() { new DateTime(2024, 1, 15), "North", "Widget", 1000 },
            new() { new DateTime(2024, 1, 20), "South", "Gadget", 1500 },
            new() { new DateTime(2024, 2, 10), "East", "Widget", 1200 },
            new() { new DateTime(2024, 2, 15), "West", "Gadget", 1800 },
            new() { new DateTime(2024, 3, 5), "North", "Widget", 1100 },
            new() { new DateTime(2024, 3, 10), "South", "Gadget", 1600 },
            new() { new DateTime(2024, 4, 12), "East", "Widget", 1300 },
            new() { new DateTime(2024, 4, 18), "West", "Gadget", 1900 },
            new() { new DateTime(2024, 5, 8), "North", "Widget", 1050 },
            new() { new DateTime(2024, 5, 22), "South", "Gadget", 1550 }
        };

        // Specify full range like the working tests: "A1:D11" (1 header row + 10 data rows)
        rangeCommands.SetValues(batch, "SalesData", "A1:D11", allData);
    }

    private static void CreateRegionalSalesTable(IExcelBatch batch)
    {
        var sheetCommands = new SheetCommands();
        var rangeCommands = new RangeCommands();
        var tableCommands = new TableCommands();

        sheetCommands.Create(batch, "RegionalData");

        // Write headers AND data together with full range specification
        var allData = new List<List<object?>>
        {
            // Headers
            new() { "Quarter", "Region", "Sales", "Units" },
            // Data rows
            new() { "Q1", "North", 5000, 100 },
            new() { "Q1", "South", 6000, 120 },
            new() { "Q1", "East", 5500, 110 },
            new() { "Q1", "West", 7000, 140 },
            new() { "Q2", "North", 5500, 110 },
            new() { "Q2", "South", 6500, 130 },
            new() { "Q2", "East", 6000, 120 },
            new() { "Q2", "West", 7500, 150 }
        };

        // Specify full range: "A1:D9" (1 header row + 8 data rows)
        rangeCommands.SetValues(batch, "RegionalData", "A1:D9", allData);

        // Create Excel Table
        tableCommands.Create(batch, "RegionalData", "RegionalSalesTable", "A1:D9", true);
    }

    public Task DisposeAsync()
    {
        try
        {
            if (Directory.Exists(_tempDir))
                Directory.Delete(_tempDir, true);
        }
        catch
        {
            // Ignore cleanup errors
        }

        return Task.CompletedTask;
    }
}

/// <summary>
/// Results of realistic PivotTable workbook creation
/// </summary>
public class PivotTableRealisticCreationResult
{
    public bool Success { get; set; }
    public bool FileCreated { get; set; }
    public int TablesCreated { get; set; }
    public int RangePivotTablesCreated { get; set; }
    public int TablePivotTablesCreated { get; set; }
    public int DataModelPivotTablesCreated { get; set; }
    public int MeasuresCreated { get; set; }
    public long CreationTimeMs { get; set; }
    public string? ErrorMessage { get; set; }
}
