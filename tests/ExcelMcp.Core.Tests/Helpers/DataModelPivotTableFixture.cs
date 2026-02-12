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
/// Unified fixture that creates ONE comprehensive Data Model + PivotTable workbook per test CLASS.
/// Consolidates DataModelTestsFixture and PivotTableRealisticFixture into one fixture.
///
/// Creates:
/// - Data Model tables with relationships (SalesTable → CustomersTable, SalesTable → ProductsTable)
/// - DAX measures for aggregation
/// - PivotTables from multiple source types (range, table, Data Model)
/// - Disambiguation test data for OLAP field matching tests
///
/// The fixture initialization IS the test for creation.
/// - Created ONCE before any tests run
/// - Shared READ-ONLY by all tests in the class
/// - Each test gets its own batch (isolation at batch level)
/// </summary>
public class DataModelPivotTableFixture : IAsyncLifetime
{
    private readonly string _tempDir;

    /// <summary>
    /// Path to the test file
    /// </summary>
    public string TestFilePath { get; private set; } = null!;

    /// <summary>
    /// Results of creation (exposed for validation)
    /// </summary>
    public DataModelPivotTableCreationResult CreationResult { get; private set; } = null!;

    public DataModelPivotTableFixture()
    {
        _tempDir = Path.Join(Path.GetTempPath(), $"DataModelPivotTableTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    /// <summary>
    /// Called ONCE before any tests in the class run.
    /// Creates comprehensive Data Model with relationships, measures, and PivotTables.
    /// </summary>
    public Task InitializeAsync()
    {
        var sw = Stopwatch.StartNew();

        TestFilePath = Path.Join(_tempDir, "DataModelPivotTables.xlsx");
        CreationResult = new DataModelPivotTableCreationResult();

        try
        {
            using (var manager = new SessionManager())
            {
                var sessionId = manager.CreateSessionForNewFile(TestFilePath, show: false);
                manager.CloseSession(sessionId, save: true);
            }
            CreationResult.FileCreated = true;

            using var batch = ExcelSession.BeginBatch(TestFilePath);

            var sheetCommands = new SheetCommands();
            var tableCommands = new TableCommands();
            var dataModelCommands = new DataModelCommands();
            var pivotCommands = new PivotTableCommands();

            // ========================================
            // PART 1: Data Model Tables with Relationships (from DataModelTestsFixture)
            // ========================================

            // Create SalesTable (main fact table)
            CreateSalesTable(batch);

            // Create CustomersTable (dimension)
            CreateCustomersTable(batch);

            // Create ProductsTable (dimension)
            CreateProductsTable(batch);

            CreationResult.TablesCreated = 3;

            // Add tables to Data Model
            tableCommands.AddToDataModel(batch, "SalesTable");
            tableCommands.AddToDataModel(batch, "CustomersTable");
            tableCommands.AddToDataModel(batch, "ProductsTable");
            CreationResult.TablesLoadedToModel = 3;

            // Create relationships
            dataModelCommands.CreateRelationship(
                batch,
                "SalesTable", "CustomerID",
                "CustomersTable", "CustomerID",
                active: true);

            dataModelCommands.CreateRelationship(
                batch,
                "SalesTable", "ProductID",
                "ProductsTable", "ProductID",
                active: true);

            CreationResult.RelationshipsCreated = 2;

            // Create DAX measures on SalesTable
            dataModelCommands.CreateMeasure(
                batch,
                "SalesTable",
                "Total Sales",
                "SUM(SalesTable[Amount])",
                "Total sales amount",
                "#,##0.00");

            dataModelCommands.CreateMeasure(
                batch,
                "SalesTable",
                "Average Sale",
                "AVERAGE(SalesTable[Amount])",
                "Average sale amount",
                "#,##0.00");

            dataModelCommands.CreateMeasure(
                batch,
                "SalesTable",
                "Total Customers",
                "DISTINCTCOUNT(SalesTable[CustomerID])",
                "Count of unique customers",
                "#,##0");

            CreationResult.MeasuresCreated = 3;

            // ========================================
            // PART 2: Regional Sales Table + PivotTables (from PivotTableRealisticFixture)
            // ========================================

            // Create RegionalSalesTable for PivotTable tests
            CreateRegionalSalesTable(batch);

            // Create DisambiguationTable for OLAP field matching tests
            CreateDisambiguationTable(batch);

            CreationResult.TablesCreated += 2;  // Now 5 total

            // Add to Data Model
            tableCommands.AddToDataModel(batch, "RegionalSalesTable");
            tableCommands.AddToDataModel(batch, "DisambiguationTable");
            CreationResult.TablesLoadedToModel += 2;  // Now 5 total

            // Create measures for RegionalSalesTable
            dataModelCommands.CreateMeasure(
                batch,
                "RegionalSalesTable",
                "TotalRevenue",
                "SUM([Sales])",
                "Total Revenue from all regions",
                "#,##0");

            CreationResult.MeasuresCreated++;

            // Create disambiguation measures (names that could be confused with column names)
            dataModelCommands.CreateMeasure(
                batch,
                "DisambiguationTable",
                "ACR",  // Could be confused with "ACRTypeKey" column
                "SUM([Amount])",
                "ACR Amount measure",
                "#,##0.00");

            dataModelCommands.CreateMeasure(
                batch,
                "DisambiguationTable",
                "Discount",  // Could be confused with "DiscountCode" column
                "SUM([Amount]) * 0.1",
                "Discount measure",
                "#,##0.00");

            CreationResult.MeasuresCreated += 2;  // Now 6 total

            // ========================================
            // PART 3: Create PivotTables
            // ========================================

            // 1. Range-based PivotTable (from SalesData range)
            sheetCommands.Create(batch, "PivotData");
            var rangePivot = pivotCommands.CreateFromRange(
                batch,
                "SalesData",
                "A1:F11",  // SalesTable range
                "PivotData",
                "A1",
                "SalesByRegion");

            if (!rangePivot.Success)
                throw new InvalidOperationException(
                    $"CREATION TEST FAILED: Range PivotTable creation failed: {rangePivot.ErrorMessage}");

            CreationResult.RangePivotTablesCreated = 1;

            // 2. Table-based PivotTable (from RegionalSalesTable)
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
            pivotCommands.AddRowField(batch, "RegionalSummary", "Quarter", null);
            pivotCommands.AddColumnField(batch, "RegionalSummary", "Region", null);
            pivotCommands.AddValueField(batch, "RegionalSummary", "Sales", AggregationFunction.Sum, "Total Sales");

            CreationResult.TablePivotTablesCreated = 1;

            // 3. Data Model PivotTable (from RegionalSalesTable in Data Model)
            sheetCommands.Create(batch, "ModelData");
            var dataModelPivot = pivotCommands.CreateFromDataModel(
                batch,
                "RegionalSalesTable",
                "ModelData",
                "A1",
                "DataModelPivot");

            if (!dataModelPivot.Success)
                throw new InvalidOperationException(
                    $"CREATION TEST FAILED: Data Model PivotTable creation failed: {dataModelPivot.ErrorMessage}");

            // Add fields from Data Model (OLAP requires [TableName].[ColumnName] format)
            pivotCommands.AddRowField(batch, "DataModelPivot", "[RegionalSalesTable].[Region]", null);
            pivotCommands.AddValueField(batch, "DataModelPivot", "[Measures].[TotalRevenue]", AggregationFunction.Sum, "Revenue");

            CreationResult.DataModelPivotTablesCreated = 1;

            // 4. Disambiguation PivotTable (for OLAP field matching tests)
            sheetCommands.Create(batch, "DisambiguationPivot");
            var disambigPivot = pivotCommands.CreateFromDataModel(
                batch,
                "DisambiguationTable",
                "DisambiguationPivot",
                "A1",
                "DisambiguationTest");

            if (!disambigPivot.Success)
                throw new InvalidOperationException(
                    $"CREATION TEST FAILED: Disambiguation PivotTable creation failed: {disambigPivot.ErrorMessage}");

            CreationResult.DataModelPivotTablesCreated++;

            // Save the workbook
            batch.Save();

            sw.Stop();
            CreationResult.Success = true;
            CreationResult.CreationTimeMs = sw.ElapsedMilliseconds;

            Console.WriteLine($"✅ DataModelPivotTable fixture created in {sw.ElapsedMilliseconds}ms:");
            Console.WriteLine($"   - {CreationResult.TablesCreated} source tables");
            Console.WriteLine($"   - {CreationResult.TablesLoadedToModel} tables in Data Model");
            Console.WriteLine($"   - {CreationResult.RelationshipsCreated} relationships");
            Console.WriteLine($"   - {CreationResult.MeasuresCreated} DAX measures");
            Console.WriteLine($"   - {CreationResult.RangePivotTablesCreated} range-based PivotTables");
            Console.WriteLine($"   - {CreationResult.TablePivotTablesCreated} table-based PivotTables");
            Console.WriteLine($"   - {CreationResult.DataModelPivotTablesCreated} Data Model PivotTables");
        }
        catch (Exception ex)
        {
            CreationResult.Success = false;
            CreationResult.ErrorMessage = ex.Message;
            sw.Stop();
            Console.WriteLine($"❌ DataModelPivotTable fixture creation FAILED after {sw.ElapsedMilliseconds}ms: {ex.Message}");
            throw;
        }

        return Task.CompletedTask;
    }

    /// <summary>
    /// Creates SalesTable with data for Data Model relationship tests.
    /// Columns: SalesID, Date, CustomerID, ProductID, Amount, Quantity
    /// </summary>
    private static void CreateSalesTable(IExcelBatch batch)
    {
        var sheetCommands = new SheetCommands();
        var rangeCommands = new RangeCommands();
        var tableCommands = new TableCommands();

        sheetCommands.Create(batch, "SalesData");

        var allData = new List<List<object?>>
        {
            new() { "SalesID", "Date", "CustomerID", "ProductID", "Amount", "Quantity" },
            new() { 1, new DateTime(2024, 1, 15), 101, 1001, 150.00, 2 },
            new() { 2, new DateTime(2024, 1, 20), 102, 1002, 250.00, 3 },
            new() { 3, new DateTime(2024, 2, 10), 101, 1003, 175.00, 1 },
            new() { 4, new DateTime(2024, 2, 15), 103, 1001, 300.00, 4 },
            new() { 5, new DateTime(2024, 3, 5), 102, 1002, 125.00, 2 },
            new() { 6, new DateTime(2024, 3, 10), 104, 1003, 450.00, 5 },
            new() { 7, new DateTime(2024, 4, 12), 101, 1001, 200.00, 2 },
            new() { 8, new DateTime(2024, 4, 18), 103, 1002, 350.00, 4 },
            new() { 9, new DateTime(2024, 5, 8), 105, 1003, 275.00, 3 },
            new() { 10, new DateTime(2024, 5, 22), 102, 1001, 180.00, 2 }
        };

        rangeCommands.SetValues(batch, "SalesData", "A1:F11", allData);
        tableCommands.Create(batch, "SalesData", "SalesTable", "A1:F11", true);
    }

    /// <summary>
    /// Creates CustomersTable for relationship tests.
    /// Columns: CustomerID, Name, Region, Country
    /// </summary>
    private static void CreateCustomersTable(IExcelBatch batch)
    {
        var sheetCommands = new SheetCommands();
        var rangeCommands = new RangeCommands();
        var tableCommands = new TableCommands();

        sheetCommands.Create(batch, "Customers");

        var allData = new List<List<object?>>
        {
            new() { "CustomerID", "Name", "Region", "Country" },
            new() { 101, "Acme Corp", "North", "USA" },
            new() { 102, "Beta Inc", "South", "USA" },
            new() { 103, "Gamma LLC", "East", "Canada" },
            new() { 104, "Delta Co", "West", "Canada" },
            new() { 105, "Epsilon Ltd", "North", "UK" }
        };

        rangeCommands.SetValues(batch, "Customers", "A1:D6", allData);
        tableCommands.Create(batch, "Customers", "CustomersTable", "A1:D6", true);
    }

    /// <summary>
    /// Creates ProductsTable for relationship tests.
    /// Columns: ProductID, ProductName, Category, UnitPrice
    /// </summary>
    private static void CreateProductsTable(IExcelBatch batch)
    {
        var sheetCommands = new SheetCommands();
        var rangeCommands = new RangeCommands();
        var tableCommands = new TableCommands();

        sheetCommands.Create(batch, "Products");

        var allData = new List<List<object?>>
        {
            new() { "ProductID", "ProductName", "Category", "UnitPrice" },
            new() { 1001, "Widget A", "Widgets", 75.00 },
            new() { 1002, "Gadget B", "Gadgets", 125.00 },
            new() { 1003, "Device C", "Devices", 175.00 }
        };

        rangeCommands.SetValues(batch, "Products", "A1:D4", allData);
        tableCommands.Create(batch, "Products", "ProductsTable", "A1:D4", true);
    }

    /// <summary>
    /// Creates RegionalSalesTable for PivotTable tests.
    /// Columns: Quarter, Region, Sales, Units
    /// </summary>
    private static void CreateRegionalSalesTable(IExcelBatch batch)
    {
        var sheetCommands = new SheetCommands();
        var rangeCommands = new RangeCommands();
        var tableCommands = new TableCommands();

        sheetCommands.Create(batch, "RegionalData");

        var allData = new List<List<object?>>
        {
            new() { "Quarter", "Region", "Sales", "Units" },
            new() { "Q1", "North", 5000, 100 },
            new() { "Q1", "South", 6000, 120 },
            new() { "Q1", "East", 5500, 110 },
            new() { "Q1", "West", 7000, 140 },
            new() { "Q2", "North", 5500, 110 },
            new() { "Q2", "South", 6500, 130 },
            new() { "Q2", "East", 6000, 120 },
            new() { "Q2", "West", 7500, 150 }
        };

        rangeCommands.SetValues(batch, "RegionalData", "A1:D9", allData);
        tableCommands.Create(batch, "RegionalData", "RegionalSalesTable", "A1:D9", true);
    }

    /// <summary>
    /// Creates DisambiguationTable for OLAP field matching tests.
    /// Has columns that could be confused with measure names:
    /// - "ACRTypeKey" column vs "ACR" measure
    /// - "DiscountCode" column vs "Discount" measure
    /// </summary>
    private static void CreateDisambiguationTable(IExcelBatch batch)
    {
        var sheetCommands = new SheetCommands();
        var rangeCommands = new RangeCommands();
        var tableCommands = new TableCommands();

        sheetCommands.Create(batch, "DisambiguationData");

        var allData = new List<List<object?>>
        {
            new() { "ID", "ACRTypeKey", "DiscountCode", "Amount", "Category" },
            new() { 1, "ACR001", "DISC10", 1000.00, "TypeA" },
            new() { 2, "ACR002", "DISC20", 2500.00, "TypeB" },
            new() { 3, "ACR001", "DISC10", 1500.00, "TypeA" },
            new() { 4, "ACR003", "DISC30", 800.00, "TypeC" },
            new() { 5, "ACR002", "DISC20", 3000.00, "TypeB" }
        };

        rangeCommands.SetValues(batch, "DisambiguationData", "A1:E6", allData);
        tableCommands.Create(batch, "DisambiguationData", "DisambiguationTable", "A1:E6", true);
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
/// Results of comprehensive Data Model + PivotTable workbook creation.
/// </summary>
public class DataModelPivotTableCreationResult
{
    public bool Success { get; set; }
    public bool FileCreated { get; set; }
    public int TablesCreated { get; set; }
    public int TablesLoadedToModel { get; set; }
    public int RelationshipsCreated { get; set; }
    public int MeasuresCreated { get; set; }
    public int RangePivotTablesCreated { get; set; }
    public int TablePivotTablesCreated { get; set; }
    public int DataModelPivotTablesCreated { get; set; }
    public long CreationTimeMs { get; set; }
    public string? ErrorMessage { get; set; }
}




