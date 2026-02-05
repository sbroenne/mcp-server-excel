// =============================================================================
// DIAGNOSTIC TESTS - Direct Excel COM API Behavior for Data Model
// =============================================================================
// Purpose: Understand what Excel COM API actually does for Data Model operations
// These tests document the REAL behavior of Excel's Data Model/Power Pivot COM API
// =============================================================================

// Suppress invalid-dynamic-call warnings - this diagnostic test file intentionally uses
// dynamic COM interop patterns to explore Excel's behavior. The Range[cell] pattern is
// standard COM interop for Excel and cannot be statically analyzed.
#pragma warning disable CS1061 // Member access on dynamic type - expected for COM interop exploration

using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Diagnostics.Tests.Integration.Diagnostics;

/// <summary>
/// Diagnostic tests for Data Model (Power Pivot) COM API behavior.
/// These tests use raw COM calls to understand Excel's actual behavior.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "Diagnostics")]
[Trait("Feature", "DataModel")]
[Trait("RequiresExcel", "true")]
[Trait("RunType", "OnDemand")]
public class DataModelComApiBehaviorTests : IClassFixture<TempDirectoryFixture>, IDisposable
{
    private readonly string _tempDir;
    private readonly ITestOutputHelper _output;
    private dynamic? _excel;
    private dynamic? _workbook;
    private readonly string _testFile;

    // Simple M code for creating Data Model tables
    private const string SalesQuery = """
        let
            Source = #table(
                {"Product", "Amount", "Quantity"},
                {{"Widget", 100, 5}, {"Gadget", 200, 3}, {"Gizmo", 150, 7}}
            )
        in
            Source
        """;

    private const string ProductsQuery = """
        let
            Source = #table(
                {"ProductName", "Category", "Price"},
                {{"Widget", "Electronics", 20}, {"Gadget", "Electronics", 66.67}, {"Gizmo", "Tools", 21.43}}
            )
        in
            Source
        """;

    public DataModelComApiBehaviorTests(TempDirectoryFixture fixture, ITestOutputHelper output)
    {
        _tempDir = fixture.TempDir;
        _output = output;
        _testFile = Path.Combine(_tempDir, $"DMDiag_{Guid.NewGuid():N}.xlsx");

        // Create Excel instance directly via COM
        var excelType = Type.GetTypeFromProgID("Excel.Application");
        _excel = Activator.CreateInstance(excelType!);
        _excel.Visible = false;
        _excel.DisplayAlerts = false;

        // Create new workbook
        _workbook = _excel.Workbooks.Add();
        _workbook.SaveAs(_testFile);

        _output.WriteLine($"Test file: {_testFile}");
    }

    public void Dispose()
    {
        try
        {
            if (_workbook != null)
            {
                _workbook.Close(false);
                ComUtilities.Release(ref _workbook);
            }
            if (_excel != null)
            {
                _excel.Quit();
                ComUtilities.Release(ref _excel);
            }
        }
        catch (Exception ex)
        {
            _output.WriteLine($"Cleanup error: {ex.Message}");
        }
        GC.SuppressFinalize(this);
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    // =========================================================================
    // SCENARIO 1: Access Data Model
    // =========================================================================

    [Fact]
    public void Scenario1_AccessDataModel()
    {
        _output.WriteLine("=== SCENARIO 1: Access Data Model ===");

        dynamic? model = null;
        dynamic? modelTables = null;
        dynamic? modelRelationships = null;

        try
        {
            model = _workbook.Model;
            _output.WriteLine($"Model object obtained: {model != null}");

            modelTables = model.ModelTables;
            _output.WriteLine($"Initial ModelTables count: {modelTables.Count}");

            modelRelationships = model.ModelRelationships;
            _output.WriteLine($"Initial ModelRelationships count: {modelRelationships.Count}");

            // Check model name/properties
            try
            {
                string modelName = model.Name;
                _output.WriteLine($"Model name: {modelName}");
            }
            catch
            {
                _output.WriteLine("Model.Name not accessible");
            }
        }
        finally
        {
            ComUtilities.Release(ref modelRelationships);
            ComUtilities.Release(ref modelTables);
            ComUtilities.Release(ref model);
        }

        _output.WriteLine("=== SCENARIO 1 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 2: Create Table in Data Model (via Power Query)
    // =========================================================================

    [Fact]
    public void Scenario2_CreateTableInDataModel()
    {
        _output.WriteLine("=== SCENARIO 2: Create Table in Data Model ===");

        dynamic? queries = null;
        dynamic? query = null;
        dynamic? connections = null;
        dynamic? model = null;
        dynamic? modelTables = null;

        try
        {
            // First create a Power Query
            queries = _workbook.Queries;
            query = queries.Add("Sales", SalesQuery);
            _output.WriteLine("Power Query 'Sales' created");

            // Now load to Data Model using Add2 with CreateModelConnection=true
            connections = _workbook.Connections;

            string connString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Sales";

            _output.WriteLine("\n--- Adding connection with CreateModelConnection=true ---");
            dynamic? conn = null;
            try
            {
                conn = connections.Add2(
                    "Query - Sales",                   // Name
                    "Power Query - Sales",             // Description
                    connString,                        // ConnectionString
                    "SELECT * FROM [Sales]",           // CommandText
                    2,                                 // lCmdtype (xlCmdSql = 2)
                    true,                              // CreateModelConnection - LOAD TO DATA MODEL
                    false                              // ImportRelationships
                );
                _output.WriteLine("Connection added with CreateModelConnection=true");

                // Refresh to load data into model
                conn.Refresh();
                _output.WriteLine("Connection refreshed");

                ComUtilities.Release(ref conn);
            }
            catch (COMException ex)
            {
                _output.WriteLine($"Add2 failed: 0x{ex.HResult:X8} - {ex.Message}");
                ComUtilities.Release(ref conn);
            }

            // Check if table appeared in Data Model
            model = _workbook.Model;
            modelTables = model.ModelTables;
            _output.WriteLine($"\nModelTables count after load: {modelTables.Count}");

            for (int i = 1; i <= modelTables.Count; i++)
            {
                dynamic? table = modelTables.Item(i);
                _output.WriteLine($"  Table {i}: {table.Name}");

                // List columns
                dynamic? columns = table.ModelTableColumns;
                _output.WriteLine($"    Columns: {columns.Count}");
                for (int j = 1; j <= columns.Count; j++)
                {
                    dynamic? col = columns.Item(j);
                    _output.WriteLine($"      - {col.Name} ({col.DataType})");
                    ComUtilities.Release(ref col);
                }
                ComUtilities.Release(ref columns);
                ComUtilities.Release(ref table);
            }
        }
        finally
        {
            ComUtilities.Release(ref modelTables);
            ComUtilities.Release(ref model);
            ComUtilities.Release(ref connections);
            ComUtilities.Release(ref query);
            ComUtilities.Release(ref queries);
        }

        _output.WriteLine("=== SCENARIO 2 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 3: Add DAX Measure
    // =========================================================================

    [Fact]
    public void Scenario3_AddDaxMeasure()
    {
        _output.WriteLine("=== SCENARIO 3: Add DAX Measure ===");

        dynamic? model = null;
        dynamic? modelTables = null;
        dynamic? measures = null;

        try
        {
            // First ensure we have a table in the model
            CreateDataModelTable("Sales", SalesQuery);

            model = _workbook.Model;
            modelTables = model.ModelTables;

            if (modelTables.Count == 0)
            {
                _output.WriteLine("No tables in Data Model. Skipping measure test.");
                return;
            }

            dynamic? table = modelTables.Item(1);
            string tableName = table.Name;
            _output.WriteLine($"Adding measure to table: {tableName}");

            // Get measures collection
            measures = model.ModelMeasures;
            int measureCountBefore = measures.Count;
            _output.WriteLine($"Measures before add: {measureCountBefore}");

            // Add a measure
            _output.WriteLine("\n--- Adding DAX measure ---");
            dynamic? formatInfo = null;
            try
            {
                // ModelMeasures.Add signature: (MeasureName, AssociatedTable, Formula, FormatInformation, [Description])
                // FormatInformation is REQUIRED - get from Model.ModelFormatGeneral property
                formatInfo = model.ModelFormatGeneral;
                dynamic? measure = measures.Add(
                    "TotalAmount",                           // MeasureName
                    table,                                   // AssociatedTable
                    "SUM(Query[Amount])",                    // Formula (use table name "Query" from M code)
                    formatInfo,                              // FormatInformation (REQUIRED)
                    "Total of all amounts"                   // Description (optional)
                );

                _output.WriteLine($"Measure added: {measure.Name}");
                _output.WriteLine($"Formula: {measure.Formula}");

                ComUtilities.Release(ref measure);
            }
            catch (COMException ex)
            {
                _output.WriteLine($"Add measure failed: 0x{ex.HResult:X8}");
                _output.WriteLine($"Message: {ex.Message}");
            }
            finally
            {
                ComUtilities.Release(ref formatInfo);
            }

            // Verify
            ComUtilities.Release(ref measures);
            measures = model.ModelMeasures;
            _output.WriteLine($"\nMeasures after add: {measures.Count}");

            ComUtilities.Release(ref table);
        }
        finally
        {
            ComUtilities.Release(ref measures);
            ComUtilities.Release(ref modelTables);
            ComUtilities.Release(ref model);
        }

        _output.WriteLine("=== SCENARIO 3 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 4: Update DAX Measure
    // =========================================================================

    [Fact]
    public void Scenario4_UpdateDaxMeasure()
    {
        _output.WriteLine("=== SCENARIO 4: Update DAX Measure ===");

        dynamic? model = null;
        dynamic? measures = null;

        try
        {
            // Setup: Create table and measure
            CreateDataModelTable("Sales", SalesQuery);
            CreateMeasure("TotalAmount", "SUM(Sales[Amount])");

            model = _workbook.Model;
            measures = model.ModelMeasures;

            if (measures.Count == 0)
            {
                _output.WriteLine("No measures found. Skipping update test.");
                return;
            }

            // Find and update the measure
            dynamic? measure = null;
            for (int i = 1; i <= measures.Count; i++)
            {
                dynamic? m = measures.Item(i);
                if (m.Name == "TotalAmount")
                {
                    measure = m;
                    break;
                }
                ComUtilities.Release(ref m);
            }

            if (measure == null)
            {
                _output.WriteLine("Measure 'TotalAmount' not found");
                return;
            }

            _output.WriteLine($"Original formula: {measure.Formula}");

            // Update formula
            _output.WriteLine("\n--- Updating measure formula ---");
            try
            {
                measure.Formula = "SUM(Sales[Amount]) * 1.1";
                _output.WriteLine($"New formula: {measure.Formula}");
            }
            catch (COMException ex)
            {
                _output.WriteLine($"Update failed: 0x{ex.HResult:X8} - {ex.Message}");
            }

            // Update description
            try
            {
                measure.Description = "Updated: Total with 10% markup";
                _output.WriteLine($"New description: {measure.Description}");
            }
            catch (COMException ex)
            {
                _output.WriteLine($"Description update failed: {ex.Message}");
            }

            ComUtilities.Release(ref measure);
        }
        finally
        {
            ComUtilities.Release(ref measures);
            ComUtilities.Release(ref model);
        }

        _output.WriteLine("=== SCENARIO 4 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 5: Delete DAX Measure
    // =========================================================================

    [Fact]
    public void Scenario5_DeleteDaxMeasure()
    {
        _output.WriteLine("=== SCENARIO 5: Delete DAX Measure ===");

        dynamic? model = null;
        dynamic? measures = null;

        try
        {
            // Setup
            CreateDataModelTable("Sales", SalesQuery);
            CreateMeasure("ToDelete", "SUM(Sales[Amount])");

            model = _workbook.Model;
            measures = model.ModelMeasures;

            int countBefore = measures.Count;
            _output.WriteLine($"Measures before delete: {countBefore}");

            // Find and delete the measure
            for (int i = 1; i <= measures.Count; i++)
            {
                dynamic? m = measures.Item(i);
                if (m.Name == "ToDelete")
                {
                    _output.WriteLine("\n--- Deleting measure 'ToDelete' ---");
                    m.Delete();
                    _output.WriteLine("Measure deleted");
                    ComUtilities.Release(ref m);
                    break;
                }
                ComUtilities.Release(ref m);
            }

            // Verify
            ComUtilities.Release(ref measures);
            measures = model.ModelMeasures;
            _output.WriteLine($"Measures after delete: {measures.Count}");

            Assert.Equal(countBefore - 1, (int)measures.Count);
        }
        finally
        {
            ComUtilities.Release(ref measures);
            ComUtilities.Release(ref model);
        }

        _output.WriteLine("=== SCENARIO 5 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 6: Create Relationship Between Tables
    // =========================================================================

    [Fact]
    public void Scenario6_CreateRelationship()
    {
        _output.WriteLine("=== SCENARIO 6: Create Relationship ===");

        dynamic? model = null;
        dynamic? modelTables = null;
        dynamic? relationships = null;

        try
        {
            // Create two tables with related columns
            CreateDataModelTable("Sales", SalesQuery);
            CreateDataModelTable("Products", ProductsQuery);

            model = _workbook.Model;
            modelTables = model.ModelTables;
            relationships = model.ModelRelationships;

            _output.WriteLine($"Tables in model: {modelTables.Count}");
            _output.WriteLine($"Relationships before: {relationships.Count}");

            if (modelTables.Count < 2)
            {
                _output.WriteLine("Need at least 2 tables for relationship test");
                return;
            }

            // Find the tables and columns
            dynamic? salesTable = null;
            dynamic? productsTable = null;
            dynamic? salesProductCol = null;
            dynamic? productsNameCol = null;

            for (int i = 1; i <= modelTables.Count; i++)
            {
                dynamic? t = modelTables.Item(i);
                string name = t.Name;
                if (name == "Sales")
                    salesTable = t;
                else if (name == "Products")
                    productsTable = t;
                else
                    ComUtilities.Release(ref t);
            }

            if (salesTable != null && productsTable != null)
            {
                // Get the columns
                dynamic? salesCols = salesTable.ModelTableColumns;
                for (int i = 1; i <= salesCols.Count; i++)
                {
                    dynamic? col = salesCols.Item(i);
                    if (col.Name == "Product")
                    {
                        salesProductCol = col;
                        break;
                    }
                    ComUtilities.Release(ref col);
                }
                ComUtilities.Release(ref salesCols);

                dynamic? prodCols = productsTable.ModelTableColumns;
                for (int i = 1; i <= prodCols.Count; i++)
                {
                    dynamic? col = prodCols.Item(i);
                    if (col.Name == "ProductName")
                    {
                        productsNameCol = col;
                        break;
                    }
                    ComUtilities.Release(ref col);
                }
                ComUtilities.Release(ref prodCols);

                if (salesProductCol != null && productsNameCol != null)
                {
                    _output.WriteLine("\n--- Creating relationship ---");
                    try
                    {
                        dynamic? rel = relationships.Add(
                            salesProductCol,     // ForeignKeyColumn (many side)
                            productsNameCol      // PrimaryKeyColumn (one side)
                        );

                        _output.WriteLine($"Relationship created");
                        _output.WriteLine($"  From: {rel.ForeignKeyColumn.Name} in {rel.ForeignKeyTable.Name}");
                        _output.WriteLine($"  To: {rel.PrimaryKeyColumn.Name} in {rel.PrimaryKeyTable.Name}");
                        _output.WriteLine($"  Active: {rel.Active}");

                        ComUtilities.Release(ref rel);
                    }
                    catch (COMException ex)
                    {
                        _output.WriteLine($"Create relationship failed: 0x{ex.HResult:X8}");
                        _output.WriteLine($"Message: {ex.Message}");
                    }
                }
                else
                {
                    _output.WriteLine("Could not find matching columns for relationship");
                }

                ComUtilities.Release(ref salesProductCol);
                ComUtilities.Release(ref productsNameCol);
            }

            // Verify
            ComUtilities.Release(ref relationships);
            relationships = model.ModelRelationships;
            _output.WriteLine($"\nRelationships after: {relationships.Count}");

            ComUtilities.Release(ref salesTable);
            ComUtilities.Release(ref productsTable);
        }
        finally
        {
            ComUtilities.Release(ref relationships);
            ComUtilities.Release(ref modelTables);
            ComUtilities.Release(ref model);
        }

        _output.WriteLine("=== SCENARIO 6 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 7: Delete Relationship
    // =========================================================================

    [Fact]
    public void Scenario7_DeleteRelationship()
    {
        _output.WriteLine("=== SCENARIO 7: Delete Relationship ===");

        dynamic? model = null;
        dynamic? relationships = null;

        try
        {
            // Setup: Create tables and relationship
            CreateDataModelTable("Sales", SalesQuery);
            CreateDataModelTable("Products", ProductsQuery);
            CreateRelationshipBetweenTables();

            model = _workbook.Model;
            relationships = model.ModelRelationships;

            int countBefore = relationships.Count;
            _output.WriteLine($"Relationships before delete: {countBefore}");

            if (countBefore == 0)
            {
                _output.WriteLine("No relationships to delete");
                return;
            }

            // Delete first relationship
            dynamic? rel = relationships.Item(1);
            _output.WriteLine($"\n--- Deleting relationship ---");
            rel.Delete();
            _output.WriteLine("Relationship deleted");
            ComUtilities.Release(ref rel);

            // Verify
            ComUtilities.Release(ref relationships);
            relationships = model.ModelRelationships;
            _output.WriteLine($"Relationships after delete: {relationships.Count}");
        }
        finally
        {
            ComUtilities.Release(ref relationships);
            ComUtilities.Release(ref model);
        }

        _output.WriteLine("=== SCENARIO 7 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 8: Delete Query That Loaded to Data Model
    // =========================================================================

    [Fact]
    public void Scenario8_DeleteQueryWithDataModelTable()
    {
        _output.WriteLine("=== SCENARIO 8: Delete Query That Loaded to Data Model ===");

        dynamic? queries = null;
        dynamic? model = null;
        dynamic? modelTables = null;

        try
        {
            // Create query and load to data model
            CreateDataModelTable("OrphanTest", SalesQuery);

            queries = _workbook.Queries;
            model = _workbook.Model;
            modelTables = model.ModelTables;

            int queryCountBefore = queries.Count;
            int tableCountBefore = modelTables.Count;

            _output.WriteLine($"Queries before delete: {queryCountBefore}");
            _output.WriteLine($"Model tables before delete: {tableCountBefore}");

            // Find and delete the query
            dynamic? query = null;
            for (int i = 1; i <= queries.Count; i++)
            {
                dynamic? q = queries.Item(i);
                if (q.Name == "OrphanTest")
                {
                    query = q;
                    break;
                }
                ComUtilities.Release(ref q);
            }

            if (query != null)
            {
                _output.WriteLine("\n--- Deleting query 'OrphanTest' ---");
                query.Delete();
                _output.WriteLine("Query deleted");
                ComUtilities.Release(ref query);
            }

            // KEY QUESTION: What happens to the Data Model table?
            ComUtilities.Release(ref modelTables);
            modelTables = model.ModelTables;

            _output.WriteLine($"\nQueries after delete: {queries.Count}");
            _output.WriteLine($"Model tables after delete: {modelTables.Count}");

            if (modelTables.Count == tableCountBefore)
            {
                _output.WriteLine("DATA MODEL TABLE SURVIVES! Query deletion does NOT remove model table.");

                // List remaining tables
                for (int i = 1; i <= modelTables.Count; i++)
                {
                    dynamic? t = modelTables.Item(i);
                    _output.WriteLine($"  Orphaned table: {t.Name}");
                    ComUtilities.Release(ref t);
                }
            }
            else
            {
                _output.WriteLine("DATA MODEL TABLE REMOVED! Query deletion removes model table too.");
            }
        }
        finally
        {
            ComUtilities.Release(ref modelTables);
            ComUtilities.Release(ref model);
            ComUtilities.Release(ref queries);
        }

        _output.WriteLine("=== SCENARIO 8 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 9: Measure Referencing Deleted Table
    // =========================================================================

    [Fact]
    public void Scenario9_MeasureWithDeletedTable()
    {
        _output.WriteLine("=== SCENARIO 9: Measure Referencing Deleted Table ===");

        dynamic? queries = null;
        dynamic? model = null;
        dynamic? measures = null;

        try
        {
            // Create table and measure
            CreateDataModelTable("ToDelete", SalesQuery);
            CreateMeasure("OrphanedMeasure", "SUM(ToDelete[Amount])");

            model = _workbook.Model;
            measures = model.ModelMeasures;
            queries = _workbook.Queries;

            _output.WriteLine($"Measures before table delete: {measures.Count}");

            // Delete the query (which may or may not delete the model table)
            _output.WriteLine("\n--- Deleting source query ---");
            for (int i = 1; i <= queries.Count; i++)
            {
                dynamic? q = queries.Item(i);
                if (q.Name == "ToDelete")
                {
                    q.Delete();
                    _output.WriteLine("Query deleted");
                    ComUtilities.Release(ref q);
                    break;
                }
                ComUtilities.Release(ref q);
            }

            // Check measures
            ComUtilities.Release(ref measures);
            measures = model.ModelMeasures;
            _output.WriteLine($"Measures after table delete: {measures.Count}");

            // Try to access the measure
            for (int i = 1; i <= measures.Count; i++)
            {
                dynamic? m = measures.Item(i);
                try
                {
                    _output.WriteLine($"Measure: {m.Name}");
                    _output.WriteLine($"  Formula: {m.Formula}");
                    _output.WriteLine($"  Table: {m.AssociatedTable?.Name ?? "(null)"}");
                }
                catch (COMException ex)
                {
                    _output.WriteLine($"  ERROR accessing measure: {ex.Message}");
                }
                ComUtilities.Release(ref m);
            }
        }
        finally
        {
            ComUtilities.Release(ref measures);
            ComUtilities.Release(ref model);
            ComUtilities.Release(ref queries);
        }

        _output.WriteLine("=== SCENARIO 9 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 10: Multiple Measures on Same Table
    // =========================================================================

    [Fact]
    public void Scenario10_MultipleMeasures()
    {
        _output.WriteLine("=== SCENARIO 10: Multiple Measures on Same Table ===");

        dynamic? model = null;
        dynamic? measures = null;

        try
        {
            CreateDataModelTable("Sales", SalesQuery);

            CreateMeasure("TotalAmount", "SUM(Sales[Amount])");
            CreateMeasure("TotalQty", "SUM(Sales[Quantity])");
            CreateMeasure("AvgAmount", "AVERAGE(Sales[Amount])");
            CreateMeasure("CountRows", "COUNTROWS(Sales)");

            model = _workbook.Model;
            measures = model.ModelMeasures;

            _output.WriteLine($"Total measures created: {measures.Count}");

            for (int i = 1; i <= measures.Count; i++)
            {
                dynamic? m = measures.Item(i);
                _output.WriteLine($"  {m.Name}: {m.Formula}");
                ComUtilities.Release(ref m);
            }
        }
        finally
        {
            ComUtilities.Release(ref measures);
            ComUtilities.Release(ref model);
        }

        _output.WriteLine("=== SCENARIO 10 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 11: Model Refresh
    // =========================================================================

    [Fact]
    public void Scenario11_ModelRefresh()
    {
        _output.WriteLine("=== SCENARIO 11: Model Refresh ===");

        dynamic? model = null;

        try
        {
            CreateDataModelTable("Sales", SalesQuery);

            model = _workbook.Model;

            _output.WriteLine("--- Attempting model.Refresh() ---");
            try
            {
                model.Refresh();
                _output.WriteLine("model.Refresh() succeeded");
            }
            catch (COMException ex)
            {
                _output.WriteLine($"model.Refresh() failed: 0x{ex.HResult:X8}");
                _output.WriteLine($"Message: {ex.Message}");
            }

            // Alternative: Refresh via connection
            _output.WriteLine("\n--- Attempting connection refresh ---");
            dynamic? connections = _workbook.Connections;
            for (int i = 1; i <= connections.Count; i++)
            {
                dynamic? conn = connections.Item(i);
                try
                {
                    conn.Refresh();
                    _output.WriteLine($"Connection '{conn.Name}' refreshed");
                }
                catch (COMException ex)
                {
                    _output.WriteLine($"Connection '{conn.Name}' refresh failed: {ex.Message}");
                }
                ComUtilities.Release(ref conn);
            }
            ComUtilities.Release(ref connections);
        }
        finally
        {
            ComUtilities.Release(ref model);
        }

        _output.WriteLine("=== SCENARIO 11 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 12: CUBEVALUE Recalculation After Data Model Refresh
    // Issue #313: CUBEVALUE formulas return error codes after refresh
    // =========================================================================

    [Fact]
    public void Scenario12_CubeValueRecalculationAfterRefresh()
    {
        _output.WriteLine("=== SCENARIO 12: CUBEVALUE Recalculation After Refresh ===");
        _output.WriteLine("This test verifies if CUBEVALUE formulas automatically recalculate after Data Model refresh.");

        dynamic? model = null;
        dynamic? sheet = null;
        dynamic? range = null;

        try
        {
            // Step 1: Create Data Model with a table
            _output.WriteLine("\n--- Step 1: Create Data Model table ---");
            CreateDataModelTable("Sales", SalesQuery);

            // Step 2: Create a DAX measure
            _output.WriteLine("\n--- Step 2: Create DAX measure ---");
            CreateMeasure("TotalAmount", "SUM(Sales[Amount])");

            // Step 3: Check current calculation mode
            _output.WriteLine("\n--- Step 3: Check calculation mode ---");
            int calcMode = _excel.Calculation;
            _output.WriteLine($"Current calculation mode: {calcMode} (xlCalculationAutomatic=-4105, xlCalculationManual=-4135)");

            // Step 4: Add CUBEVALUE formula to worksheet
            _output.WriteLine("\n--- Step 4: Add CUBEVALUE formula ---");
            sheet = _workbook.Worksheets.Item(1);

            // First, let's check what cube connections exist
            _output.WriteLine("Checking available cube connections...");
            try
            {
                dynamic? connections = _workbook.Connections;
                for (int i = 1; i <= connections.Count; i++)
                {
                    dynamic? conn = connections.Item(i);
                    _output.WriteLine($"  Connection {i}: '{conn.Name}' Type={conn.Type}");
                    ComUtilities.Release(ref conn);
                }
                ComUtilities.Release(ref connections);
            }
            catch (COMException ex)
            {
                _output.WriteLine($"  Error listing connections: {ex.Message}");
            }

            // Also check the Model's connection
            try
            {
                _output.WriteLine($"Model.Name: {model?.Name ?? "(model is null)"}");
            }
#pragma warning disable CA1031 // Intentional: diagnostic test logs all exceptions
            catch (Exception ex)
#pragma warning restore CA1031
            {
                _output.WriteLine($"  Error getting model name: {ex.Message}");
            }

            // Try different CUBEVALUE connection names
            // CUBEVALUE format: =CUBEVALUE(connection, member_expression)
            // The connection should be the Data Model connection name
            string cubeFormula = "=CUBEVALUE(\"ThisWorkbookDataModel\",\"[Measures].[TotalAmount]\")";
            range = sheet.Range["A1"];
            range.Formula = cubeFormula;
            _output.WriteLine($"Set formula in A1: {cubeFormula}");

            // Also try the alternative connection format in A2
            dynamic? rangeA2 = sheet.Range["A2"];
            string cubeFormula2 = "=CUBEVALUE(\"Query - Sales\",\"[Measures].[TotalAmount]\")";
            rangeA2.Formula = cubeFormula2;
            _output.WriteLine($"Set formula in A2: {cubeFormula2}");
            ComUtilities.Release(ref rangeA2);

            // Try different measure formats in A3-A5
            dynamic? rangeA3 = sheet.Range["A3"];
            rangeA3.Formula = "=CUBEVALUE(\"ThisWorkbookDataModel\",\"TotalAmount\")";
            _output.WriteLine($"Set formula in A3: =CUBEVALUE(\"ThisWorkbookDataModel\",\"TotalAmount\")");
            ComUtilities.Release(ref rangeA3);

            dynamic? rangeA4 = sheet.Range["A4"];
            rangeA4.Formula = "=CUBEVALUE(\"ThisWorkbookDataModel\",\"[Sales].[TotalAmount]\")";
            _output.WriteLine($"Set formula in A4: =CUBEVALUE(\"ThisWorkbookDataModel\",\"[Sales].[TotalAmount]\")");
            ComUtilities.Release(ref rangeA4);

            // Try a CUBEMEMBER first approach
            dynamic? rangeA5 = sheet.Range["A5"];
            rangeA5.Formula = "=CUBEMEMBER(\"ThisWorkbookDataModel\",\"[Measures].[TotalAmount]\")";
            _output.WriteLine($"Set formula in A5: =CUBEMEMBER(\"ThisWorkbookDataModel\",\"[Measures].[TotalAmount]\")");
            ComUtilities.Release(ref rangeA5);

            // Step 5: Read value BEFORE refresh
            _output.WriteLine("\n--- Step 5: Read A1 value BEFORE any explicit recalculation ---");
            object valueBefore = range.Value2;
            _output.WriteLine($"A1 Value (raw): {valueBefore}");
            _output.WriteLine($"A1 Value type: {valueBefore?.GetType().Name ?? "null"}");

            // Check for error codes
            if (valueBefore is int or double)
            {
                double numVal = Convert.ToDouble(valueBefore, System.Globalization.CultureInfo.InvariantCulture);
                if (numVal < 0)
                {
                    _output.WriteLine($"⚠️ NEGATIVE VALUE - likely Excel error code!");
                    DescribeExcelErrorCode(numVal);
                }
                else
                {
                    _output.WriteLine($"✅ Numeric value: {numVal} (expected ~450 from SUM of 100+200+150)");
                }
            }

            // Step 6: Refresh Data Model
            _output.WriteLine("\n--- Step 6: Refresh Data Model ---");
            model = _workbook.Model;
            try
            {
                model.Refresh();
                _output.WriteLine("model.Refresh() succeeded");
            }
            catch (COMException ex)
            {
                _output.WriteLine($"model.Refresh() failed: 0x{ex.HResult:X8} - {ex.Message}");
                // Try connection refresh as fallback
                _output.WriteLine("Attempting connection refresh instead...");
                RefreshAllConnections();
            }

            // Step 7: Read value AFTER refresh (without explicit Calculate)
            _output.WriteLine("\n--- Step 7: Read A1 value AFTER refresh (no explicit Calculate) ---");
            object valueAfterRefresh = range.Value2;
            _output.WriteLine($"A1 Value (raw): {valueAfterRefresh}");
            _output.WriteLine($"A1 Value type: {valueAfterRefresh?.GetType().Name ?? "null"}");

            if (valueAfterRefresh is int or double)
            {
                double numVal = Convert.ToDouble(valueAfterRefresh, System.Globalization.CultureInfo.InvariantCulture);
                if (numVal < 0)
                {
                    _output.WriteLine($"⚠️ STILL ERROR CODE after refresh!");
                    DescribeExcelErrorCode(numVal);
                }
                else
                {
                    _output.WriteLine($"✅ Numeric value: {numVal}");
                }
            }

            // Step 8: Call Application.Calculate and re-read
            _output.WriteLine("\n--- Step 8: Call Application.Calculate() ---");
            _excel.Calculate();
            _output.WriteLine("Application.Calculate() called");

            object valueAfterCalculate = range.Value2;
            _output.WriteLine($"A1 Value (raw): {valueAfterCalculate}");
            _output.WriteLine($"A1 Value type: {valueAfterCalculate?.GetType().Name ?? "null"}");

            if (valueAfterCalculate is int or double)
            {
                double numVal = Convert.ToDouble(valueAfterCalculate, System.Globalization.CultureInfo.InvariantCulture);
                if (numVal < 0)
                {
                    _output.WriteLine($"⚠️ STILL ERROR CODE after Calculate!");
                    DescribeExcelErrorCode(numVal);
                }
                else
                {
                    _output.WriteLine($"✅ Numeric value: {numVal}");
                }
            }

            // Step 9: Call Application.CalculateFull and re-read
            _output.WriteLine("\n--- Step 9: Call Application.CalculateFull() ---");
            _excel.CalculateFull();
            _output.WriteLine("Application.CalculateFull() called");

            object valueAfterFullCalc = range.Value2;
            _output.WriteLine($"A1 Value (raw): {valueAfterFullCalc}");
            _output.WriteLine($"A1 Value type: {valueAfterFullCalc?.GetType().Name ?? "null"}");

            if (valueAfterFullCalc is int or double)
            {
                double numVal = Convert.ToDouble(valueAfterFullCalc, System.Globalization.CultureInfo.InvariantCulture);
                if (numVal < 0)
                {
                    _output.WriteLine($"⚠️ STILL ERROR CODE after CalculateFull!");
                    DescribeExcelErrorCode(numVal);
                }
                else
                {
                    _output.WriteLine($"✅ Numeric value: {numVal} (expected ~450)");
                }
            }

            // Step 10: Also check A2 (Query - Sales connection)
            _output.WriteLine("\n--- Step 10: Check A2-A5 (different formula formats) ---");
            dynamic? rangeA2Check = sheet.Range["A2"];
            dynamic? rangeA3Check = sheet.Range["A3"];
            dynamic? rangeA4Check = sheet.Range["A4"];
            dynamic? rangeA5Check = sheet.Range["A5"];

            object a2Value = rangeA2Check.Value2;
            object a3Value = rangeA3Check.Value2;
            object a4Value = rangeA4Check.Value2;
            object a5Value = rangeA5Check.Value2;

            _output.WriteLine($"A2 (Query - Sales + [Measures].[TotalAmount]): {FormatValue(a2Value)}");
            _output.WriteLine($"A3 (ThisWorkbookDataModel + TotalAmount): {FormatValue(a3Value)}");
            _output.WriteLine($"A4 (ThisWorkbookDataModel + [Sales].[TotalAmount]): {FormatValue(a4Value)}");
            _output.WriteLine($"A5 (CUBEMEMBER [Measures].[TotalAmount]): {FormatValue(a5Value)}");

            ComUtilities.Release(ref rangeA2Check);
            ComUtilities.Release(ref rangeA3Check);
            ComUtilities.Release(ref rangeA4Check);
            ComUtilities.Release(ref rangeA5Check);

            // Step 11: Try CalculateFullRebuild (Ctrl+Alt+Shift+F9)
            _output.WriteLine("\n--- Step 11: Call Application.CalculateFullRebuild() ---");
            try
            {
                _excel.CalculateFullRebuild();
                _output.WriteLine("Application.CalculateFullRebuild() called");

                object valueAfterRebuild = range.Value2;
                _output.WriteLine($"A1 Value (raw): {valueAfterRebuild}");
                if (valueAfterRebuild is int or double)
                {
                    double numVal = Convert.ToDouble(valueAfterRebuild, System.Globalization.CultureInfo.InvariantCulture);
                    if (numVal < 0)
                    {
                        _output.WriteLine($"⚠️ STILL ERROR CODE after CalculateFullRebuild!");
                        DescribeExcelErrorCode(numVal);
                    }
                    else
                    {
                        _output.WriteLine($"✅ Numeric value: {numVal}");
                    }
                }
            }
            catch (COMException ex)
            {
                _output.WriteLine($"CalculateFullRebuild failed: {ex.Message}");
            }

            // Step 12: Save workbook
            // NOTE: The "close and reopen" part was never implemented, so we just save
            _output.WriteLine("\n--- Step 12: Save workbook ---");
            _workbook.Save();
            _output.WriteLine("Workbook saved");

            // Check value after save
            object valueAfterSave = range.Value2;
            _output.WriteLine($"A1 Value after save: {valueAfterSave}");

            // Summary
            _output.WriteLine("\n--- Summary ---");
            _output.WriteLine($"Before refresh:       {valueBefore}");
            _output.WriteLine($"After refresh:        {valueAfterRefresh}");
            _output.WriteLine($"After Calculate():    {valueAfterCalculate}");
            _output.WriteLine($"After CalculateFull():{valueAfterFullCalc}");
            _output.WriteLine($"After Save:           {valueAfterSave}");
            _output.WriteLine("");
            _output.WriteLine("CONCLUSION: If all values are negative error codes, CUBEVALUE");
            _output.WriteLine("formulas may require the workbook to be opened in visible Excel");
            _output.WriteLine("with Data Model initialized before formulas can calculate.");

            // Step 13: Try with Excel Visible
            _output.WriteLine("\n--- Step 13: Set Excel.Visible = true and recalculate ---");
            _excel.Visible = true;
            Thread.Sleep(2000);  // Give Excel time to render
            _excel.CalculateFullRebuild();
            Thread.Sleep(1000);
            object valueAfterVisible = range.Value2;
            _output.WriteLine($"A1 Value with Excel Visible: {valueAfterVisible}");
            if (valueAfterVisible is int or double)
            {
                double numVal = Convert.ToDouble(valueAfterVisible, System.Globalization.CultureInfo.InvariantCulture);
                if (numVal < 0)
                {
                    _output.WriteLine($"⚠️ STILL ERROR CODE even with Excel visible!");
                    DescribeExcelErrorCode(numVal);
                }
                else
                {
                    _output.WriteLine($"✅ SUCCESS! Value: {numVal} - Visible Excel fixed it!");
                }
            }
            _excel.Visible = false;  // Hide again for cleanup
        }
        catch (COMException ex)
        {
            _output.WriteLine($"Test failed with COM exception: 0x{ex.HResult:X8}");
            _output.WriteLine($"Message: {ex.Message}");
        }
        finally
        {
            ComUtilities.Release(ref range);
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref model);
        }

        _output.WriteLine("=== SCENARIO 12 COMPLETE ===\n");
    }

    private static string FormatValue(object? value)
    {
        if (value == null) return "null";
        if (value is int or double)
        {
            double numVal = Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture);
            if (numVal < 0)
            {
                int code = Convert.ToInt32(numVal);
                string errName = code switch
                {
                    -2146826288 => "#NULL!",
                    -2146826281 => "#DIV/0!",
                    -2146826246 => "#VALUE!",
                    -2146826259 => "#REF!",
                    -2146826252 => "#NAME?",
                    -2146826265 => "#NUM!",
                    -2146826245 => "#N/A",
                    _ => $"Error {code}"
                };
                return $"{code} ({errName})";
            }
            return $"{numVal} ✅";
        }
        return value.ToString() ?? "empty";
    }

    private void DescribeExcelErrorCode(double errorCode)
    {
        int code = Convert.ToInt32(errorCode);
        string description = code switch
        {
            -2146826288 => "#NULL! - Incorrect range operator or missing intersection",
            -2146826281 => "#DIV/0! - Division by zero",
            -2146826246 => "#VALUE! - Wrong argument type",
            -2146826259 => "#REF! - Invalid cell reference",
            -2146826252 => "#NAME? - Unrecognized formula name",
            -2146826265 => "#NUM! - Invalid numeric value",
            -2146826245 => "#N/A - Value not found, CUBE member not found, or Data Model not refreshed/calculated",
            _ => $"Unknown error code: {code}"
        };
        _output.WriteLine($"  Error code {code} = {description}");
    }

    private void RefreshAllConnections()
    {
        dynamic? connections = null;
        try
        {
            connections = _workbook.Connections;
            for (int i = 1; i <= connections.Count; i++)
            {
                dynamic? conn = connections.Item(i);
                try
                {
                    conn.Refresh();
                    _output.WriteLine($"  Refreshed connection: {conn.Name}");
                }
                catch (COMException ex)
                {
                    _output.WriteLine($"  Failed to refresh {conn.Name}: {ex.Message}");
                }
                finally
                {
                    ComUtilities.Release(ref conn);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref connections);
        }
    }

    // =========================================================================
    // Helper Methods
    // =========================================================================

    private void CreateDataModelTable(string queryName, string mCode)
    {
        dynamic? queries = null;
        dynamic? query = null;
        dynamic? connections = null;
        dynamic? conn = null;

        try
        {
            queries = _workbook.Queries;

            // Check if query already exists
            for (int i = 1; i <= queries.Count; i++)
            {
                dynamic? q = queries.Item(i);
                if (q.Name == queryName)
                {
                    ComUtilities.Release(ref q);
                    _output.WriteLine($"Query '{queryName}' already exists");
                    return;
                }
                ComUtilities.Release(ref q);
            }

            query = queries.Add(queryName, mCode);

            connections = _workbook.Connections;
            string connString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";

            conn = connections.Add2(
                $"Query - {queryName}",
                $"Power Query - {queryName}",
                connString,
                $"SELECT * FROM [{queryName}]",
                2,      // xlCmdSql
                true,   // CreateModelConnection
                false   // ImportRelationships
            );
            conn.Refresh();
            _output.WriteLine($"Created and loaded '{queryName}' to Data Model");
        }
        catch (COMException ex)
        {
            _output.WriteLine($"CreateDataModelTable failed: {ex.Message}");
        }
        finally
        {
            ComUtilities.Release(ref conn);
            ComUtilities.Release(ref connections);
            ComUtilities.Release(ref query);
            ComUtilities.Release(ref queries);
        }
    }

    private void CreateMeasure(string name, string formula)
    {
        dynamic? model = null;
        dynamic? modelTables = null;
        dynamic? measures = null;
        dynamic? table = null;
        dynamic? formatInfo = null;

        try
        {
            model = _workbook.Model;
            modelTables = model.ModelTables;

            if (modelTables.Count == 0)
            {
                _output.WriteLine($"Cannot create measure '{name}': No tables in model");
                return;
            }

            table = modelTables.Item(1);
            measures = model.ModelMeasures;

            // ModelMeasures.Add signature: (MeasureName, AssociatedTable, Formula, FormatInformation, [Description])
            // FormatInformation is REQUIRED - get from Model.ModelFormatGeneral property
            formatInfo = model.ModelFormatGeneral;
            dynamic? measure = measures.Add(name, table, formula, formatInfo);
            _output.WriteLine($"Created measure '{name}'");
            ComUtilities.Release(ref measure);
        }
        catch (COMException ex)
        {
            _output.WriteLine($"CreateMeasure failed: {ex.Message}");
        }
        finally
        {
            ComUtilities.Release(ref formatInfo);
            ComUtilities.Release(ref table);
            ComUtilities.Release(ref measures);
            ComUtilities.Release(ref modelTables);
            ComUtilities.Release(ref model);
        }
    }

    private void CreateRelationshipBetweenTables()
    {
        dynamic? model = null;
        dynamic? modelTables = null;
        dynamic? relationships = null;

        try
        {
            model = _workbook.Model;
            modelTables = model.ModelTables;
            relationships = model.ModelRelationships;

            dynamic? salesTable = null;
            dynamic? productsTable = null;

            for (int i = 1; i <= modelTables.Count; i++)
            {
                dynamic? t = modelTables.Item(i);
                string name = t.Name;
                if (name == "Sales")
                    salesTable = t;
                else if (name == "Products")
                    productsTable = t;
                else
                    ComUtilities.Release(ref t);
            }

            if (salesTable == null || productsTable == null)
            {
                _output.WriteLine("Cannot create relationship: Missing tables");
                ComUtilities.Release(ref salesTable);
                ComUtilities.Release(ref productsTable);
                return;
            }

            // Find columns
            dynamic? salesProductCol = null;
            dynamic? productsNameCol = null;

            dynamic? salesCols = salesTable.ModelTableColumns;
            for (int i = 1; i <= salesCols.Count; i++)
            {
                dynamic? col = salesCols.Item(i);
                if (col.Name == "Product")
                {
                    salesProductCol = col;
                    break;
                }
                ComUtilities.Release(ref col);
            }
            ComUtilities.Release(ref salesCols);

            dynamic? prodCols = productsTable.ModelTableColumns;
            for (int i = 1; i <= prodCols.Count; i++)
            {
                dynamic? col = prodCols.Item(i);
                if (col.Name == "ProductName")
                {
                    productsNameCol = col;
                    break;
                }
                ComUtilities.Release(ref col);
            }
            ComUtilities.Release(ref prodCols);

            if (salesProductCol != null && productsNameCol != null)
            {
                dynamic? rel = relationships.Add(salesProductCol, productsNameCol);
                _output.WriteLine("Created relationship Sales[Product] -> Products[ProductName]");
                ComUtilities.Release(ref rel);
            }

            ComUtilities.Release(ref salesProductCol);
            ComUtilities.Release(ref productsNameCol);
            ComUtilities.Release(ref salesTable);
            ComUtilities.Release(ref productsTable);
        }
        catch (COMException ex)
        {
            _output.WriteLine($"CreateRelationship failed: {ex.Message}");
        }
        finally
        {
            ComUtilities.Release(ref relationships);
            ComUtilities.Release(ref modelTables);
            ComUtilities.Release(ref model);
        }
    }

    // =========================================================================
    // SCENARIO 13: CUBEVALUE with fresh Excel instance (Hidden vs Visible)
    // Runs CUBEVALUE test with a completely fresh Excel instance
    // =========================================================================

    [Theory]
    [InlineData(false, "Hidden")]
    [InlineData(true, "Visible")]
    public void Scenario13_CubeValueFreshExcelInstance(bool excelVisible, string visibilityLabel)
    {
        _output.WriteLine($"=== SCENARIO 13: CUBEVALUE Fresh Excel Instance ({visibilityLabel}) ===");
        _output.WriteLine($"Excel.Visible = {excelVisible}");
        _output.WriteLine("");

        dynamic? excel = null;
        dynamic? workbook = null;
        dynamic? model = null;
        dynamic? sheet = null;
        dynamic? range = null;
        string testFile = Path.Combine(_tempDir, $"CubeValue_{visibilityLabel}_{Guid.NewGuid():N}.xlsx");

        try
        {
            // Step 1: Create fresh Excel instance
            _output.WriteLine("--- Step 1: Create fresh Excel instance ---");
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            excel = Activator.CreateInstance(excelType!);
            excel.Visible = excelVisible;
            excel.DisplayAlerts = false;
            _output.WriteLine($"Excel instance created. Visible={excel.Visible}");

            if (excelVisible)
            {
                // Give Excel time to fully initialize when visible
                Thread.Sleep(2000);
            }

            // Step 2: Create workbook
            _output.WriteLine("\n--- Step 2: Create workbook ---");
            workbook = excel.Workbooks.Add();
            workbook.SaveAs(testFile);
            _output.WriteLine($"Workbook created: {testFile}");

            // Step 3: Create Power Query and load to Data Model
            _output.WriteLine("\n--- Step 3: Create Power Query with Data Model load ---");
            string queryName = "Sales";
            string mCode = """
                let
                    Source = #table(
                        {"Product", "Amount", "Quantity"},
                        {{"Widget", 100, 5}, {"Gadget", 200, 3}, {"Gizmo", 150, 7}}
                    )
                in
                    Source
                """;

            dynamic? queries = workbook.Queries;
            dynamic? query = queries.Add(queryName, mCode);
            _output.WriteLine($"Query '{queryName}' created");

            // Create connection and load to Data Model
            string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
            string commandText = $"SELECT * FROM [{queryName}]";

            dynamic? connections = workbook.Connections;
            dynamic? conn = connections.Add2(
                $"Query - {queryName}",
                $"Power Query - {queryName}",
                connectionString,
                commandText,
                2,    // xlCmdSql
                true, // CreateModelConnection - LOAD TO DATA MODEL
                false // ImportRelationships
            );
            _output.WriteLine("Connection created with CreateModelConnection=true");

            // Refresh to load data
            conn.Refresh();
            _output.WriteLine("Connection refreshed");
            ComUtilities.Release(ref conn);
            ComUtilities.Release(ref connections);
            ComUtilities.Release(ref query);
            ComUtilities.Release(ref queries);

            if (excelVisible)
            {
                Thread.Sleep(1000);
            }

            // Step 4: Create DAX measure
            _output.WriteLine("\n--- Step 4: Create DAX measure ---");
            model = workbook.Model;
            dynamic? modelTables = model.ModelTables;
            _output.WriteLine($"ModelTables.Count: {modelTables.Count}");
            dynamic? targetTable = null;
            string targetTableName = "";

            for (int i = 1; i <= modelTables.Count; i++)
            {
                dynamic? t = modelTables.Item(i);
                string tName = t.Name?.ToString() ?? "(null)";
                _output.WriteLine($"  Table[{i}]: {tName}");
                if (targetTable == null)
                {
                    targetTable = t;
                    targetTableName = tName;
                }
                else
                {
                    ComUtilities.Release(ref t);
                }
            }
            ComUtilities.Release(ref modelTables);

            if (targetTable != null)
            {
                dynamic? measures = model.ModelMeasures;
                dynamic? formatInfo = model.ModelFormatGeneral;  // REQUIRED
                try
                {
                    // Use the actual table name in the DAX formula
                    string daxFormula = $"SUM({targetTableName}[Amount])";
                    dynamic? measure = measures.Add(
                        "TotalAmount",
                        targetTable,
                        daxFormula,
                        formatInfo
                    );
                    _output.WriteLine($"Created measure 'TotalAmount' = {daxFormula}");
                    ComUtilities.Release(ref measure);
                }
                catch (COMException ex)
                {
                    _output.WriteLine($"Failed to create measure: 0x{ex.HResult:X} - {ex.Message}");
                }
                finally
                {
                    ComUtilities.Release(ref formatInfo);
                    ComUtilities.Release(ref measures);
                }
            }
            else
            {
                _output.WriteLine("ERROR: No tables found in model!");
            }
            ComUtilities.Release(ref targetTable);

            // Step 5: Discover Data Model connection name
            _output.WriteLine("\n--- Step 5: Discover Data Model connection name ---");
            dynamic? dataModelConn = null;
            string dataModelConnName = "ThisWorkbookDataModel";
            try
            {
                dataModelConn = model.DataModelConnection;
                dataModelConnName = dataModelConn?.Name?.ToString() ?? "ThisWorkbookDataModel";
                _output.WriteLine($"Model.DataModelConnection.Name: {dataModelConnName}");
            }
#pragma warning disable CA1031 // Intentional: diagnostic test logs all exceptions
            catch (Exception ex)
#pragma warning restore CA1031
            {
                _output.WriteLine($"Could not get DataModelConnection: {ex.Message}");
            }
            finally
            {
                ComUtilities.Release(ref dataModelConn);
            }

            // Also list all connections
            _output.WriteLine("All connections in workbook:");
            dynamic? conns = workbook.Connections;
            for (int i = 1; i <= conns.Count; i++)
            {
                dynamic? c = conns.Item(i);
                string cName = c?.Name?.ToString() ?? "(null)";
                int cType = Convert.ToInt32(c?.Type ?? 0);
                bool inModel = false;
                try { inModel = c.InModel; } catch (COMException) { /* InModel property may not exist on all connection types */ }
                _output.WriteLine($"  [{i}] Name: '{cName}', Type: {cType}, InModel: {inModel}");
                ComUtilities.Release(ref c);
            }
            ComUtilities.Release(ref conns);

            // Step 5b: Try to create a proper Model Workbook Connection (per MS docs)
            _output.WriteLine("\n--- Step 5b: Try Model.CreateModelWorkbookConnection ---");
            dynamic? modelWorkbookConn = null;
            try
            {
                // Get the first table from the model to use as parameter
                dynamic? modelTables2 = model.ModelTables;
                dynamic? firstTable = modelTables2.Item(1);
                modelWorkbookConn = model.CreateModelWorkbookConnection(firstTable);
                string connName = modelWorkbookConn?.Name?.ToString() ?? "(null)";
                _output.WriteLine($"Created Model Workbook Connection: '{connName}'");
                ComUtilities.Release(ref firstTable);
                ComUtilities.Release(ref modelTables2);
            }
            catch (COMException ex)
            {
                _output.WriteLine($"CreateModelWorkbookConnection failed: 0x{ex.HResult:X8} - {ex.Message}");
            }
            finally
            {
                ComUtilities.Release(ref modelWorkbookConn);
            }

            // List connections again after attempting to create model connection
            _output.WriteLine("Connections after CreateModelWorkbookConnection:");
            conns = workbook.Connections;
            for (int i = 1; i <= conns.Count; i++)
            {
                dynamic? c = conns.Item(i);
                string cName = c?.Name?.ToString() ?? "(null)";
                int cType = Convert.ToInt32(c?.Type ?? 0);
                _output.WriteLine($"  [{i}] Name: '{cName}', Type: {cType}");
                ComUtilities.Release(ref c);
            }
            ComUtilities.Release(ref conns);

            // Step 6: Add CUBEVALUE formula
            _output.WriteLine("\n--- Step 6: Add CUBEVALUE formulas (testing different syntax) ---");
            sheet = workbook.Worksheets.Item(1);

            // Build formulas with discovered names
            // Important: DAX measures are referenced as [Measures].[MeasureName]
            // According to MS docs, the special Data Model connection is named "Workbook Data Model" (with spaces)
            string[] formulas = new[]
            {
                $"=CUBEVALUE(\"{dataModelConnName}\",\"[Measures].[TotalAmount]\")",
                $"=CUBEVALUE(\"{dataModelConnName}\",\"{targetTableName}[TotalAmount]\")",
                "=CUBEVALUE(\"Workbook Data Model\",\"[Measures].[TotalAmount]\")",
                "=CUBEVALUE(\"ThisWorkbookDataModel\",\"[Measures].[TotalAmount]\")",
                "=CUBEVALUE(\"Query - Sales\",\"[Measures].[TotalAmount]\")",
                $"=CUBEVALUE(\"{dataModelConnName}\",\"[{targetTableName}].[Measures].[TotalAmount]\")",
            };

            string[] formulaDescriptions = new[]
            {
                $"DataModelConnection.Name ({dataModelConnName})",
                "TableName[MeasureName] pattern",
                "Workbook Data Model (MS docs name)",
                "ThisWorkbookDataModel (hardcoded)",
                "Query - Sales connection",
                "[TableName].[Measures].[MeasureName] pattern",
            };

            for (int i = 0; i < formulas.Length; i++)
            {
                string cell = $"A{i + 1}";
                range = sheet.Range[cell];
                range.Formula = formulas[i];
                _output.WriteLine($"{cell}: {formulaDescriptions[i]}");
                _output.WriteLine($"     Formula: {formulas[i]}");
                ComUtilities.Release(ref range);
            }

            // Also test CUBEMEMBER to see if that works
            _output.WriteLine("\n--- Also testing CUBEMEMBER function ---");
            dynamic? memberRange = sheet.Range["B1"];
            string cubeMemberFormula = $"=CUBEMEMBER(\"{dataModelConnName}\",\"[Measures].[TotalAmount]\")";
            memberRange.Formula = cubeMemberFormula;
            _output.WriteLine($"B1: CUBEMEMBER formula: {cubeMemberFormula}");
            ComUtilities.Release(ref memberRange);

            range = sheet.Range["A1"];  // Reset for value reading

            // Step 6b: Read values immediately
            _output.WriteLine("\n--- Step 6b: Read values immediately ---");
            for (int i = 0; i < formulas.Length; i++)
            {
                string cell = $"A{i + 1}";
                dynamic? r = sheet.Range[cell];
                object val = r.Value2;
                _output.WriteLine($"{cell}: {FormatValue(val)} - {formulaDescriptions[i]}");
                ComUtilities.Release(ref r);
            }

            // Check CUBEMEMBER
            dynamic? memberRangeCheck = sheet.Range["B1"];
            object memberVal = memberRangeCheck.Value2;
            _output.WriteLine($"B1 (CUBEMEMBER): {FormatValue(memberVal)}");
            ComUtilities.Release(ref memberRangeCheck);

            // Step 7: Refresh Data Model
            _output.WriteLine("\n--- Step 7: Refresh Data Model ---");
            try
            {
                model.Refresh();
                _output.WriteLine("model.Refresh() succeeded");
            }
            catch (COMException ex)
            {
                _output.WriteLine($"model.Refresh() failed: 0x{ex.HResult:X8}");
            }

            if (excelVisible)
            {
                Thread.Sleep(1000);
            }

            _output.WriteLine("Values after Data Model refresh:");
            for (int i = 0; i < formulas.Length; i++)
            {
                string cell = $"A{i + 1}";
                dynamic? r = sheet.Range[cell];
                object val = r.Value2;
                _output.WriteLine($"  {cell}: {FormatValue(val)}");
                ComUtilities.Release(ref r);
            }

            // Step 8: Calculate
            _output.WriteLine("\n--- Step 8: Application.Calculate() ---");
            try
            {
                excel.Calculate();
                _output.WriteLine("Calculate() succeeded");
            }
            catch (COMException ex)
            {
                _output.WriteLine($"Calculate() failed: 0x{ex.HResult:X8} - {ex.Message}");
            }

            if (excelVisible)
            {
                Thread.Sleep(500);
            }

            _output.WriteLine("Values after Calculate:");
            for (int i = 0; i < formulas.Length; i++)
            {
                string cell = $"A{i + 1}";
                dynamic? r = sheet.Range[cell];
                object val = r.Value2;
                _output.WriteLine($"  {cell}: {FormatValue(val)}");
                ComUtilities.Release(ref r);
            }

            // Step 9: CalculateFull
            _output.WriteLine("\n--- Step 9: Application.CalculateFull() ---");
            try
            {
                excel.CalculateFull();
                _output.WriteLine("CalculateFull() succeeded");
            }
            catch (COMException ex)
            {
                _output.WriteLine($"CalculateFull() failed: 0x{ex.HResult:X8} - {ex.Message}");
            }

            if (excelVisible)
            {
                Thread.Sleep(500);
            }

            _output.WriteLine("Values after CalculateFull:");
            for (int i = 0; i < formulas.Length; i++)
            {
                string cell = $"A{i + 1}";
                dynamic? r = sheet.Range[cell];
                object val = r.Value2;
                _output.WriteLine($"  {cell}: {FormatValue(val)}");
                ComUtilities.Release(ref r);
            }

            // Step 10: CalculateFullRebuild
            _output.WriteLine("\n--- Step 10: Application.CalculateFullRebuild() ---");
            try
            {
                excel.CalculateFullRebuild();
                _output.WriteLine("CalculateFullRebuild() succeeded");
            }
            catch (COMException ex)
            {
                _output.WriteLine($"CalculateFullRebuild() failed: 0x{ex.HResult:X8} - {ex.Message}");
            }

            if (excelVisible)
            {
                Thread.Sleep(500);
            }

            _output.WriteLine("Values after CalculateFullRebuild:");
            object? bestValue = null;
            for (int i = 0; i < formulas.Length; i++)
            {
                string cell = $"A{i + 1}";
                dynamic? r = sheet.Range[cell];
                object val = r.Value2;
                _output.WriteLine($"  {cell}: {FormatValue(val)} - {formulaDescriptions[i]}");
                if (val is int or double && bestValue == null)
                {
                    bestValue = val;
                }
                ComUtilities.Release(ref r);
            }

            // Also check CUBEMEMBER
            dynamic? memberRangeFinal = sheet.Range["B1"];
            object memberValFinal = memberRangeFinal.Value2;
            _output.WriteLine($"  B1 (CUBEMEMBER): {FormatValue(memberValFinal)}");
            ComUtilities.Release(ref memberRangeFinal);

            // Summary
            _output.WriteLine("\n=== SUMMARY ===");
            _output.WriteLine($"Excel.Visible: {excelVisible}");
            _output.WriteLine($"Data Model Connection: {dataModelConnName}");
            _output.WriteLine($"Target Table: {targetTableName}");
            _output.WriteLine($"Best numeric result: {(bestValue != null ? bestValue.ToString() : "None - all formulas returned errors")}");

            // Determine success
            bool success = bestValue != null;
            _output.WriteLine(success ? "\n✅ At least one CUBEVALUE formula returned a numeric value" : "\n❌ ALL CUBEVALUE formulas still show errors");
        }
        catch (COMException ex)
        {
            _output.WriteLine($"\n❌ Test failed with COM exception: 0x{ex.HResult:X8}");
            _output.WriteLine($"Message: {ex.Message}");
        }
        finally
        {
            _output.WriteLine("\n--- Cleanup ---");
            ComUtilities.Release(ref range);
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref model);

            if (workbook != null)
            {
                try
                {
                    workbook.Close(false);
                }
#pragma warning disable CA1031 // Intentional: cleanup code must not throw
                catch (Exception) { /* Ignore cleanup errors */ }
#pragma warning restore CA1031
                ComUtilities.Release(ref workbook);
            }

            if (excel != null)
            {
                try
                {
                    excel.Quit();
                }
#pragma warning disable CA1031 // Intentional: cleanup code must not throw
                catch (Exception) { /* Ignore cleanup errors */ }
#pragma warning restore CA1031
                ComUtilities.Release(ref excel);
            }

            // Clean up test file
            try
            {
                if (File.Exists(testFile))
                    File.Delete(testFile);
            }
#pragma warning disable CA1031 // Intentional: cleanup code must not throw
            catch (Exception) { /* Ignore file cleanup errors */ }
#pragma warning restore CA1031
        }

        _output.WriteLine($"=== SCENARIO 13 ({visibilityLabel}) COMPLETE ===\n");
    }
}




