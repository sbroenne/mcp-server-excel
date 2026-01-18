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

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Diagnostics;

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

    // =========================================================================
    // SCENARIO 14: Execute DAX EVALUATE Query via xlSrcModel + xlCmdDAX
    // =========================================================================
    // This tests the approach proposed in GitHub Issue #356:
    // https://github.com/sbroenne/mcp-server-excel/issues/356
    //
    // The hypothesis is that using ListObjects.Add with xlSrcModel (4)
    // and ModelConnection.CommandType = xlCmdDAX (8) may work differently
    // from CUBEVALUE worksheet functions, because it goes through Excel's
    // table refresh mechanism rather than worksheet formula evaluation.
    // =========================================================================

    [Fact]
    public void Scenario14_DaxEvaluateQuery_ViaListObjectXlSrcModel()
    {
        _output.WriteLine("=== SCENARIO 14: DAX EVALUATE via xlSrcModel + xlCmdDAX ===");
        _output.WriteLine("Testing approach from GitHub Issue #356");
        _output.WriteLine("https://github.com/sbroenne/mcp-server-excel/issues/356\n");

        // Constants for Excel enums
        const int xlSrcModel = 4;        // XlListObjectSourceType.xlSrcModel
        const int xlCmdDAX = 8;          // XlCmdType.xlCmdDAX
        const int xlYes = 1;             // XlYesNoGuess.xlYes

        dynamic? excel = null;
        dynamic? workbook = null;
        dynamic? sheet = null;
        dynamic? model = null;
        dynamic? queries = null;
        dynamic? query = null;
        dynamic? connections = null;
        dynamic? listObjects = null;
        dynamic? listObject = null;
        dynamic? tableObject = null;
        dynamic? wbConnection = null;
        dynamic? modelConnection = null;
        dynamic? dataModelConnection = null;

        string testFile = Path.Combine(_tempDir, $"DMDiag_Scenario14_{Guid.NewGuid():N}.xlsx");

        try
        {
            // Step 1: Create Excel instance
            _output.WriteLine("--- Step 1: Create Excel instance ---");
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            excel = Activator.CreateInstance(excelType!);
            excel.Visible = false;
            excel.DisplayAlerts = false;
            _output.WriteLine("Excel instance created (hidden mode)");

            // Step 2: Create workbook and save
            _output.WriteLine("\n--- Step 2: Create workbook ---");
            workbook = excel.Workbooks.Add();
            workbook.SaveAs(testFile);
            _output.WriteLine($"Workbook saved: {testFile}");

            sheet = workbook.Worksheets[1];
            _output.WriteLine($"Active sheet: {sheet.Name}");

            // Step 3: Create Power Query and load to Data Model
            _output.WriteLine("\n--- Step 3: Create Power Query 'Sales' and load to Data Model ---");
            queries = workbook.Queries;
            query = queries.Add("Sales", SalesQuery);
            _output.WriteLine("Power Query 'Sales' created");

            connections = workbook.Connections;
            string connString = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Sales";

            dynamic? conn = null;
            try
            {
                conn = connections.Add2(
                    "Query - Sales",
                    "Power Query - Sales",
                    connString,
                    "SELECT * FROM [Sales]",
                    2,      // xlCmdSql
                    true,   // CreateModelConnection = true (load to Data Model)
                    false
                );
                _output.WriteLine("Connection added with CreateModelConnection=true");
                conn.Refresh();
                _output.WriteLine("Connection refreshed - data loaded to Data Model");
                ComUtilities.Release(ref conn);
            }
            catch (COMException ex)
            {
                _output.WriteLine($"Failed to create model connection: 0x{ex.HResult:X8} - {ex.Message}");
                ComUtilities.Release(ref conn);
                throw;
            }

            // Step 4: Verify Data Model has the table
            _output.WriteLine("\n--- Step 4: Verify Data Model ---");
            model = workbook.Model;
            dynamic? modelTables = model.ModelTables;
            int tableCount = modelTables.Count;
            _output.WriteLine($"Model tables count: {tableCount}");

            string? targetTableName = null;
            for (int i = 1; i <= tableCount; i++)
            {
                dynamic? table = modelTables.Item(i);
                string tableName = table.Name;
                int rowCount = table.RecordCount;
                _output.WriteLine($"  Table {i}: '{tableName}' ({rowCount} rows)");
                if (targetTableName == null)
                {
                    targetTableName = tableName;
                }
                ComUtilities.Release(ref table);
            }
            ComUtilities.Release(ref modelTables);

            if (targetTableName == null)
            {
                _output.WriteLine("❌ No tables in Data Model - cannot proceed");
                return;
            }

            // Step 5: Get DataModelConnection
            _output.WriteLine("\n--- Step 5: Get DataModelConnection ---");
            dataModelConnection = model.DataModelConnection;
            string dmConnName = dataModelConnection.Name;
            _output.WriteLine($"DataModelConnection name: {dmConnName}");

            // Step 6: Attempt to create ListObject with xlSrcModel
            _output.WriteLine("\n--- Step 6: Create ListObject with xlSrcModel ---");
            _output.WriteLine($"Using XlListObjectSourceType.xlSrcModel = {xlSrcModel}");

            listObjects = sheet.ListObjects;
            dynamic? destRange = sheet.Range["E1"];

            try
            {
                // Try: ListObjects.Add(xlSrcModel, DataModelConnection, LinkSource, HasHeaders, Destination)
                listObject = listObjects.Add(
                    xlSrcModel,             // SourceType = xlSrcModel (4)
                    dataModelConnection,    // Source = DataModelConnection
                    true,                   // LinkSource
                    xlYes,                  // HasHeaders
                    destRange               // Destination
                );
                _output.WriteLine("✅ ListObject created successfully with xlSrcModel!");

                // Step 7: Access TableObject and ModelConnection
                _output.WriteLine("\n--- Step 7: Access TableObject.WorkbookConnection.ModelConnection ---");
                tableObject = listObject.TableObject;
                _output.WriteLine($"TableObject obtained");

                wbConnection = tableObject.WorkbookConnection;
                string wbConnName = wbConnection.Name;
                _output.WriteLine($"WorkbookConnection name: {wbConnName}");

                modelConnection = wbConnection.ModelConnection;
                _output.WriteLine("ModelConnection obtained");

                // Check current CommandType
                int currentCmdType = Convert.ToInt32(modelConnection.CommandType);
                _output.WriteLine($"Current CommandType: {currentCmdType} (xlCmdTable=3, xlCmdDAX=8)");

                object currentCmdText = modelConnection.CommandText;
                _output.WriteLine($"Current CommandText: {currentCmdText}");

                // Step 8: Try to set CommandType to xlCmdDAX
                _output.WriteLine("\n--- Step 8: Set CommandType to xlCmdDAX (8) ---");
                try
                {
                    modelConnection.CommandType = xlCmdDAX;
                    _output.WriteLine("✅ CommandType set to xlCmdDAX successfully!");
                }
                catch (COMException ex)
                {
                    _output.WriteLine($"❌ Failed to set CommandType to xlCmdDAX: 0x{ex.HResult:X8}");
                    _output.WriteLine($"   Message: {ex.Message}");
                }

                // Step 9: Set DAX EVALUATE query
                _output.WriteLine("\n--- Step 9: Set DAX EVALUATE query ---");
                // Try various DAX query formats
                string[] daxQueries = new[]
                {
                    $"EVALUATE '{targetTableName}'",
                    $"EVALUATE SUMMARIZECOLUMNS('{targetTableName}'[Product], \"Total\", SUM('{targetTableName}'[Amount]))",
                    $"EVALUATE ROW(\"Total\", SUM('{targetTableName}'[Amount]))",
                    "EVALUATE {1}"  // Simplest possible DAX query
                };

                foreach (string daxQuery in daxQueries)
                {
                    _output.WriteLine($"\nTrying DAX query: {daxQuery}");
                    try
                    {
                        modelConnection.CommandText = daxQuery;
                        _output.WriteLine("  CommandText set successfully");

                        // Step 10: Refresh to execute DAX
                        _output.WriteLine("  Attempting refresh...");
                        listObject.Refresh();
                        _output.WriteLine("  ✅ Refresh succeeded!");

                        // Step 11: Read results
                        dynamic? dataRange = listObject.DataBodyRange;
                        if (dataRange != null)
                        {
                            object[,]? values = dataRange.Value2 as object[,];
                            if (values != null)
                            {
                                int rows = values.GetLength(0);
                                int cols = values.GetLength(1);
                                _output.WriteLine($"  Result: {rows} rows x {cols} columns");
                                for (int r = 1; r <= Math.Min(rows, 5); r++)
                                {
                                    var rowValues = new List<string>();
                                    for (int c = 1; c <= cols; c++)
                                    {
                                        rowValues.Add(values[r, c]?.ToString() ?? "(null)");
                                    }
                                    _output.WriteLine($"    Row {r}: {string.Join(", ", rowValues)}");
                                }
                                if (rows > 5)
                                {
                                    _output.WriteLine($"    ... and {rows - 5} more rows");
                                }
                            }
                            else
                            {
                                object singleValue = dataRange.Value2;
                                _output.WriteLine($"  Single value result: {singleValue}");
                            }
                            ComUtilities.Release(ref dataRange);
                        }
                        else
                        {
                            _output.WriteLine("  DataBodyRange is null (no data returned)");
                        }

                        // If we got here without exception, DAX EVALUATE works!
                        _output.WriteLine("\n🎉 SUCCESS: DAX EVALUATE query execution works via xlSrcModel + xlCmdDAX!");
                        break;
                    }
                    catch (COMException ex)
                    {
                        _output.WriteLine($"  ❌ Failed: 0x{ex.HResult:X8} - {ex.Message}");
                    }
                }
            }
            catch (COMException ex)
            {
                _output.WriteLine($"❌ ListObjects.Add with xlSrcModel failed: 0x{ex.HResult:X8}");
                _output.WriteLine($"   Message: {ex.Message}");

                // Alternative approach: Try creating a new connection with xlCmdDAX directly
                _output.WriteLine("\n--- Alternative: Try Connections.Add2 with xlCmdDAX ---");
                try
                {
                    string daxConnString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={targetTableName}";
                    string daxQuery = $"EVALUATE '{targetTableName}'";

                    dynamic? daxConn = connections.Add2(
                        "DAX Query Connection",
                        "Execute DAX EVALUATE",
                        daxConnString,
                        daxQuery,
                        xlCmdDAX,   // Try xlCmdDAX directly
                        true,       // CreateModelConnection
                        false
                    );
                    _output.WriteLine("✅ Connection created with xlCmdDAX!");

                    daxConn.Refresh();
                    _output.WriteLine("✅ Connection refreshed!");

                    ComUtilities.Release(ref daxConn);
                }
                catch (COMException ex2)
                {
                    _output.WriteLine($"❌ Alternative also failed: 0x{ex2.HResult:X8} - {ex2.Message}");
                }
            }
            finally
            {
                ComUtilities.Release(ref destRange);
            }

            // Summary
            _output.WriteLine("\n=== SCENARIO 14 SUMMARY ===");
            _output.WriteLine("This test explores whether DAX EVALUATE queries can be executed");
            _output.WriteLine("via COM automation using the xlSrcModel + xlCmdDAX approach.");
            _output.WriteLine("Results above indicate whether this is a viable path for Issue #356.");
        }
        catch (Exception ex)
        {
            _output.WriteLine($"\n❌ Unexpected exception: {ex.GetType().Name}");
            _output.WriteLine($"Message: {ex.Message}");
            _output.WriteLine($"StackTrace: {ex.StackTrace}");
        }
        finally
        {
            _output.WriteLine("\n--- Cleanup ---");
            ComUtilities.Release(ref modelConnection);
            ComUtilities.Release(ref wbConnection);
            ComUtilities.Release(ref tableObject);
            ComUtilities.Release(ref listObject);
            ComUtilities.Release(ref listObjects);
            ComUtilities.Release(ref dataModelConnection);
            ComUtilities.Release(ref model);
            ComUtilities.Release(ref query);
            ComUtilities.Release(ref queries);
            ComUtilities.Release(ref connections);
            ComUtilities.Release(ref sheet);

            if (workbook != null)
            {
                try { workbook.Close(false); }
                catch { /* Ignore */ }
                ComUtilities.Release(ref workbook);
            }

            if (excel != null)
            {
                try { excel.Quit(); }
                catch { /* Ignore */ }
                ComUtilities.Release(ref excel);
            }

            try
            {
                if (File.Exists(testFile))
                    File.Delete(testFile);
            }
            catch { /* Ignore */ }
        }

        _output.WriteLine("=== SCENARIO 14 COMPLETE ===\n");
    }

    /// <summary>
    /// Scenario 15: Test DAX EVALUATE via Model.CreateModelWorkbookConnection and ADOConnection
    ///
    /// This tests two alternative approaches discovered in Microsoft documentation:
    /// 1. Model.CreateModelWorkbookConnection - Creates a WorkbookConnection to a specific model table
    /// 2. ModelConnection.ADOConnection - Direct ADO/ADOMD access to execute queries
    ///
    /// Related to GitHub Issue #356
    /// </summary>
    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Layer", "Diagnostics")]
    [Trait("RequiresExcel", "true")]
    [Trait("Feature", "DataModel")]
    [Trait("RunType", "OnDemand")]
    public void Scenario15_DaxQuery_ViaCreateModelWorkbookConnectionAndADO()
    {
        _output.WriteLine("=== SCENARIO 15: DAX via CreateModelWorkbookConnection + ADOConnection ===");
        _output.WriteLine("Testing alternative approaches from Microsoft documentation");
        _output.WriteLine("https://github.com/sbroenne/mcp-server-excel/issues/356\n");

        // Test file with unique name
        string testFile = Path.Combine(_tempDir, $"DMDiag_Scenario15_{Guid.NewGuid():N}.xlsx");
        _output.WriteLine($"Test file: {testFile}");

        dynamic? excel = null;
        dynamic? workbook = null;
        dynamic? sheet = null;
        dynamic? connections = null;
        dynamic? queries = null;
        dynamic? query = null;
        dynamic? model = null;
        dynamic? dataModelConnection = null;
        dynamic? modelConnection = null;
        dynamic? adoConnection = null;
        dynamic? modelWbConn = null;

        try
        {
            // Constants
            const int xlCmdDAX = 8;
            const int xlCmdTable = 3;

            // Step 1: Create Excel instance
            _output.WriteLine("--- Step 1: Create Excel instance ---");
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            excel = Activator.CreateInstance(excelType!);
            excel.Visible = false;
            excel.DisplayAlerts = false;
            _output.WriteLine("Excel instance created (hidden mode)");

            // Step 2: Create workbook
            _output.WriteLine("\n--- Step 2: Create workbook ---");
            dynamic? workbooks = excel.Workbooks;
            workbook = workbooks.Add();
            ComUtilities.Release(ref workbooks);
            workbook.SaveAs(testFile);
            _output.WriteLine($"Workbook saved: {testFile}");

            sheet = workbook.ActiveSheet;
            string sheetName = sheet.Name;
            _output.WriteLine($"Active sheet: {sheetName}");

            connections = workbook.Connections;
            queries = workbook.Queries;

            // Step 3: Create Power Query and load to Data Model
            _output.WriteLine("\n--- Step 3: Create Power Query 'Products' and load to Data Model ---");
            string mCode = @"let
    Source = #table(
        type table [ProductID = Int64.Type, ProductName = text, Price = number],
        {
            {1, ""Widget"", 10.99},
            {2, ""Gadget"", 24.95},
            {3, ""Gizmo"", 15.50}
        }
    )
in
    Source";

            query = queries.Add("Products", mCode);
            _output.WriteLine("Power Query 'Products' created");

            // Load to Data Model via connection
            string connString = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Products;Extended Properties=\"\"";
            dynamic? conn = null;
            try
            {
                conn = connections.Add2(
                    "Query - Products",
                    "Power Query - Products",
                    connString,
                    "SELECT * FROM [Products]",
                    2,      // xlCmdSql
                    true,   // CreateModelConnection = true (load to Data Model)
                    false
                );
                _output.WriteLine("Connection added with CreateModelConnection=true");
                conn.Refresh();
                _output.WriteLine("Connection refreshed - data loaded to Data Model");
                ComUtilities.Release(ref conn);
            }
            catch (COMException ex)
            {
                _output.WriteLine($"Failed to create model connection: 0x{ex.HResult:X8} - {ex.Message}");
                ComUtilities.Release(ref conn);
                throw;
            }

            // Step 4: Verify Data Model has the table
            _output.WriteLine("\n--- Step 4: Verify Data Model ---");
            model = workbook.Model;
            dynamic? modelTables = model.ModelTables;
            int tableCount = modelTables.Count;
            _output.WriteLine($"Model tables count: {tableCount}");

            string? targetTableName = null;
            for (int i = 1; i <= tableCount; i++)
            {
                dynamic? table = modelTables.Item(i);
                string tableName = table.Name;
                int rowCount = table.RecordCount;
                _output.WriteLine($"  Table {i}: '{tableName}' ({rowCount} rows)");
                if (targetTableName == null)
                {
                    targetTableName = tableName;
                }
                ComUtilities.Release(ref table);
            }
            ComUtilities.Release(ref modelTables);

            if (targetTableName == null)
            {
                _output.WriteLine("❌ No tables in Data Model - cannot proceed");
                return;
            }

            // Step 5: Get DataModelConnection and its ModelConnection
            _output.WriteLine("\n--- Step 5: Get DataModelConnection ---");
            dataModelConnection = model.DataModelConnection;
            string dmConnName = dataModelConnection.Name;
            _output.WriteLine($"DataModelConnection name: {dmConnName}");

            // Step 6: Try Model.CreateModelWorkbookConnection
            _output.WriteLine("\n--- Step 6: Try Model.CreateModelWorkbookConnection ---");
            _output.WriteLine($"Creating model workbook connection for table: {targetTableName}");
            try
            {
                modelWbConn = model.CreateModelWorkbookConnection(targetTableName);
                _output.WriteLine($"✅ CreateModelWorkbookConnection succeeded!");

                string connName = modelWbConn.Name;
                _output.WriteLine($"  Connection name: {connName}");

                // Try to access ModelConnection from this connection
                modelConnection = modelWbConn.ModelConnection;
                _output.WriteLine("  ModelConnection obtained");

                int cmdType = Convert.ToInt32(modelConnection.CommandType);
                _output.WriteLine($"  CommandType: {cmdType} (xlCmdTable=3, xlCmdDAX=8)");

                object cmdText = modelConnection.CommandText;
                _output.WriteLine($"  CommandText: {cmdText}");

                // Try to change CommandType to xlCmdDAX
                _output.WriteLine("\n  Attempting to set CommandType to xlCmdDAX (8)...");
                try
                {
                    modelConnection.CommandType = xlCmdDAX;
                    _output.WriteLine("  ✅ CommandType set to xlCmdDAX!");

                    // Try setting a DAX query
                    string daxQuery = $"EVALUATE '{targetTableName}'";
                    _output.WriteLine($"  Setting CommandText to: {daxQuery}");
                    modelConnection.CommandText = daxQuery;
                    _output.WriteLine("  ✅ CommandText set!");

                    // Try refreshing
                    _output.WriteLine("  Attempting refresh...");
                    modelWbConn.Refresh();
                    _output.WriteLine("  ✅ Refresh completed!");

                    _output.WriteLine("\n🎉 SUCCESS: Model.CreateModelWorkbookConnection + xlCmdDAX works!");
                }
                catch (COMException ex)
                {
                    _output.WriteLine($"  ❌ Failed to set CommandType to xlCmdDAX: 0x{ex.HResult:X8}");
                    _output.WriteLine($"     Message: {ex.Message}");

                    // Reset to xlCmdTable
                    try { modelConnection.CommandType = xlCmdTable; } catch { }
                }
            }
            catch (COMException ex)
            {
                _output.WriteLine($"❌ CreateModelWorkbookConnection failed: 0x{ex.HResult:X8}");
                _output.WriteLine($"   Message: {ex.Message}");
            }

            // Step 7: Try ADOConnection approach
            _output.WriteLine("\n--- Step 7: Try ADOConnection approach ---");
            _output.WriteLine("Attempting to get ADOConnection from ModelConnection...");

            try
            {
                // Get fresh DataModelConnection
                if (dataModelConnection == null)
                {
                    dataModelConnection = model.DataModelConnection;
                }

                // Access the ModelConnection
                dynamic? dmModelConn = dataModelConnection.ModelConnection;
                _output.WriteLine($"DataModelConnection.ModelConnection obtained");

                // Check if CommandType is xlCmdCube (expected for DataModelConnection)
                try
                {
                    int cmdType = Convert.ToInt32(dmModelConn.CommandType);
                    _output.WriteLine($"  CommandType: {cmdType}");
                }
                catch (COMException ex)
                {
                    _output.WriteLine($"  CommandType access failed: {ex.Message}");
                }

                // Try to get ADOConnection
                try
                {
                    adoConnection = dmModelConn.ADOConnection;
                    _output.WriteLine("✅ ADOConnection obtained!");

                    // Check connection state
                    try
                    {
                        int state = adoConnection.State;
                        _output.WriteLine($"  ADO Connection State: {state} (1=Open, 0=Closed)");
                    }
                    catch { _output.WriteLine("  Could not read State property"); }

                    // Try to get connection string
                    try
                    {
                        string adoConnStr = adoConnection.ConnectionString;
                        _output.WriteLine($"  ConnectionString: {adoConnStr}");
                    }
                    catch { _output.WriteLine("  Could not read ConnectionString property"); }

                    // Try to execute a DAX query via ADO
                    _output.WriteLine("\n  Attempting to execute DAX query via ADO.Execute...");
                    string daxQuery = $"EVALUATE '{targetTableName}'";
                    _output.WriteLine($"  Query: {daxQuery}");

                    try
                    {
                        dynamic? recordset = adoConnection.Execute(daxQuery);
                        _output.WriteLine("  ✅ Execute succeeded!");

                        // Read results from recordset
                        if (recordset != null)
                        {
                            try
                            {
                                bool eof = recordset.EOF;
                                int fieldCount = recordset.Fields.Count;
                                _output.WriteLine($"  Recordset: EOF={eof}, Fields={fieldCount}");

                                // Read field names
                                var fieldNames = new List<string>();
                                for (int f = 0; f < fieldCount; f++)
                                {
                                    string fieldName = recordset.Fields.Item(f).Name;
                                    fieldNames.Add(fieldName);
                                }
                                _output.WriteLine($"  Fields: {string.Join(", ", fieldNames)}");

                                // Read some rows
                                int rowNum = 0;
                                while (!recordset.EOF && rowNum < 5)
                                {
                                    var rowValues = new List<string>();
                                    for (int f = 0; f < fieldCount; f++)
                                    {
                                        object val = recordset.Fields.Item(f).Value;
                                        rowValues.Add(val?.ToString() ?? "(null)");
                                    }
                                    _output.WriteLine($"  Row {rowNum + 1}: {string.Join(", ", rowValues)}");
                                    recordset.MoveNext();
                                    rowNum++;
                                }

                                _output.WriteLine("\n🎉 SUCCESS: DAX query execution via ADOConnection works!");
                            }
                            catch (Exception readEx)
                            {
                                _output.WriteLine($"  Error reading recordset: {readEx.Message}");
                            }
                            finally
                            {
                                try { recordset.Close(); } catch { }
                                ComUtilities.Release(ref recordset);
                            }
                        }
                    }
                    catch (COMException ex)
                    {
                        _output.WriteLine($"  ❌ Execute failed: 0x{ex.HResult:X8}");
                        _output.WriteLine($"     Message: {ex.Message}");

                        // Try alternative: ADOMD via late binding
                        _output.WriteLine("\n  Trying alternative DAX queries...");
                        string[] queries2 = new[]
                        {
                            "EVALUATE {1}",
                            $"EVALUATE SELECTCOLUMNS('{targetTableName}', \"ID\", [ProductID])",
                            "SELECT * FROM $SYSTEM.DBSCHEMA_TABLES"  // MDX query to list tables
                        };

                        foreach (var q in queries2)
                        {
                            _output.WriteLine($"\n  Trying: {q}");
                            try
                            {
                                dynamic? rs = adoConnection.Execute(q);
                                _output.WriteLine("    ✅ Succeeded!");
                                try { rs.Close(); } catch { }
                                ComUtilities.Release(ref rs);
                            }
                            catch (COMException ex2)
                            {
                                _output.WriteLine($"    ❌ Failed: {ex2.Message}");
                            }
                        }
                    }
                }
                catch (COMException ex)
                {
                    _output.WriteLine($"❌ ADOConnection access failed: 0x{ex.HResult:X8}");
                    _output.WriteLine($"   Message: {ex.Message}");
                }

                ComUtilities.Release(ref dmModelConn);
            }
            catch (Exception ex)
            {
                _output.WriteLine($"❌ ADOConnection approach failed: {ex.GetType().Name}");
                _output.WriteLine($"   Message: {ex.Message}");
            }

            // Step 8: Try PivotCache ADOConnection
            _output.WriteLine("\n--- Step 8: Try PivotCache.ADOConnection approach ---");
            _output.WriteLine("Creating PivotTable connected to Data Model...");

            dynamic? pivotCaches = null;
            dynamic? pivotCache = null;
            dynamic? pivotTable = null;
            dynamic? pivotAdoConn = null;

            try
            {
                pivotCaches = workbook.PivotCaches();

                // Create cache from Data Model (xlExternal with xlPivotTableVersion15)
                const int xlExternal = 2;
                const int xlPivotTableVersion15 = 5; // Excel 2013+

                // Use connection to Data Model
                string pivotConnString = $"Data Model;DSN=Excel Data Model;Provider=MSDASQL;";

                try
                {
                    pivotCache = pivotCaches.Create(
                        xlExternal,
                        dataModelConnection,
                        xlPivotTableVersion15
                    );
                    _output.WriteLine("✅ PivotCache created from DataModelConnection!");

                    // Try to get ADOConnection
                    try
                    {
                        pivotAdoConn = pivotCache.ADOConnection;
                        _output.WriteLine("✅ PivotCache.ADOConnection obtained!");

                        // Execute DAX query
                        string daxQuery = $"EVALUATE '{targetTableName}'";
                        _output.WriteLine($"  Executing: {daxQuery}");

                        try
                        {
                            dynamic? rs = pivotAdoConn.Execute(daxQuery);
                            _output.WriteLine("  ✅ Execute succeeded!");

                            // Read results
                            int fieldCount = rs.Fields.Count;
                            _output.WriteLine($"  Fields: {fieldCount}");

                            int rowCount = 0;
                            while (!rs.EOF && rowCount < 3)
                            {
                                _output.WriteLine($"  Row {rowCount + 1} read");
                                rs.MoveNext();
                                rowCount++;
                            }

                            _output.WriteLine("\n🎉 SUCCESS: PivotCache.ADOConnection + DAX works!");

                            try { rs.Close(); } catch { }
                            ComUtilities.Release(ref rs);
                        }
                        catch (COMException ex)
                        {
                            _output.WriteLine($"  ❌ Execute failed: 0x{ex.HResult:X8} - {ex.Message}");
                        }
                    }
                    catch (COMException ex)
                    {
                        _output.WriteLine($"❌ PivotCache.ADOConnection failed: 0x{ex.HResult:X8}");
                        _output.WriteLine($"   Message: {ex.Message}");
                    }
                }
                catch (COMException ex)
                {
                    _output.WriteLine($"❌ PivotCache.Create failed: 0x{ex.HResult:X8}");
                    _output.WriteLine($"   Message: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                _output.WriteLine($"❌ PivotCache approach failed: {ex.GetType().Name}");
                _output.WriteLine($"   Message: {ex.Message}");
            }
            finally
            {
                ComUtilities.Release(ref pivotAdoConn);
                ComUtilities.Release(ref pivotTable);
                ComUtilities.Release(ref pivotCache);
                ComUtilities.Release(ref pivotCaches);
            }

            // Summary
            _output.WriteLine("\n=== SCENARIO 15 SUMMARY ===");
            _output.WriteLine("This test explored alternative approaches for DAX query execution:");
            _output.WriteLine("1. Model.CreateModelWorkbookConnection + xlCmdDAX");
            _output.WriteLine("2. ModelConnection.ADOConnection.Execute");
            _output.WriteLine("3. PivotCache.ADOConnection.Execute");
            _output.WriteLine("Results above indicate which approach (if any) is viable for Issue #356.");
        }
        catch (Exception ex)
        {
            _output.WriteLine($"\n❌ Unexpected exception: {ex.GetType().Name}");
            _output.WriteLine($"Message: {ex.Message}");
            _output.WriteLine($"StackTrace: {ex.StackTrace}");
        }
        finally
        {
            _output.WriteLine("\n--- Cleanup ---");
            ComUtilities.Release(ref adoConnection);
            ComUtilities.Release(ref modelWbConn);
            ComUtilities.Release(ref modelConnection);
            ComUtilities.Release(ref dataModelConnection);
            ComUtilities.Release(ref model);
            ComUtilities.Release(ref query);
            ComUtilities.Release(ref queries);
            ComUtilities.Release(ref connections);
            ComUtilities.Release(ref sheet);

            if (workbook != null)
            {
                try { workbook.Close(false); }
                catch { /* Ignore */ }
                ComUtilities.Release(ref workbook);
            }

            if (excel != null)
            {
                try { excel.Quit(); }
                catch { /* Ignore */ }
                ComUtilities.Release(ref excel);
            }

            try
            {
                if (File.Exists(testFile))
                    File.Delete(testFile);
            }
            catch { /* Ignore */ }
        }

        _output.WriteLine("=== SCENARIO 15 COMPLETE ===\n");
    }

    /// <summary>
    /// Scenario 16: Test creating Excel Tables backed by DAX queries
    ///
    /// This tests whether we can create ListObjects (Excel Tables) that are populated
    /// by DAX EVALUATE queries, enabling Power BI-style DAX tables in worksheets.
    ///
    /// Approaches to test:
    /// 1. Create DAX connection via Model.CreateModelWorkbookConnection, then create ListObject
    /// 2. Create QueryTable pointing to DAX connection
    /// 3. Use TableObject from an existing ListObject and change to DAX
    ///
    /// Related to GitHub Issue #356
    /// </summary>
    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Layer", "Diagnostics")]
    [Trait("RequiresExcel", "true")]
    [Trait("Feature", "DataModel")]
    [Trait("RunType", "OnDemand")]
    public void Scenario16_DaxBackedExcelTable()
    {
        _output.WriteLine("=== SCENARIO 16: DAX-Backed Excel Tables ===");
        _output.WriteLine("Testing whether ListObjects can be created from DAX EVALUATE queries");
        _output.WriteLine("https://github.com/sbroenne/mcp-server-excel/issues/356\n");

        // Test file with unique name
        string testFile = Path.Combine(_tempDir, $"DMDiag_Scenario16_{Guid.NewGuid():N}.xlsx");
        _output.WriteLine($"Test file: {testFile}");

        dynamic? excel = null;
        dynamic? workbook = null;
        dynamic? sheet = null;
        dynamic? connections = null;
        dynamic? queries = null;
        dynamic? query = null;
        dynamic? model = null;
        dynamic? modelWbConn = null;
        dynamic? queryTable = null;
        dynamic? listObject = null;

        try
        {
            // Constants
            const int xlCmdDAX = 8;
            const int xlCmdTable = 3;
            const int xlYes = 1;

            // Step 1: Create Excel instance
            _output.WriteLine("--- Step 1: Create Excel instance ---");
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            excel = Activator.CreateInstance(excelType!);
            excel.Visible = false;
            excel.DisplayAlerts = false;
            _output.WriteLine("Excel instance created (hidden mode)");

            // Step 2: Create workbook with sample data in Data Model
            _output.WriteLine("\n--- Step 2: Create workbook ---");
            dynamic? workbooks = excel.Workbooks;
            workbook = workbooks.Add();
            ComUtilities.Release(ref workbooks);
            workbook.SaveAs(testFile);
            _output.WriteLine($"Workbook saved: {testFile}");

            sheet = workbook.ActiveSheet;
            string sheetName = sheet.Name;
            _output.WriteLine($"Active sheet: {sheetName}");

            connections = workbook.Connections;
            queries = workbook.Queries;

            // Step 3: Create Power Query with sales data and load to Data Model
            _output.WriteLine("\n--- Step 3: Create Power Query 'SalesData' and load to Data Model ---");
            string mCode = @"let
    Source = #table(
        type table [Region = text, Product = text, Amount = number, Quantity = Int64.Type],
        {
            {""North"", ""Widget"", 100.00, 10},
            {""North"", ""Gadget"", 200.00, 5},
            {""South"", ""Widget"", 150.00, 15},
            {""South"", ""Gadget"", 300.00, 8},
            {""East"", ""Widget"", 80.00, 8},
            {""East"", ""Gadget"", 250.00, 6}
        }
    )
in
    Source";

            query = queries.Add("SalesData", mCode);
            _output.WriteLine("Power Query 'SalesData' created");

            // Load to Data Model via connection
            string connString = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=SalesData;Extended Properties=\"\"";
            dynamic? conn = null;
            try
            {
                conn = connections.Add2(
                    "Query - SalesData",
                    "Power Query - SalesData",
                    connString,
                    "SELECT * FROM [SalesData]",
                    2,      // xlCmdSql
                    true,   // CreateModelConnection = true (load to Data Model)
                    false
                );
                _output.WriteLine("Connection added with CreateModelConnection=true");
                conn.Refresh();
                _output.WriteLine("Connection refreshed - data loaded to Data Model");
                ComUtilities.Release(ref conn);
            }
            catch (COMException ex)
            {
                _output.WriteLine($"Failed to create model connection: 0x{ex.HResult:X8} - {ex.Message}");
                ComUtilities.Release(ref conn);
                throw;
            }

            // Verify Data Model
            model = workbook.Model;
            dynamic? modelTables = model.ModelTables;
            int tableCount = modelTables.Count;
            _output.WriteLine($"Model tables count: {tableCount}");

            string? modelTableName = null;
            for (int i = 1; i <= tableCount; i++)
            {
                dynamic? table = modelTables.Item(i);
                string tableName = table.Name;
                int rowCount = table.RecordCount;
                _output.WriteLine($"  Table {i}: '{tableName}' ({rowCount} rows)");
                modelTableName = tableName;
                ComUtilities.Release(ref table);
            }
            ComUtilities.Release(ref modelTables);

            if (modelTableName == null)
            {
                _output.WriteLine("❌ No tables in Data Model - cannot proceed");
                return;
            }

            // Step 4: Create DAX connection via Model.CreateModelWorkbookConnection
            _output.WriteLine("\n--- Step 4: Create DAX-backed connection ---");
            _output.WriteLine($"Creating model workbook connection for table: {modelTableName}");

            modelWbConn = model.CreateModelWorkbookConnection(modelTableName);
            string daxConnName = modelWbConn.Name;
            _output.WriteLine($"✅ Connection created: {daxConnName}");

            // Get ModelConnection and change to DAX
            dynamic? modelConnection = modelWbConn.ModelConnection;
            _output.WriteLine($"Current CommandType: {Convert.ToInt32(modelConnection.CommandType)}");

            // Set to xlCmdDAX with a DAX query
            string daxQuery = $"EVALUATE SUMMARIZECOLUMNS('{modelTableName}'[Region], \"TotalAmount\", SUM('{modelTableName}'[Amount]), \"TotalQty\", SUM('{modelTableName}'[Quantity]))";
            _output.WriteLine($"\nSetting CommandType to xlCmdDAX (8)");
            _output.WriteLine($"DAX Query: {daxQuery}");

            try
            {
                modelConnection.CommandType = xlCmdDAX;
                modelConnection.CommandText = daxQuery;
                _output.WriteLine("✅ CommandType and CommandText set successfully!");

                // Refresh to validate
                modelWbConn.Refresh();
                _output.WriteLine("✅ Connection refresh with DAX succeeded!");
            }
            catch (COMException ex)
            {
                _output.WriteLine($"❌ Failed to set DAX: 0x{ex.HResult:X8} - {ex.Message}");
                // Reset to table mode
                try
                {
                    modelConnection.CommandType = xlCmdTable;
                    modelConnection.CommandText = modelTableName;
                }
                catch { }
            }

            ComUtilities.Release(ref modelConnection);

            // Step 5: Try to create a ListObject from the DAX connection
            _output.WriteLine("\n--- Step 5: Create ListObject from DAX connection ---");

            dynamic? sheet2 = workbook.Worksheets.Add();
            string sheet2Name = sheet2.Name;
            _output.WriteLine($"Created new sheet: {sheet2Name}");

            dynamic? listObjects = sheet2.ListObjects;
            dynamic? destRange = sheet2.Range["A1"];

            // Approach 5a: Try ListObjects.Add with the DAX connection
            _output.WriteLine("\n5a: Trying ListObjects.Add with DAX WorkbookConnection...");
            try
            {
                const int xlSrcModel = 4;
                listObject = listObjects.Add(
                    xlSrcModel,         // SourceType = xlSrcModel
                    modelWbConn,        // Source = our DAX WorkbookConnection
                    true,               // LinkSource
                    xlYes,              // HasHeaders
                    destRange           // Destination
                );
                _output.WriteLine("✅ ListObject created with xlSrcModel!");

                // Refresh to populate data
                _output.WriteLine("Refreshing ListObject to populate data...");
                listObject.Refresh();
                _output.WriteLine("✅ ListObject refreshed!");

                // Read results
                string loName = listObject.Name;
                _output.WriteLine($"ListObject name: {loName}");

                dynamic? headerRange = listObject.HeaderRowRange;
                if (headerRange != null)
                {
                    object[,]? headers = headerRange.Value2 as object[,];
                    if (headers != null)
                    {
                        var headerList = new List<string>();
                        for (int c = 1; c <= headers.GetLength(1); c++)
                        {
                            headerList.Add(headers[1, c]?.ToString() ?? "(null)");
                        }
                        _output.WriteLine($"Headers: {string.Join(", ", headerList)}");
                    }
                    ComUtilities.Release(ref headerRange);
                }

                dynamic? dataRange = listObject.DataBodyRange;
                if (dataRange != null)
                {
                    object[,]? data = dataRange.Value2 as object[,];
                    if (data != null)
                    {
                        int rows = data.GetLength(0);
                        int cols = data.GetLength(1);
                        _output.WriteLine($"Data: {rows} rows x {cols} columns");
                        for (int r = 1; r <= Math.Min(rows, 5); r++)
                        {
                            var rowValues = new List<string>();
                            for (int c = 1; c <= cols; c++)
                            {
                                rowValues.Add(data[r, c]?.ToString() ?? "(null)");
                            }
                            _output.WriteLine($"  Row {r}: {string.Join(", ", rowValues)}");
                        }
                    }
                    ComUtilities.Release(ref dataRange);
                }

                _output.WriteLine("\n🎉 SUCCESS: DAX-backed Excel Table created via ListObjects.Add!");
            }
            catch (COMException ex)
            {
                _output.WriteLine($"❌ ListObjects.Add failed: 0x{ex.HResult:X8} - {ex.Message}");

                // Approach 5b: Try QueryTables.Add instead
                _output.WriteLine("\n5b: Trying QueryTables.Add with connection string...");
                ComUtilities.Release(ref destRange);
                destRange = sheet2.Range["A1"];

                try
                {
                    // Build connection string for the DAX connection
                    string qtConnString = $"OLEDB;Provider=MSOLAP.8;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=;Data Source=$Embedded$";

                    dynamic? queryTables = sheet2.QueryTables;
                    queryTable = queryTables.Add(
                        qtConnString,
                        destRange,
                        daxQuery
                    );
                    _output.WriteLine("QueryTable created!");

                    // Set properties
                    queryTable.CommandType = xlCmdDAX;
                    _output.WriteLine("CommandType set to xlCmdDAX");

                    // Refresh
                    queryTable.Refresh(false);
                    _output.WriteLine("✅ QueryTable refresh succeeded!");

                    // Check for ListObject
                    dynamic? qtListObject = queryTable.ListObject;
                    if (qtListObject != null)
                    {
                        _output.WriteLine($"ListObject created: {qtListObject.Name}");
                        ComUtilities.Release(ref qtListObject);
                    }

                    _output.WriteLine("\n🎉 SUCCESS: DAX-backed Excel Table created via QueryTables.Add!");
                    ComUtilities.Release(ref queryTables);
                }
                catch (COMException ex2)
                {
                    _output.WriteLine($"❌ QueryTables.Add failed: 0x{ex2.HResult:X8} - {ex2.Message}");

                    // Approach 5c: Try using ADO to populate range, then convert to table
                    _output.WriteLine("\n5c: Trying ADO.Execute + Range.CopyFromRecordset...");
                    try
                    {
                        dynamic? dataModelConn = model.DataModelConnection;
                        dynamic? dmModelConn = dataModelConn.ModelConnection;
                        dynamic? adoConn = dmModelConn.ADOConnection;

                        _output.WriteLine($"Executing DAX: {daxQuery}");
                        dynamic? recordset = adoConn.Execute(daxQuery);
                        _output.WriteLine("✅ DAX query executed!");

                        // Get field names for headers
                        int fieldCount = recordset.Fields.Count;
                        _output.WriteLine($"Fields: {fieldCount}");

                        // Write headers
                        for (int f = 0; f < fieldCount; f++)
                        {
                            string fieldName = recordset.Fields.Item(f).Name;
                            sheet2.Cells[1, f + 1].Value2 = fieldName;
                        }

                        // Use CopyFromRecordset to bulk copy data
                        dynamic? dataStart = sheet2.Range["A2"];
                        int rowsCopied = dataStart.CopyFromRecordset(recordset);
                        _output.WriteLine($"✅ CopyFromRecordset: {rowsCopied} rows copied!");
                        ComUtilities.Release(ref dataStart);

                        // Convert range to ListObject
                        dynamic? usedRange = sheet2.UsedRange;
                        string usedAddress = usedRange.Address;
                        _output.WriteLine($"Used range: {usedAddress}");

                        dynamic? tableRange = sheet2.Range[usedAddress];
                        dynamic? newListObject = listObjects.Add(
                            1,          // xlSrcRange
                            tableRange,
                            null,
                            xlYes,
                            null
                        );
                        _output.WriteLine($"✅ ListObject created: {newListObject.Name}");

                        // Verify data
                        dynamic? body = newListObject.DataBodyRange;
                        if (body != null)
                        {
                            object[,]? data = body.Value2 as object[,];
                            if (data != null)
                            {
                                _output.WriteLine($"Table has {data.GetLength(0)} data rows");
                                for (int r = 1; r <= Math.Min(3, data.GetLength(0)); r++)
                                {
                                    var vals = new List<string>();
                                    for (int c = 1; c <= data.GetLength(1); c++)
                                    {
                                        vals.Add(data[r, c]?.ToString() ?? "(null)");
                                    }
                                    _output.WriteLine($"  Row {r}: {string.Join(", ", vals)}");
                                }
                            }
                            ComUtilities.Release(ref body);
                        }

                        _output.WriteLine("\n🎉 SUCCESS: Excel Table created from DAX via CopyFromRecordset!");
                        _output.WriteLine("NOTE: This table is NOT auto-linked to DAX - it's a snapshot.");
                        _output.WriteLine("      To refresh, re-run the DAX query and update the table.");

                        ComUtilities.Release(ref newListObject);
                        ComUtilities.Release(ref tableRange);
                        ComUtilities.Release(ref usedRange);

                        try { recordset.Close(); }
                        catch { }
                        ComUtilities.Release(ref recordset);
                        ComUtilities.Release(ref adoConn);
                        ComUtilities.Release(ref dmModelConn);
                        ComUtilities.Release(ref dataModelConn);
                    }
                    catch (COMException ex3)
                    {
                        _output.WriteLine($"❌ CopyFromRecordset failed: 0x{ex3.HResult:X8} - {ex3.Message}");
                    }
                }
            }
            finally
            {
                ComUtilities.Release(ref destRange);
            }

            // Summary
            _output.WriteLine("\n=== SCENARIO 16 SUMMARY ===");
            _output.WriteLine("This test explored creating Excel Tables backed by DAX queries.");
            _output.WriteLine("Approaches tested:");
            _output.WriteLine("  5a: ListObjects.Add(xlSrcModel, daxConnection)");
            _output.WriteLine("  5b: QueryTables.Add with MSOLAP connection string");
            _output.WriteLine("  5c: ADO.Execute + CopyFromRecordset + Convert to Table");
            _output.WriteLine("\nResults above indicate which approach (if any) is viable.");

            ComUtilities.Release(ref listObjects);
            ComUtilities.Release(ref sheet2);
        }
        catch (Exception ex)
        {
            _output.WriteLine($"\n❌ Unexpected exception: {ex.GetType().Name}");
            _output.WriteLine($"Message: {ex.Message}");
            _output.WriteLine($"StackTrace: {ex.StackTrace}");
        }
        finally
        {
            _output.WriteLine("\n--- Cleanup ---");
            ComUtilities.Release(ref listObject);
            ComUtilities.Release(ref queryTable);
            ComUtilities.Release(ref modelWbConn);
            ComUtilities.Release(ref model);
            ComUtilities.Release(ref query);
            ComUtilities.Release(ref queries);
            ComUtilities.Release(ref connections);
            ComUtilities.Release(ref sheet);

            if (workbook != null)
            {
                try { workbook.Close(false); }
                catch { /* Ignore */ }
                ComUtilities.Release(ref workbook);
            }

            if (excel != null)
            {
                try { excel.Quit(); }
                catch { /* Ignore */ }
                ComUtilities.Release(ref excel);
            }

            try
            {
                if (File.Exists(testFile))
                    File.Delete(testFile);
            }
            catch { /* Ignore */ }
        }

        _output.WriteLine("=== SCENARIO 16 COMPLETE ===\n");
    }
}
