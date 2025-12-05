// =============================================================================
// DIAGNOSTIC TESTS - Direct Excel COM API Behavior for Data Model
// =============================================================================
// Purpose: Understand what Excel COM API actually does for Data Model operations
// These tests document the REAL behavior of Excel's Data Model/Power Pivot COM API
// =============================================================================

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
}
