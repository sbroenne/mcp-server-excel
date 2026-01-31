// =============================================================================
// DIAGNOSTIC TESTS - Direct Excel COM API Behavior
// =============================================================================
// Purpose: Understand what Excel COM API actually does, without our abstractions
// These tests document the REAL behavior of Excel's Power Query COM API
// =============================================================================

using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Diagnostics.Tests.Integration.Diagnostics;

/// <summary>
/// Diagnostic tests for Power Query COM API behavior.
/// These tests use raw COM calls to understand Excel's actual behavior.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "Diagnostics")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
[Trait("RunType", "OnDemand")]
public class PowerQueryComApiBehaviorTests : IClassFixture<TempDirectoryFixture>, IDisposable
{
    private readonly string _tempDir;
    private readonly ITestOutputHelper _output;
    private dynamic? _excel;
    private dynamic? _workbook;
    private readonly string _testFile;

    // Simple M code that creates inline data
    private const string SimpleQuery = """
        let
            Source = #table({"Name", "Value"}, {{"A", 1}, {"B", 2}, {"C", 3}})
        in
            Source
        """;

    private const string ModifiedQuery = """
        let
            Source = #table({"Name", "Value", "Extra"}, {{"A", 1, "X"}, {"B", 2, "Y"}, {"C", 3, "Z"}})
        in
            Source
        """;

    private const string ColumnRemovedQuery = """
        let
            Source = #table({"Name"}, {{"A"}, {"B"}, {"C"}})
        in
            Source
        """;

    public PowerQueryComApiBehaviorTests(TempDirectoryFixture fixture, ITestOutputHelper output)
    {
        _tempDir = fixture.TempDir;
        _output = output;
        _testFile = Path.Combine(_tempDir, $"PQDiag_{Guid.NewGuid():N}.xlsx");

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
    // SCENARIO 1: Basic Query Creation - Load to Table
    // =========================================================================

    [Fact]
    public void Scenario1_CreateQuery_LoadToTable()
    {
        _output.WriteLine("=== SCENARIO 1: Create Query → Load to Table ===");

        // Step 1: Add query to Queries collection
        dynamic? queries = null;
        dynamic? query = null;
        dynamic? sheets = null;
        dynamic? sheet = null;
        dynamic? listObjects = null;

        try
        {
            queries = _workbook.Queries;
            int initialCount = queries.Count;
            _output.WriteLine($"Initial query count: {initialCount}");

            // Add query - this creates the query definition only
            query = queries.Add("TestQuery", SimpleQuery);
            _output.WriteLine($"Query added. Name: {query.Name}");
            _output.WriteLine($"Query count after add: {queries.Count}");

            // Check: Does adding a query automatically create a table? NO
            sheets = _workbook.Worksheets;
            sheet = sheets.Item(1);
            listObjects = sheet.ListObjects;
            _output.WriteLine($"ListObjects count after query add: {listObjects.Count}");

            Assert.Equal(initialCount + 1, (int)queries.Count);
            Assert.Equal(0, (int)listObjects.Count); // Query alone doesn't create table

            // Step 2: To load to table, we need to create a QueryTable
            _output.WriteLine("\n--- Creating QueryTable to load data ---");

            dynamic? range = null;
            dynamic? queryTables = null;
            dynamic? qt = null;

            try
            {
                range = sheet.Range["A1"];
                queryTables = sheet.QueryTables;

                string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=TestQuery";
                qt = queryTables.Add(connectionString, range);
                qt.CommandType = 2; // xlCmdSql
                qt.CommandText = "SELECT * FROM [TestQuery]";

                _output.WriteLine("QueryTable created. Refreshing...");
                qt.Refresh(false); // false = synchronous

                _output.WriteLine($"QueryTable refreshed. RowNumbers: {qt.ResultRange?.Rows?.Count}");

                // Check ListObjects now
                int listObjectCount = listObjects.Count;
                _output.WriteLine($"ListObjects count after refresh: {listObjectCount}");

                // Document behavior
                if (listObjectCount > 0)
                {
                    dynamic? lo = listObjects.Item(1);
                    _output.WriteLine($"ListObject name: {lo?.Name}");
                    ComUtilities.Release(ref lo);
                }
            }
            finally
            {
                ComUtilities.Release(ref qt);
                ComUtilities.Release(ref queryTables);
                ComUtilities.Release(ref range);
            }
        }
        finally
        {
            ComUtilities.Release(ref listObjects);
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref sheets);
            ComUtilities.Release(ref query);
            ComUtilities.Release(ref queries);
        }

        _output.WriteLine("=== SCENARIO 1 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 2: Update Query (Add Column)
    // =========================================================================

    [Fact]
    public void Scenario2_UpdateQuery_AddColumn()
    {
        _output.WriteLine("=== SCENARIO 2: Update Query → Add Column ===");

        dynamic? queries = null;
        dynamic? query = null;

        try
        {
            queries = _workbook.Queries;

            // Create and load query first
            query = queries.Add("UpdateTest", SimpleQuery);
            LoadQueryToTable("UpdateTest", "A1");

            int colsBefore = GetFirstTableColumnCount();
            _output.WriteLine($"Original columns: {colsBefore}");
            _output.WriteLine($"Original formula length: {((string)query.Formula).Length}");

            // Update the formula - NetOffice shows this is just a property set
            _output.WriteLine("\n--- Updating query formula ---");
            query.Formula = ModifiedQuery;

            _output.WriteLine($"New formula length: {((string)query.Formula).Length}");
            _output.WriteLine($"Formula updated successfully: {query.Formula.Contains("Extra")}");

            // Key question: Does the table automatically update? NO - need refresh
            _output.WriteLine("\n--- Checking if table auto-updates (it shouldn't) ---");
            int colsAfterUpdate = GetFirstTableColumnCount();
            _output.WriteLine($"Columns after formula update (before refresh): {colsAfterUpdate}");
            _output.WriteLine($"Table auto-updated? {colsAfterUpdate != colsBefore}");

            // Refresh to see new column
            _output.WriteLine("\n--- Refreshing table ---");
            RefreshFirstTable();

            int colsAfterRefresh = GetFirstTableColumnCount();
            _output.WriteLine($"Columns after refresh: {colsAfterRefresh}");
            _output.WriteLine($"New column appeared? {colsAfterRefresh > colsBefore}");
        }
        finally
        {
            ComUtilities.Release(ref query);
            ComUtilities.Release(ref queries);
        }

        _output.WriteLine("=== SCENARIO 2 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 3: Update Query (Remove Column)
    // =========================================================================

    [Fact]
    public void Scenario3_UpdateQuery_RemoveColumn()
    {
        _output.WriteLine("=== SCENARIO 3: Update Query → Remove Column ===");

        dynamic? queries = null;
        dynamic? query = null;

        try
        {
            queries = _workbook.Queries;
            query = queries.Add("RemoveColTest", SimpleQuery);
            LoadQueryToTable("RemoveColTest", "A1");

            int colsBefore = GetFirstTableColumnCount();
            _output.WriteLine($"Original columns: {colsBefore}");

            // Remove a column
            _output.WriteLine("\n--- Updating query to remove column ---");
            query.Formula = ColumnRemovedQuery;
            _output.WriteLine("Updated to 1 column (Name only)");

            // Refresh to apply schema change
            _output.WriteLine("\n--- Refreshing table ---");
            RefreshFirstTable();

            int colsAfterRefresh = GetFirstTableColumnCount();
            _output.WriteLine($"Columns after refresh: {colsAfterRefresh}");
            _output.WriteLine($"Column removed? {colsAfterRefresh < colsBefore}");
        }
        finally
        {
            ComUtilities.Release(ref query);
            ComUtilities.Release(ref queries);
        }

        _output.WriteLine("=== SCENARIO 3 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 4: Delete Query - What Happens to Table?
    // =========================================================================

    [Fact]
    public void Scenario4_DeleteQuery_TableBehavior()
    {
        _output.WriteLine("=== SCENARIO 4: Delete Query → What Happens to Table? ===");

        dynamic? queries = null;
        dynamic? query = null;
        dynamic? sheet = null;
        dynamic? listObjects = null;

        try
        {
            queries = _workbook.Queries;
            query = queries.Add("DeleteTest", SimpleQuery);
            LoadQueryToTable("DeleteTest", "A1");

            sheet = _workbook.Worksheets.Item(1);
            listObjects = sheet.ListObjects;

            int tableCountBefore = listObjects.Count;
            _output.WriteLine($"Tables before delete: {tableCountBefore}");

            // Get table name before delete
            string? tableName = null;
            if (tableCountBefore > 0)
            {
                dynamic? lo = listObjects.Item(1);
                tableName = lo.Name;
                _output.WriteLine($"Table name: {tableName}");
                ComUtilities.Release(ref lo);
            }

            // DELETE THE QUERY - Key test!
            _output.WriteLine("\n--- Deleting query ---");
            query.Delete();
            ComUtilities.Release(ref query);
            query = null;

            _output.WriteLine($"Query count after delete: {queries.Count}");

            // KEY QUESTION: What happened to the table?
            ComUtilities.Release(ref listObjects);
            listObjects = sheet.ListObjects;
            int tableCountAfter = listObjects.Count;
            _output.WriteLine($"Tables after delete: {tableCountAfter}");

            if (tableCountAfter > 0)
            {
                _output.WriteLine("TABLE SURVIVES! Query deletion does NOT delete the table.");
                dynamic? lo = listObjects.Item(1);
                _output.WriteLine($"Orphaned table name: {lo.Name}");

                // Can we still access the data?
                dynamic? dataRange = lo.DataBodyRange;
                if (dataRange != null)
                {
                    _output.WriteLine($"Data rows: {dataRange.Rows.Count}");
                    ComUtilities.Release(ref dataRange);
                }
                ComUtilities.Release(ref lo);
            }
            else
            {
                _output.WriteLine("TABLE DELETED! Query deletion removes the table too.");
            }
        }
        finally
        {
            ComUtilities.Release(ref listObjects);
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref query);
            ComUtilities.Release(ref queries);
        }

        _output.WriteLine("=== SCENARIO 4 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 5: Re-create Query After Delete
    // =========================================================================

    [Fact]
    public void Scenario5_RecreateQueryAfterDelete()
    {
        _output.WriteLine("=== SCENARIO 5: Re-create Query After Delete ===");

        dynamic? queries = null;
        dynamic? query = null;

        try
        {
            queries = _workbook.Queries;

            // Create, load, delete
            query = queries.Add("RecreateTest", SimpleQuery);
            LoadQueryToTable("RecreateTest", "A1");
            query.Delete();
            ComUtilities.Release(ref query);
            query = null;

            _output.WriteLine("Query deleted. Attempting to recreate with same name...");

            // Can we recreate with same name?
            try
            {
                query = queries.Add("RecreateTest", ModifiedQuery);
                _output.WriteLine($"SUCCESS: Query recreated. Name: {query.Name}");

                // Can we load it again?
                LoadQueryToTable("RecreateTest", "E1"); // Different location
                _output.WriteLine("Query loaded to new location successfully");
            }
            catch (COMException ex)
            {
                _output.WriteLine($"FAILED to recreate: {ex.Message}");
            }
        }
        finally
        {
            ComUtilities.Release(ref query);
            ComUtilities.Release(ref queries);
        }

        _output.WriteLine("=== SCENARIO 5 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 6: Change to Connection Only
    // =========================================================================

    [Fact]
    public void Scenario6_ChangeToConnectionOnly()
    {
        _output.WriteLine("=== SCENARIO 6: Change Query to Connection Only ===");

        dynamic? queries = null;
        dynamic? query = null;
        dynamic? sheet = null;
        dynamic? listObjects = null;
        dynamic? connections = null;

        try
        {
            queries = _workbook.Queries;
            query = queries.Add("ConnOnlyTest", SimpleQuery);

            // First load to table
            LoadQueryToTable("ConnOnlyTest", "A1");

            sheet = _workbook.Worksheets.Item(1);
            listObjects = sheet.ListObjects;
            _output.WriteLine($"Tables after initial load: {listObjects.Count}");

            // Now how do we change to "connection only"?
            // In Excel UI, this means the query exists but doesn't load anywhere
            _output.WriteLine("\n--- Attempting to change to connection only ---");

            // Option 1: Delete the ListObject but keep the query
            if (listObjects.Count > 0)
            {
                dynamic? lo = listObjects.Item(1);
                string loName = lo.Name;
                _output.WriteLine($"Deleting ListObject: {loName}");

                // Unlist converts table to range
                lo.Unlist();
                ComUtilities.Release(ref lo);

                _output.WriteLine("ListObject deleted (Unlist called)");
            }

            // Verify query still exists
            ComUtilities.Release(ref listObjects);
            listObjects = sheet.ListObjects;
            _output.WriteLine($"Tables after Unlist: {listObjects.Count}");
            _output.WriteLine($"Queries count: {queries.Count}");

            // Check connection status
            connections = _workbook.Connections;
            _output.WriteLine($"Connections count: {connections.Count}");

            for (int i = 1; i <= connections.Count; i++)
            {
                dynamic? conn = connections.Item(i);
                _output.WriteLine($"Connection {i}: {conn.Name}, Type: {conn.Type}");
                ComUtilities.Release(ref conn);
            }
        }
        finally
        {
            ComUtilities.Release(ref connections);
            ComUtilities.Release(ref listObjects);
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref query);
            ComUtilities.Release(ref queries);
        }

        _output.WriteLine("=== SCENARIO 6 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 7: Query Error Handling
    // =========================================================================

    [Fact]
    public void Scenario7_QueryWithError()
    {
        _output.WriteLine("=== SCENARIO 7: Query With Error ===");

        const string errorQuery = """
            let
                Source = NonExistentFunction()
            in
                Source
            """;

        dynamic? queries = null;
        dynamic? query = null;

        try
        {
            queries = _workbook.Queries;

            // Can we add a query with invalid M code?
            _output.WriteLine("Adding query with invalid M code...");
            query = queries.Add("ErrorQuery", errorQuery);
            _output.WriteLine($"Query added successfully (no validation on add)");

            // Error should occur on refresh
            _output.WriteLine("\n--- Attempting to load (should fail) ---");
            try
            {
                LoadQueryToTable("ErrorQuery", "A1");
                _output.WriteLine("UNEXPECTED: Query loaded without error");
            }
            catch (COMException ex)
            {
                _output.WriteLine($"EXPECTED ERROR on refresh: 0x{ex.HResult:X8}");
                _output.WriteLine($"Message: {ex.Message}");
            }

            // Query should still exist despite error
            _output.WriteLine($"\nQueries count after error: {queries.Count}");
        }
        finally
        {
            ComUtilities.Release(ref query);
            ComUtilities.Release(ref queries);
        }

        _output.WriteLine("=== SCENARIO 7 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 8: Refresh Behavior
    // =========================================================================

    [Fact]
    public void Scenario8_RefreshBehavior()
    {
        _output.WriteLine("=== SCENARIO 8: Refresh Behavior ===");

        dynamic? queries = null;
        dynamic? query = null;
        dynamic? connections = null;

        try
        {
            queries = _workbook.Queries;
            query = queries.Add("RefreshTest", SimpleQuery);
            LoadQueryToTable("RefreshTest", "A1");

            connections = _workbook.Connections;
            _output.WriteLine($"Connections: {connections.Count}");

            // Find the connection for this query
            dynamic? pqConnection = null;
            for (int i = 1; i <= connections.Count; i++)
            {
                dynamic? conn = connections.Item(i);
                string connName = conn.Name;
                if (connName.Contains("RefreshTest"))
                {
                    pqConnection = conn;
                    _output.WriteLine($"Found connection: {connName}");
                    break;
                }
                ComUtilities.Release(ref conn);
            }

            if (pqConnection != null)
            {
                // Test different refresh methods
                _output.WriteLine("\n--- Testing connection.Refresh() ---");
                try
                {
                    pqConnection.Refresh();
                    _output.WriteLine("connection.Refresh() succeeded");
                }
                catch (COMException ex)
                {
                    _output.WriteLine($"connection.Refresh() failed: {ex.Message}");
                }

                ComUtilities.Release(ref pqConnection);
            }

            // Also test RefreshAll
            _output.WriteLine("\n--- Testing workbook.RefreshAll() ---");
            try
            {
                _workbook.RefreshAll();
                _output.WriteLine("RefreshAll() called (async - may not complete immediately)");
            }
            catch (COMException ex)
            {
                _output.WriteLine($"RefreshAll() failed: {ex.Message}");
            }
        }
        finally
        {
            ComUtilities.Release(ref connections);
            ComUtilities.Release(ref query);
            ComUtilities.Release(ref queries);
        }

        _output.WriteLine("=== SCENARIO 8 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 9: Multiple Queries
    // =========================================================================

    [Fact]
    public void Scenario9_MultipleQueries()
    {
        _output.WriteLine("=== SCENARIO 9: Multiple Queries ===");

        dynamic? queries = null;
        dynamic? query1 = null;
        dynamic? query2 = null;
        dynamic? query3 = null;

        try
        {
            queries = _workbook.Queries;

            // Create multiple queries
            query1 = queries.Add("Query1", SimpleQuery);
            query2 = queries.Add("Query2", ModifiedQuery);
            query3 = queries.Add("Query3", ColumnRemovedQuery);

            _output.WriteLine($"Created 3 queries. Total: {queries.Count}");

            // Load each to different locations
            LoadQueryToTable("Query1", "A1");
            LoadQueryToTable("Query2", "E1");
            LoadQueryToTable("Query3", "I1");

            _output.WriteLine("All queries loaded to tables");

            // Delete middle query
            _output.WriteLine("\n--- Deleting Query2 ---");
            query2.Delete();
            ComUtilities.Release(ref query2);
            query2 = null;

            _output.WriteLine($"Queries after delete: {queries.Count}");

            // Verify other queries still work
            query1.Formula = ModifiedQuery; // Update Query1
            RefreshFirstTable(); // Refresh via table instead of connection name
            _output.WriteLine("Query1 still works after Query2 deletion");
        }
        finally
        {
            ComUtilities.Release(ref query3);
            ComUtilities.Release(ref query2);
            ComUtilities.Release(ref query1);
            ComUtilities.Release(ref queries);
        }

        _output.WriteLine("=== SCENARIO 9 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 10: Special Characters in Query Name
    // =========================================================================

    [Fact]
    public void Scenario10_SpecialCharactersInName()
    {
        _output.WriteLine("=== SCENARIO 10: Special Characters in Query Name ===");

        dynamic? queries = null;

        try
        {
            queries = _workbook.Queries;

            var testNames = new[]
            {
                "Query With Spaces",
                "Query-With-Dashes",
                "Query_With_Underscores",
                "Query.With.Dots",   // Expected to fail - dots not allowed
                "Query123Numbers",
                // "Query/Slash", // Likely invalid
                // "Query:Colon", // Likely invalid
            };

            foreach (var name in testNames)
            {
                try
                {
                    dynamic? q = queries.Add(name, SimpleQuery);
                    _output.WriteLine($"✓ Created: '{name}'");
                    ComUtilities.Release(ref q);
                }
                catch (Exception ex) when (ex is COMException || ex is ArgumentException)
                {
                    _output.WriteLine($"✗ Failed: '{name}' - {ex.Message}");
                }
            }

            _output.WriteLine($"\nTotal queries created: {queries.Count}");
        }
        finally
        {
            ComUtilities.Release(ref queries);
        }

        _output.WriteLine("=== SCENARIO 10 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 11: Load to Data Model
    // =========================================================================

    [Fact]
    public void Scenario11_LoadToDataModel()
    {
        _output.WriteLine("=== SCENARIO 11: Load to Data Model ===");

        dynamic? queries = null;
        dynamic? query = null;
        dynamic? connections = null;
        dynamic? model = null;

        try
        {
            queries = _workbook.Queries;
            query = queries.Add("DataModelTest", SimpleQuery);

            // To load to Data Model, we need to use a different approach
            // The connection needs CreateModelConnection = true
            _output.WriteLine("Query created. Attempting to load to Data Model...");

            connections = _workbook.Connections;

            // Check if a Power Query connection was auto-created
            _output.WriteLine($"Connections after query add: {connections.Count}");

            // Try to access the Data Model
            try
            {
                model = _workbook.Model;
                dynamic? modelTables = model.ModelTables;
                _output.WriteLine($"Model tables before load: {modelTables.Count}");

                // To load to Data Model, we typically need to:
                // 1. Create connection with CreateModelConnection = true
                // 2. Or use the UI's "Load To..." option

                // Let's try creating a connection that loads to model
                string connString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=DataModelTest";

                // Add connection with model flag
                try
                {
                    dynamic? newConn = connections.Add2(
                        "Query - DataModelTest",           // Name
                        "Power Query - DataModelTest",     // Description
                        connString,                        // ConnectionString
                        "SELECT * FROM [DataModelTest]",   // CommandText
                        2,                                 // lCmdtype (xlCmdSql)
                        true,                              // CreateModelConnection - KEY!
                        false                              // ImportRelationships
                    );
                    _output.WriteLine("Connection with CreateModelConnection=true created");

                    // Refresh to load data
                    newConn.Refresh();
                    _output.WriteLine("Connection refreshed");

                    ComUtilities.Release(ref newConn);
                }
                catch (COMException ex)
                {
                    _output.WriteLine($"Add2 with model flag failed: 0x{ex.HResult:X8} - {ex.Message}");
                }

                // Check model tables after
                ComUtilities.Release(ref modelTables);
                modelTables = model.ModelTables;
                _output.WriteLine($"Model tables after load attempt: {modelTables.Count}");

                ComUtilities.Release(ref modelTables);
            }
            catch (COMException ex)
            {
                _output.WriteLine($"Data Model access failed: {ex.Message}");
            }
        }
        finally
        {
            ComUtilities.Release(ref model);
            ComUtilities.Release(ref connections);
            ComUtilities.Release(ref query);
            ComUtilities.Release(ref queries);
        }

        _output.WriteLine("=== SCENARIO 11 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 12: Unload Data Model Query (Bug Discovery Test)
    // =========================================================================

    [Fact]
    public void Scenario12_UnloadDataModelQuery_ConnectionNotRemoved()
    {
        _output.WriteLine("=== SCENARIO 12: Unload Query Loaded to Data Model Only ===");
        _output.WriteLine("PURPOSE: Verify if Unload (removing worksheet tables) handles Data Model connections\n");

        dynamic? queries = null;
        dynamic? query = null;
        dynamic? connections = null;
        dynamic? model = null;

        try
        {
            queries = _workbook.Queries;
            query = queries.Add("DataModelUnloadTest", SimpleQuery);
            _output.WriteLine("Query created: DataModelUnloadTest");

            connections = _workbook.Connections;
            int connectionsBefore = connections.Count;

            // Load to Data Model ONLY (no worksheet table)
            _output.WriteLine("\n--- Loading to Data Model only (no worksheet table) ---");
            string connString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=DataModelUnloadTest";

            try
            {
                dynamic? modelConn = connections.Add2(
                    "Query - DataModelUnloadTest",
                    "Power Query - DataModelUnloadTest",
                    connString,
                    "SELECT * FROM [DataModelUnloadTest]",
                    2,     // xlCmdSql
                    true,  // CreateModelConnection = TRUE (Data Model only)
                    false  // ImportRelationships
                );
                modelConn.Refresh();
                _output.WriteLine("Data Model connection created and refreshed");
                ComUtilities.Release(ref modelConn);
            }
            catch (COMException ex)
            {
                _output.WriteLine($"Failed to create Data Model connection: {ex.Message}");
                return;
            }

            // Verify Data Model has the table
            model = _workbook.Model;
            dynamic? modelTables = model.ModelTables;
            _output.WriteLine($"Model tables after load: {modelTables.Count}");
            ComUtilities.Release(ref modelTables);

            // Verify no ListObjects (worksheet tables)
            dynamic? sheet = _workbook.Worksheets.Item(1);
            dynamic? listObjects = sheet.ListObjects;
            _output.WriteLine($"Worksheet tables (ListObjects): {listObjects.Count}");
            ComUtilities.Release(ref listObjects);
            ComUtilities.Release(ref sheet);

            // Now simulate what our Unload method does: iterate ListObjects and delete them
            _output.WriteLine("\n--- Simulating Unload (only checks ListObjects) ---");
            sheet = _workbook.Worksheets.Item(1);
            listObjects = sheet.ListObjects;
            int tablesToDelete = 0;

            for (int i = listObjects.Count; i >= 1; i--)
            {
                tablesToDelete++;
                // Our Unload only looks at ListObjects
            }
            _output.WriteLine($"ListObjects found to unlist: {tablesToDelete}");
            _output.WriteLine("BUG: Unload only checks ListObjects - it IGNORES Data Model connections!");

            ComUtilities.Release(ref listObjects);
            ComUtilities.Release(ref sheet);

            // Check if Data Model connection still exists
            _output.WriteLine("\n--- Checking Data Model state after 'Unload' ---");
            ComUtilities.Release(ref connections);
            connections = _workbook.Connections;
            _output.WriteLine($"Connections count: {connections.Count}");

            bool modelConnectionExists = false;
            for (int i = 1; i <= connections.Count; i++)
            {
                dynamic? conn = connections.Item(i);
                string connName = conn.Name;
                if (connName.Contains("DataModelUnloadTest"))
                {
                    modelConnectionExists = true;
                    _output.WriteLine($"FINDING: Data Model connection STILL EXISTS: {connName}");
                }
                ComUtilities.Release(ref conn);
            }

            modelTables = model.ModelTables;
            _output.WriteLine($"Model tables after 'Unload': {modelTables.Count}");
            ComUtilities.Release(ref modelTables);

            // Document the bug
            _output.WriteLine("\n=== BUG CONFIRMATION ===");
            _output.WriteLine("Our Unload method only iterates through worksheet ListObjects.");
            _output.WriteLine("For queries loaded ONLY to Data Model, there are NO ListObjects to unlist.");
            _output.WriteLine($"Data Model connection still exists: {modelConnectionExists}");
            _output.WriteLine("Query is NOT connection-only - it's still loaded to Data Model!");

            Assert.True(modelConnectionExists, "Test proves Data Model connection survives Unload");
        }
        finally
        {
            ComUtilities.Release(ref model);
            ComUtilities.Release(ref connections);
            ComUtilities.Release(ref query);
            ComUtilities.Release(ref queries);
        }

        _output.WriteLine("=== SCENARIO 12 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 13: Unload Query Loaded to Both (Worksheet AND Data Model)
    // =========================================================================

    [Fact]
    public void Scenario13_UnloadBothDestinations_OnlyTableRemoved()
    {
        _output.WriteLine("=== SCENARIO 13: Unload Query Loaded to BOTH Worksheet AND Data Model ===");
        _output.WriteLine("PURPOSE: Verify Unload behavior when query is loaded to both destinations\n");

        dynamic? queries = null;
        dynamic? query = null;
        dynamic? connections = null;
        dynamic? model = null;
        dynamic? sheet = null;
        dynamic? listObjects = null;

        try
        {
            queries = _workbook.Queries;
            query = queries.Add("BothDestTest", SimpleQuery);
            _output.WriteLine("Query created: BothDestTest");

            // Step 1: Load to worksheet first
            _output.WriteLine("\n--- Step 1: Load to worksheet ---");
            LoadQueryToTable("BothDestTest", "A1");

            sheet = _workbook.Worksheets.Item(1);
            listObjects = sheet.ListObjects;
            _output.WriteLine($"Worksheet tables after load: {listObjects.Count}");
            ComUtilities.Release(ref listObjects);
            ComUtilities.Release(ref sheet);

            // Step 2: Also load to Data Model
            _output.WriteLine("\n--- Step 2: Also load to Data Model ---");
            connections = _workbook.Connections;
            string connString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=BothDestTest";

            try
            {
                dynamic? modelConn = connections.Add2(
                    "Query - BothDestTest - Model",  // Different name to avoid conflict
                    "Power Query Model Connection",
                    connString,
                    "SELECT * FROM [BothDestTest]",
                    2,     // xlCmdSql
                    true,  // CreateModelConnection = TRUE
                    false
                );
                modelConn.Refresh();
                _output.WriteLine("Data Model connection created");
                ComUtilities.Release(ref modelConn);
            }
            catch (COMException ex)
            {
                _output.WriteLine($"Failed to create Data Model connection: {ex.Message}");
            }

            model = _workbook.Model;
            dynamic? modelTables = model.ModelTables;
            _output.WriteLine($"Model tables: {modelTables.Count}");
            ComUtilities.Release(ref modelTables);

            // Step 3: Simulate Unload - remove worksheet table
            _output.WriteLine("\n--- Step 3: Simulating Unload (removing worksheet table) ---");
            sheet = _workbook.Worksheets.Item(1);
            listObjects = sheet.ListObjects;

            if (listObjects.Count > 0)
            {
                dynamic? lo = listObjects.Item(1);
                string tableName = lo.Name;
                _output.WriteLine($"Unlisting table: {tableName}");
                lo.Unlist();
                ComUtilities.Release(ref lo);
            }

            ComUtilities.Release(ref listObjects);
            listObjects = sheet.ListObjects;
            _output.WriteLine($"Worksheet tables after Unlist: {listObjects.Count}");
            ComUtilities.Release(ref listObjects);
            ComUtilities.Release(ref sheet);

            // Step 4: Check Data Model state
            _output.WriteLine("\n--- Step 4: Check Data Model state after Unload ---");
            ComUtilities.Release(ref connections);
            connections = _workbook.Connections;

            int modelConnectionCount = 0;
            for (int i = 1; i <= connections.Count; i++)
            {
                dynamic? conn = connections.Item(i);
                string connName = conn.Name;
                if (connName.Contains("BothDestTest"))
                {
                    modelConnectionCount++;
                    _output.WriteLine($"Connection still exists: {connName}");
                }
                ComUtilities.Release(ref conn);
            }

            modelTables = model.ModelTables;
            int modelTableCount = modelTables.Count;
            _output.WriteLine($"Model tables after Unload: {modelTableCount}");
            ComUtilities.Release(ref modelTables);

            // Document findings
            _output.WriteLine("\n=== FINDINGS ===");
            _output.WriteLine($"Worksheet table removed: {listObjects?.Count == 0}");
            _output.WriteLine($"Data Model connections remaining: {modelConnectionCount}");
            _output.WriteLine($"Model tables remaining: {modelTableCount}");

            if (modelConnectionCount > 0 || modelTableCount > 0)
            {
                _output.WriteLine("\nBUG: Unload only removes worksheet table, NOT Data Model connection!");
                _output.WriteLine("Query is NOT fully connection-only after Unload.");
            }

            Assert.True(modelConnectionCount > 0 || modelTableCount > 0,
                "Test proves Data Model content survives Unload");
        }
        finally
        {
            ComUtilities.Release(ref model);
            ComUtilities.Release(ref listObjects);
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref connections);
            ComUtilities.Release(ref query);
            ComUtilities.Release(ref queries);
        }

        _output.WriteLine("=== SCENARIO 13 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 14: Proper Connection-Only Implementation
    // =========================================================================

    [Fact]
    public void Scenario14_ProperConnectionOnlyImplementation()
    {
        _output.WriteLine("=== SCENARIO 14: How to Properly Make a Query Connection-Only ===");
        _output.WriteLine("PURPOSE: Document the CORRECT way to make a query connection-only\n");

        dynamic? queries = null;
        dynamic? query = null;
        dynamic? connections = null;
        dynamic? model = null;
#pragma warning disable IDE0059 // Unnecessary assignment - required for COM object lifecycle management
        dynamic? sheet = null;
        dynamic? listObjects = null;
#pragma warning restore IDE0059

        try
        {
            queries = _workbook.Queries;
            query = queries.Add("FullUnloadTest", SimpleQuery);

            // Load to BOTH destinations
            _output.WriteLine("--- Loading to both worksheet AND Data Model ---");
            LoadQueryToTable("FullUnloadTest", "A1");

            connections = _workbook.Connections;
            string connString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=FullUnloadTest";

            try
            {
                dynamic? modelConn = connections.Add2(
                    "Query - FullUnloadTest - Model",
                    "Power Query Model Connection",
                    connString,
                    "SELECT * FROM [FullUnloadTest]",
                    2, true, false
                );
                modelConn.Refresh();
                ComUtilities.Release(ref modelConn);
            }
            catch (COMException ex)
            {
                _output.WriteLine($"Model connection creation failed: {ex.Message}");
            }

            sheet = _workbook.Worksheets.Item(1);
            listObjects = sheet.ListObjects;
            model = _workbook.Model;
            dynamic? modelTables = model.ModelTables;

            _output.WriteLine($"Initial state - Worksheet tables: {listObjects.Count}, Model tables: {modelTables.Count}");
            ComUtilities.Release(ref modelTables);
            ComUtilities.Release(ref listObjects);
            ComUtilities.Release(ref sheet);

            // PROPER Unload: Remove BOTH worksheet tables AND Data Model connections
            _output.WriteLine("\n--- PROPER Unload Implementation ---");

            // Step 1: Remove worksheet tables (ListObjects)
            _output.WriteLine("Step 1: Remove worksheet tables");
            sheet = _workbook.Worksheets.Item(1);
            listObjects = sheet.ListObjects;
            for (int i = listObjects.Count; i >= 1; i--)
            {
                dynamic? lo = listObjects.Item(i);
                dynamic? qt = lo.QueryTable;
                if (qt != null)
                {
                    string? connName = null;
                    try { connName = qt.Connection?.ToString(); } catch (COMException) { /* Connection property may not exist */ }
                    if (connName?.Contains("FullUnloadTest") == true)
                    {
                        _output.WriteLine($"  Unlisting table: {lo.Name}");
                        lo.Unlist();
                    }
                    ComUtilities.Release(ref qt);
                }
                ComUtilities.Release(ref lo);
            }
            ComUtilities.Release(ref listObjects);
            ComUtilities.Release(ref sheet);

            // Step 2: Remove Data Model connections (the missing step in our Unload!)
            _output.WriteLine("Step 2: Remove Data Model connections");
            ComUtilities.Release(ref connections);
            connections = _workbook.Connections;

            // Find and delete connections for this query
            var connectionsToDelete = new List<string>();
            for (int i = 1; i <= connections.Count; i++)
            {
                dynamic? conn = connections.Item(i);
                string connName = conn.Name;
                if (connName.Contains("FullUnloadTest"))
                {
                    connectionsToDelete.Add(connName);
                }
                ComUtilities.Release(ref conn);
            }

            foreach (var connName in connectionsToDelete)
            {
                try
                {
                    dynamic? conn = connections.Item(connName);
                    _output.WriteLine($"  Deleting connection: {connName}");
                    conn.Delete();
                    ComUtilities.Release(ref conn);
                }
                catch (Exception ex)
                {
                    _output.WriteLine($"  Failed to delete {connName}: {ex.Message}");
                }
            }

            // Verify final state
            _output.WriteLine("\n--- Final State (should be connection-only) ---");
            sheet = _workbook.Worksheets.Item(1);
            listObjects = sheet.ListObjects;
            _output.WriteLine($"Worksheet tables: {listObjects.Count}");

            ComUtilities.Release(ref connections);
            connections = _workbook.Connections;
            int remainingPQConnections = 0;
            for (int i = 1; i <= connections.Count; i++)
            {
                dynamic? conn = connections.Item(i);
                if (((string)conn.Name).Contains("FullUnloadTest"))
                {
                    remainingPQConnections++;
                    _output.WriteLine($"  Remaining connection: {conn.Name}");
                }
                ComUtilities.Release(ref conn);
            }
            _output.WriteLine($"Power Query connections for this query: {remainingPQConnections}");

            modelTables = model.ModelTables;
            _output.WriteLine($"Model tables: {modelTables.Count}");

            // Verify query still exists
            _output.WriteLine($"\nQuery still exists: {queries.Count > 0}");
            _output.WriteLine($"Query name: {query.Name}");

            bool isConnectionOnly = listObjects.Count == 0 && remainingPQConnections == 0;
            _output.WriteLine($"\nIS CONNECTION-ONLY: {isConnectionOnly}");

            ComUtilities.Release(ref modelTables);
            ComUtilities.Release(ref listObjects);
            ComUtilities.Release(ref sheet);

            _output.WriteLine("\n=== IMPLEMENTATION REQUIREMENT ===");
            _output.WriteLine("To make a query connection-only, Unload must:");
            _output.WriteLine("1. Remove worksheet tables (ListObjects with matching QueryTable)");
            _output.WriteLine("2. Remove Data Model connections (connections with 'Query - {name}' pattern)");
            _output.WriteLine("3. Keep the query in Workbook.Queries collection");
        }
        finally
        {
            ComUtilities.Release(ref model);
            ComUtilities.Release(ref connections);
            ComUtilities.Release(ref query);
            ComUtilities.Release(ref queries);
        }

        _output.WriteLine("=== SCENARIO 14 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 15: Update Data Model-Only Query
    // =========================================================================

    [Fact]
    public void Scenario15_UpdateDataModelOnlyQuery()
    {
        _output.WriteLine("=== SCENARIO 15: Update Query Loaded ONLY to Data Model ===");
        _output.WriteLine("PURPOSE: Test if Update works for Data Model-only queries (no worksheet table)\n");

        dynamic? queries = null;
        dynamic? query = null;
        dynamic? connections = null;
        dynamic? model = null;

        try
        {
            queries = _workbook.Queries;
            query = queries.Add("DataModelUpdateTest", SimpleQuery);
            _output.WriteLine("Query created: DataModelUpdateTest");

            // Load to Data Model ONLY (no worksheet table)
            _output.WriteLine("\n--- Loading to Data Model only ---");
            connections = _workbook.Connections;
            string connString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=DataModelUpdateTest";

            dynamic? modelConn = null;
            try
            {
                modelConn = connections.Add2(
                    "Query - DataModelUpdateTest",
                    "Power Query - DataModelUpdateTest",
                    connString,
                    "SELECT * FROM [DataModelUpdateTest]",
                    2,     // xlCmdSql
                    true,  // CreateModelConnection = TRUE (Data Model only)
                    false  // ImportRelationships
                );
                modelConn.Refresh();
                _output.WriteLine("Data Model connection created and refreshed");
            }
            catch (COMException ex)
            {
                _output.WriteLine($"Failed to create Data Model connection: {ex.Message}");
                return;
            }

            // Verify initial state - Data Model has data
            model = _workbook.Model;
            dynamic? modelTables = model.ModelTables;
            int initialModelTableCount = modelTables.Count;
            _output.WriteLine($"Model tables after initial load: {initialModelTableCount}");
            ComUtilities.Release(ref modelTables);

            // Verify NO worksheet tables
            dynamic? sheet = _workbook.Worksheets.Item(1);
            dynamic? listObjects = sheet.ListObjects;
            _output.WriteLine($"Worksheet tables (should be 0): {listObjects.Count}");
            Assert.Equal(0, (int)listObjects.Count);
            ComUtilities.Release(ref listObjects);
            ComUtilities.Release(ref sheet);

            // Get original formula
            string originalFormula = query.Formula;
            _output.WriteLine($"\nOriginal formula contains 'Extra': {originalFormula.Contains("Extra")}");

            // =========================================================================
            // TEST 1: Update M code (adds a column)
            // =========================================================================
            _output.WriteLine("\n--- TEST 1: Update M code (add 'Extra' column) ---");
            query.Formula = ModifiedQuery;
            string newFormula = query.Formula;
            _output.WriteLine($"Formula updated. Contains 'Extra': {newFormula.Contains("Extra")}");
            Assert.True(newFormula.Contains("Extra"), "Formula should be updated");

            // =========================================================================
            // TEST 2: What happens WITHOUT refresh?
            // =========================================================================
            _output.WriteLine("\n--- TEST 2: Check Data Model WITHOUT refresh ---");
            // At this point, M code is updated but we haven't refreshed
            // Question: Is the Data Model stale?

            // We can't easily inspect Data Model column structure via COM,
            // but we can check if the connection still exists
            bool connectionExists = false;
            ComUtilities.Release(ref connections);
            connections = _workbook.Connections;
            for (int i = 1; i <= connections.Count; i++)
            {
                dynamic? conn = connections.Item(i);
                if (((string)conn.Name).Contains("DataModelUpdateTest"))
                {
                    connectionExists = true;
                    _output.WriteLine($"Connection exists: {conn.Name}");
                }
                ComUtilities.Release(ref conn);
            }
            _output.WriteLine($"Connection still exists: {connectionExists}");

            // =========================================================================
            // TEST 3: Refresh via Connection.Refresh()
            // =========================================================================
            _output.WriteLine("\n--- TEST 3: Refresh via connection.Refresh() ---");
            try
            {
                // Find and refresh the connection
                bool refreshed = false;
                for (int i = 1; i <= connections.Count; i++)
                {
                    dynamic? conn = connections.Item(i);
                    string connName = conn.Name;
                    if (connName.Contains("DataModelUpdateTest"))
                    {
                        _output.WriteLine($"Refreshing connection: {connName}");
                        conn.Refresh();
                        refreshed = true;
                        _output.WriteLine("connection.Refresh() succeeded");
                    }
                    ComUtilities.Release(ref conn);
                    if (refreshed) break;
                }

                if (!refreshed)
                {
                    _output.WriteLine("WARNING: No connection found to refresh!");
                }
            }
            catch (COMException ex)
            {
                _output.WriteLine($"connection.Refresh() FAILED: 0x{ex.HResult:X8} - {ex.Message}");
            }

            // =========================================================================
            // TEST 4: Verify Data Model state after refresh
            // =========================================================================
            _output.WriteLine("\n--- TEST 4: Data Model state after refresh ---");
            modelTables = model.ModelTables;
            _output.WriteLine($"Model tables after refresh: {modelTables.Count}");

            // List ALL model tables and their columns to verify Extra column appeared
            bool foundExtraColumn = false;
            for (int i = 1; i <= modelTables.Count; i++)
            {
                dynamic? mt = null;
                dynamic? cols = null;
                try
                {
                    mt = modelTables.Item(i);
                    string tableName = mt.Name;
                    _output.WriteLine($"Model table {i}: '{tableName}'");

                    cols = mt.ModelTableColumns;
                    _output.WriteLine($"  Column count: {cols.Count}");

                    // List ALL column names
                    for (int c = 1; c <= cols.Count; c++)
                    {
                        dynamic? col = cols.Item(c);
                        string colName = col.Name?.ToString() ?? "(null)";
                        _output.WriteLine($"    Column {c}: {colName}");
                        if (colName == "Extra")
                        {
                            foundExtraColumn = true;
                            _output.WriteLine($"    ^^^ FOUND 'Extra' column! connection.Refresh() WORKS! ^^^");
                        }
                        ComUtilities.Release(ref col);
                    }
                }
                finally
                {
                    ComUtilities.Release(ref cols);
                    ComUtilities.Release(ref mt);
                }
            }
            ComUtilities.Release(ref modelTables);

            // =========================================================================
            // FINDINGS
            // =========================================================================
            _output.WriteLine("\n=== FINDINGS ===");
            _output.WriteLine("1. M code update via query.Formula = works for Data Model queries");
            _output.WriteLine("2. Without refresh, Data Model has STALE data");
            _output.WriteLine($"3. connection.Refresh() propagates column changes: {(foundExtraColumn ? "YES - Extra column found!" : "NO - Extra column NOT found")}");
            _output.WriteLine("4. Our Update.cs currently does NOT refresh Data Model-only queries");
            _output.WriteLine("   (because it only looks for QueryTables on worksheets)");

            // ASSERT: Extra column should exist after connection.Refresh()
            Assert.True(foundExtraColumn, "Extra column should appear in Data Model after connection.Refresh()");

            ComUtilities.Release(ref modelConn);
        }
        finally
        {
            ComUtilities.Release(ref model);
            ComUtilities.Release(ref connections);
            ComUtilities.Release(ref query);
            ComUtilities.Release(ref queries);
        }

        _output.WriteLine("=== SCENARIO 15 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 16: Rename Query
    // =========================================================================

    [Fact]
    public void Scenario16_RenameQuery()
    {
        _output.WriteLine("=== SCENARIO 12: Rename Query ===");

        dynamic? queries = null;
        dynamic? query = null;
        dynamic? sheet = null;
        dynamic? listObjects = null;

        try
        {
            queries = _workbook.Queries;
            query = queries.Add("OriginalName", SimpleQuery);
            LoadQueryToTable("OriginalName", "A1");

            sheet = _workbook.Worksheets.Item(1);
            listObjects = sheet.ListObjects;

            string? tableNameBefore = null;
            if (listObjects.Count > 0)
            {
                dynamic? lo = listObjects.Item(1);
                tableNameBefore = lo.Name;
                _output.WriteLine($"Table name before rename: {tableNameBefore}");
                ComUtilities.Release(ref lo);
            }

            // Rename the query
            _output.WriteLine("\n--- Renaming query to 'NewName' ---");
            query.Name = "NewName";
            _output.WriteLine($"Query renamed. New name: {query.Name}");

            // Check if table name changed
            ComUtilities.Release(ref listObjects);
            listObjects = sheet.ListObjects;
            if (listObjects.Count > 0)
            {
                dynamic? lo = listObjects.Item(1);
                string tableNameAfter = lo.Name;
                _output.WriteLine($"Table name after rename: {tableNameAfter}");

                if (tableNameBefore == tableNameAfter)
                {
                    _output.WriteLine("TABLE NAME DID NOT CHANGE when query was renamed");
                }
                else
                {
                    _output.WriteLine("TABLE NAME CHANGED when query was renamed");
                }
                ComUtilities.Release(ref lo);
            }

            // Can we still refresh via the table?
            _output.WriteLine("\n--- Refreshing table after query rename ---");
            try
            {
                RefreshFirstTable();
                _output.WriteLine("Refresh succeeded after query rename");
            }
            catch (Exception ex)
            {
                _output.WriteLine($"Refresh failed: {ex.Message}");
            }
        }
        finally
        {
            ComUtilities.Release(ref listObjects);
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref query);
            ComUtilities.Release(ref queries);
        }

        _output.WriteLine("=== SCENARIO 12 COMPLETE ===\n");
    }

    // =========================================================================
    // Helper Methods - Direct COM calls
    // =========================================================================

    private void LoadQueryToTable(string queryName, string startCell)
    {
        dynamic? sheet = null;
        dynamic? range = null;
        dynamic? listObjects = null;
        dynamic? listObject = null;
        dynamic? queryTable = null;

        try
        {
            sheet = _workbook.Worksheets.Item(1);
            range = sheet.Range[startCell];
            listObjects = sheet.ListObjects;

            // Use ListObjects.Add with xlSrcExternal (0) - this creates a proper Excel Table
            string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName};Extended Properties=\"\"";
            listObject = listObjects.Add(
                0,                  // SourceType: 0 = xlSrcExternal
                connectionString,   // Source: connection string
                Type.Missing,       // LinkSource
                1,                  // XlListObjectHasHeaders: xlYes
                range               // Destination: starting cell
            );

            // Configure the QueryTable behind the ListObject
            queryTable = listObject.QueryTable;
            queryTable.CommandType = 2; // xlCmdSql
            queryTable.CommandText = $"SELECT * FROM [{queryName}]";
            queryTable.BackgroundQuery = false; // Synchronous
            queryTable.Refresh(false); // Synchronous
        }
        finally
        {
            ComUtilities.Release(ref queryTable);
            ComUtilities.Release(ref listObject);
            ComUtilities.Release(ref listObjects);
            ComUtilities.Release(ref range);
            ComUtilities.Release(ref sheet);
        }
    }

    /// <summary>
    /// Refreshes the first table's QueryTable directly.
    /// This is simpler than finding connections by name since ListObjects.Add creates auto-named connections.
    /// </summary>
    private void RefreshFirstTable()
    {
        dynamic? sheet = null;
        dynamic? listObjects = null;
        dynamic? listObject = null;
        dynamic? queryTable = null;

        try
        {
            sheet = _workbook.Worksheets.Item(1);
            listObjects = sheet.ListObjects;

            if (listObjects.Count == 0)
            {
                throw new InvalidOperationException("No tables found to refresh");
            }

            listObject = listObjects.Item(1);
            queryTable = listObject.QueryTable;
            queryTable.Refresh(false); // Synchronous refresh

            _output.WriteLine("Table QueryTable refreshed successfully");
        }
        finally
        {
            ComUtilities.Release(ref queryTable);
            ComUtilities.Release(ref listObject);
            ComUtilities.Release(ref listObjects);
            ComUtilities.Release(ref sheet);
        }
    }

    // =========================================================================
    // SCENARIO 17: Use Add2 for Worksheet Loading (avoid orphaned connections)
    // =========================================================================

    [Fact]
    public void Scenario17_UseAdd2ForWorksheetLoading()
    {
        _output.WriteLine("=== SCENARIO 17: Use Add2 for Worksheet Loading ===");
        _output.WriteLine("PURPOSE: Test if we can use Connections.Add2 with CreateModelConnection=false");
        _output.WriteLine("         and then create a ListObject that uses that named connection\n");

        dynamic? queries = null;
        dynamic? query = null;
        dynamic? connections = null;
        dynamic? connection = null;
        dynamic? sheet = null;
        dynamic? listObjects = null;
        dynamic? listObject = null;
        dynamic? queryTable = null;
        dynamic? range = null;

        try
        {
            // Step 1: Create query
            queries = _workbook.Queries;
            query = queries.Add("Add2WorksheetTest", SimpleQuery);
            _output.WriteLine("Query created: Add2WorksheetTest");

            // Step 2: Create connection with Add2 (CreateModelConnection = false)
            connections = _workbook.Connections;
            string connString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Add2WorksheetTest";

            _output.WriteLine("\n--- Step 2: Create connection with Add2 (CreateModelConnection=false) ---");
            connection = connections.Add2(
                "Query - Add2WorksheetTest",           // Name - proper naming!
                "Power Query - Add2WorksheetTest",     // Description
                connString,                            // ConnectionString
                "SELECT * FROM [Add2WorksheetTest]",   // CommandText
                2,                                     // lCmdtype (xlCmdSql)
                false,                                 // CreateModelConnection = FALSE (worksheet, not data model)
                false                                  // ImportRelationships
            );
            _output.WriteLine($"Connection created: {connection.Name}");

            // Step 3: Try to create ListObject using this connection
            _output.WriteLine("\n--- Step 3: Create ListObject using the named connection ---");
            sheet = _workbook.Worksheets.Item(1);
            range = sheet.Range["A1"];
            listObjects = sheet.ListObjects;

            // Try using the connection name instead of connection string
            try
            {
                // Method 1: Use connection name directly
                listObject = listObjects.Add(
                    0,                              // SourceType: 0 = xlSrcExternal
                    connection,                     // Source: try passing connection object
                    Type.Missing,                   // LinkSource
                    1,                              // XlListObjectHasHeaders: xlYes
                    range                           // Destination
                );
                _output.WriteLine("ListObjects.Add with connection object SUCCEEDED!");
            }
            catch (Exception ex)
            {
                _output.WriteLine($"ListObjects.Add with connection object FAILED: {ex.Message}");

                // Method 2: Try with connection string but specifying the connection name
                ComUtilities.Release(ref listObject);
                try
                {
                    // Use the full connection string but the connection should already exist
                    string fullConnString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Add2WorksheetTest;Extended Properties=\"\"";
                    listObject = listObjects.Add(
                        0,                          // SourceType: 0 = xlSrcExternal
                        fullConnString,             // Source: connection string
                        Type.Missing,               // LinkSource
                        1,                          // XlListObjectHasHeaders: xlYes
                        range                       // Destination
                    );
                    _output.WriteLine("ListObjects.Add with connection string SUCCEEDED!");
                }
                catch (Exception ex2)
                {
                    _output.WriteLine($"ListObjects.Add with connection string ALSO FAILED: {ex2.Message}");
                }
            }

            if (listObject != null)
            {
                // Configure and refresh
                queryTable = listObject.QueryTable;
                queryTable.CommandType = 2;
                queryTable.CommandText = "SELECT * FROM [Add2WorksheetTest]";
                queryTable.BackgroundQuery = false;

                _output.WriteLine("\n--- Step 4: Refresh and check connection name ---");
                queryTable.Refresh(false);
                _output.WriteLine("Refresh succeeded!");

                // Check how many connections exist now
                ComUtilities.Release(ref connections);
                connections = _workbook.Connections;
                _output.WriteLine($"\nConnections after ListObjects.Add:");
                for (int i = 1; i <= connections.Count; i++)
                {
                    dynamic? conn = connections.Item(i);
                    _output.WriteLine($"  Connection {i}: '{conn.Name}' (Type: {conn.Type})");
                    ComUtilities.Release(ref conn);
                }

                // Key question: Did ListObjects.Add create ANOTHER connection or use our existing one?
                int connectionCount = connections.Count;
                if (connectionCount == 1)
                {
                    _output.WriteLine("\nSUCCESS: Only 1 connection exists - ListObjects.Add used our named connection!");
                }
                else
                {
                    _output.WriteLine($"\nISSUE: {connectionCount} connections exist - ListObjects.Add created a new one!");
                }
            }
        }
        catch (Exception ex)
        {
            _output.WriteLine($"ERROR: {ex.Message}");
        }
        finally
        {
            ComUtilities.Release(ref queryTable);
            ComUtilities.Release(ref listObject);
            ComUtilities.Release(ref listObjects);
            ComUtilities.Release(ref range);
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref connection);
            ComUtilities.Release(ref connections);
            ComUtilities.Release(ref query);
            ComUtilities.Release(ref queries);
        }

        _output.WriteLine("=== SCENARIO 17 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 18: List vs View Behavior on Connection-Only Queries
    // =========================================================================
    // Bug Report: List works but View fails with 0x800A03EC on complex queries
    // This test investigates what operations fail on connection-only queries

    [Fact]
    public void Scenario18_ListVsView_ConnectionOnlyQuery()
    {
        _output.WriteLine("=== SCENARIO 18: List vs View on Connection-Only Query ===");
        _output.WriteLine("PURPOSE: Investigate why List works but View fails with 0x800A03EC\n");

        // Use a more complex M code that references functions (similar to bug report)
        const string complexQuery = """
            let
                // Simulates a query with helper function (like fnLoadMilestoneExport in bug report)
                fnHelper = (x as number) => x * 2,
                Source = #table({"ID", "Name", "Value"}, {
                    {1, "Item1", fnHelper(100)},
                    {2, "Item2", fnHelper(200)},
                    {3, "Item3", fnHelper(300)}
                }),
                AddColumn = Table.AddColumn(Source, "Doubled", each fnHelper([Value]))
            in
                AddColumn
            """;

        dynamic? queries = null;
        dynamic? query = null;
        dynamic? worksheets = null;

        try
        {
            queries = _workbook.Queries;

            // STEP 1: Create a connection-only query (no loading to worksheet)
            _output.WriteLine("--- STEP 1: Create Connection-Only Query ---");
            query = queries.Add("ComplexConnectionOnly", complexQuery);
            _output.WriteLine($"Query created: ComplexConnectionOnly");
            _output.WriteLine($"Query count: {queries.Count}");

            // STEP 2: Test List-like operations (what List() does)
            _output.WriteLine("\n--- STEP 2: Test List-like Operations ---");
            _output.WriteLine("Iterating queries like List() does...\n");

            for (int i = 1; i <= queries.Count; i++)
            {
                dynamic? q = null;
                try
                {
                    _output.WriteLine($"Query {i}:");

                    // Test accessing queries.Item(i) - like List does
                    _output.WriteLine($"  Calling queries.Item({i})...");
                    q = queries.Item(i);
                    _output.WriteLine($"  SUCCESS: queries.Item({i}) worked");

                    // Test accessing Name property
                    _output.WriteLine($"  Calling q.Name...");
                    string name = q.Name?.ToString() ?? "(null)";
                    _output.WriteLine($"  SUCCESS: Name = '{name}'");

                    // Test accessing Formula property (this is what List catches with try-catch)
                    _output.WriteLine($"  Calling q.Formula...");
                    try
                    {
                        string formula = q.Formula?.ToString() ?? "(null)";
                        _output.WriteLine($"  SUCCESS: Formula length = {formula.Length} chars");
                        _output.WriteLine($"  Formula preview: {formula[..Math.Min(50, formula.Length)]}...");
                    }
                    catch (COMException ex)
                    {
                        _output.WriteLine($"  FAILED: Formula access threw 0x{ex.HResult:X8}");
                        _output.WriteLine($"  Message: {ex.Message}");
                    }
                }
                catch (COMException ex)
                {
                    _output.WriteLine($"  FAILED at query {i}: 0x{ex.HResult:X8} - {ex.Message}");
                }
                finally
                {
                    if (q != null) ComUtilities.Release(ref q!);
                }
            }

            // STEP 3: Test View-like operations (what View() does differently)
            _output.WriteLine("\n--- STEP 3: Test View-like Operations ---");
            _output.WriteLine("Simulating View() operations...\n");

            // View does the same query lookup, but then also iterates worksheets
            dynamic? foundQuery = null;
            try
            {
                // Find query by name (same as View)
                _output.WriteLine("Finding query by name 'ComplexConnectionOnly'...");
                for (int i = 1; i <= queries.Count; i++)
                {
                    dynamic? q = null;
                    try
                    {
                        q = queries.Item(i);
                        string qName = q.Name?.ToString() ?? "";
                        if (qName.Equals("ComplexConnectionOnly", StringComparison.OrdinalIgnoreCase))
                        {
                            foundQuery = q;
                            q = null; // Don't release
                            _output.WriteLine($"  Found query at index {i}");
                            break;
                        }
                    }
                    finally
                    {
                        if (q != null) ComUtilities.Release(ref q!);
                    }
                }

                if (foundQuery == null)
                {
                    _output.WriteLine("  ERROR: Query not found!");
                }
                else
                {
                    // Read Formula (same as View)
                    _output.WriteLine("\nReading Formula property...");
                    try
                    {
                        string mCode = foundQuery.Formula?.ToString() ?? "";
                        _output.WriteLine($"  SUCCESS: Formula length = {mCode.Length}");
                    }
                    catch (COMException ex)
                    {
                        _output.WriteLine($"  FAILED: 0x{ex.HResult:X8} - {ex.Message}");
                    }

                    // Now iterate worksheets to detect load configuration (this is what View does extra)
                    _output.WriteLine("\nIterating worksheets to detect load configuration...");
                    worksheets = _workbook.Worksheets;
                    _output.WriteLine($"  Worksheet count: {worksheets.Count}");

                    for (int ws = 1; ws <= worksheets.Count; ws++)
                    {
                        dynamic? worksheet = null;
                        dynamic? queryTables = null;
                        dynamic? listObjects = null;

                        try
                        {
                            _output.WriteLine($"\n  Worksheet {ws}:");
                            worksheet = worksheets.Item(ws);
                            _output.WriteLine($"    Name: {worksheet.Name}");

                            // Check QueryTables
                            _output.WriteLine($"    Accessing QueryTables...");
                            try
                            {
                                queryTables = worksheet.QueryTables;
                                _output.WriteLine($"    SUCCESS: QueryTables.Count = {queryTables.Count}");

                                for (int qt = 1; qt <= queryTables.Count; qt++)
                                {
                                    dynamic? qTable = null;
                                    dynamic? wbConn = null;
                                    dynamic? oledbConn = null;
                                    try
                                    {
                                        _output.WriteLine($"      QueryTable {qt}:");
                                        qTable = queryTables.Item(qt);

                                        _output.WriteLine($"        Accessing WorkbookConnection...");
                                        wbConn = qTable.WorkbookConnection;
                                        if (wbConn == null)
                                        {
                                            _output.WriteLine($"        WorkbookConnection is null");
                                            continue;
                                        }
                                        _output.WriteLine($"        SUCCESS: WorkbookConnection accessed");

                                        _output.WriteLine($"        Accessing OLEDBConnection...");
                                        oledbConn = wbConn.OLEDBConnection;
                                        if (oledbConn == null)
                                        {
                                            _output.WriteLine($"        OLEDBConnection is null");
                                            continue;
                                        }
                                        _output.WriteLine($"        SUCCESS: OLEDBConnection accessed");

                                        _output.WriteLine($"        Accessing Connection string...");
                                        string connString = oledbConn.Connection?.ToString() ?? "";
                                        _output.WriteLine($"        SUCCESS: Connection string length = {connString.Length}");
                                    }
                                    catch (COMException ex)
                                    {
                                        _output.WriteLine($"        FAILED: 0x{ex.HResult:X8} - {ex.Message}");
                                    }
                                    finally
                                    {
                                        if (oledbConn != null) ComUtilities.Release(ref oledbConn!);
                                        if (wbConn != null) ComUtilities.Release(ref wbConn!);
                                        if (qTable != null) ComUtilities.Release(ref qTable!);
                                    }
                                }
                            }
                            catch (COMException ex)
                            {
                                _output.WriteLine($"    FAILED accessing QueryTables: 0x{ex.HResult:X8} - {ex.Message}");
                            }

                            // Check ListObjects
                            _output.WriteLine($"    Accessing ListObjects...");
                            try
                            {
                                listObjects = worksheet.ListObjects;
                                _output.WriteLine($"    SUCCESS: ListObjects.Count = {listObjects.Count}");

                                for (int lo = 1; lo <= listObjects.Count; lo++)
                                {
                                    dynamic? listObj = null;
                                    dynamic? loQueryTable = null;

                                    try
                                    {
                                        _output.WriteLine($"      ListObject {lo}:");
                                        listObj = listObjects.Item(lo);

                                        _output.WriteLine($"        Accessing QueryTable property...");
                                        try
                                        {
                                            loQueryTable = listObj.QueryTable;
                                            if (loQueryTable == null)
                                            {
                                                _output.WriteLine($"        QueryTable is null (manual table)");
                                            }
                                            else
                                            {
                                                _output.WriteLine($"        SUCCESS: QueryTable accessed");
                                            }
                                        }
                                        catch (COMException ex)
                                        {
                                            _output.WriteLine($"        EXPECTED: ListObject.QueryTable threw 0x{ex.HResult:X8}");
                                            _output.WriteLine($"        (Normal for ListObjects without QueryTable)");
                                        }
                                    }
                                    finally
                                    {
                                        if (loQueryTable != null) ComUtilities.Release(ref loQueryTable!);
                                        if (listObj != null) ComUtilities.Release(ref listObj!);
                                    }
                                }
                            }
                            catch (COMException ex)
                            {
                                _output.WriteLine($"    FAILED accessing ListObjects: 0x{ex.HResult:X8} - {ex.Message}");
                            }
                        }
                        catch (COMException ex)
                        {
                            _output.WriteLine($"    FAILED accessing worksheet: 0x{ex.HResult:X8} - {ex.Message}");
                        }
                        finally
                        {
                            if (listObjects != null) ComUtilities.Release(ref listObjects!);
                            if (queryTables != null) ComUtilities.Release(ref queryTables!);
                            if (worksheet != null) ComUtilities.Release(ref worksheet!);
                        }
                    }
                }
            }
            finally
            {
                if (foundQuery != null) ComUtilities.Release(ref foundQuery!);
            }

            _output.WriteLine("\n--- SUMMARY ---");
            _output.WriteLine("If List-like operations succeeded but View-like failed,");
            _output.WriteLine("the issue is in the worksheet/QueryTable/ListObject iteration.");
            _output.WriteLine("If both succeeded, the issue may be specific to the real workbook's state.");
        }
        finally
        {
            ComUtilities.Release(ref worksheets);
            ComUtilities.Release(ref query);
            ComUtilities.Release(ref queries);
        }

        _output.WriteLine("=== SCENARIO 18 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 19: Query with Dependencies (like bug report)
    // =========================================================================
    // Bug Report mentions query referencing another query (fnEnsureColumn)

    [Fact]
    public void Scenario19_QueryWithDependencies()
    {
        _output.WriteLine("=== SCENARIO 19: Query with Dependencies ===");
        _output.WriteLine("PURPOSE: Test View/Update on queries that reference other queries\n");

        // Create a base query first
        const string baseQuery = """
            let
                Source = #table({"ID", "Name"}, {{1, "A"}, {2, "B"}, {3, "C"}})
            in
                Source
            """;

        // Create a dependent query that references the base
        const string dependentQuery = """
            let
                Source = BaseQuery,
                AddValue = Table.AddColumn(Source, "Value", each [ID] * 10)
            in
                AddValue
            """;

        dynamic? queries = null;
        dynamic? baseQ = null;
        dynamic? depQ = null;

        try
        {
            queries = _workbook.Queries;

            // Create base query
            _output.WriteLine("--- Creating Base Query ---");
            baseQ = queries.Add("BaseQuery", baseQuery);
            _output.WriteLine($"Base query created. Count: {queries.Count}");

            // Create dependent query
            _output.WriteLine("\n--- Creating Dependent Query ---");
            depQ = queries.Add("DependentQuery", dependentQuery);
            _output.WriteLine($"Dependent query created. Count: {queries.Count}");

            // Test accessing Formula on both
            _output.WriteLine("\n--- Testing Formula Access ---");

            ComUtilities.Release(ref baseQ);
            ComUtilities.Release(ref depQ);

            for (int i = 1; i <= queries.Count; i++)
            {
                dynamic? q = null;
                try
                {
                    q = queries.Item(i);
                    string name = q.Name?.ToString() ?? "";
                    _output.WriteLine($"\nQuery: {name}");

                    try
                    {
                        string formula = q.Formula?.ToString() ?? "";
                        _output.WriteLine($"  Formula access: SUCCESS ({formula.Length} chars)");
                    }
                    catch (COMException ex)
                    {
                        _output.WriteLine($"  Formula access: FAILED 0x{ex.HResult:X8}");
                    }

                    // Try to update the formula (like Update does)
                    _output.WriteLine($"  Testing Formula assignment...");
                    try
                    {
                        string currentFormula = q.Formula?.ToString() ?? "";
                        // Just reassign the same formula
                        q.Formula = currentFormula;
                        _output.WriteLine($"  Formula assignment: SUCCESS");
                    }
                    catch (COMException ex)
                    {
                        _output.WriteLine($"  Formula assignment: FAILED 0x{ex.HResult:X8}");
                        _output.WriteLine($"  Message: {ex.Message}");
                    }
                }
                finally
                {
                    if (q != null) ComUtilities.Release(ref q!);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref depQ);
            ComUtilities.Release(ref baseQ);
            ComUtilities.Release(ref queries);
        }

        _output.WriteLine("\n=== SCENARIO 19 COMPLETE ===\n");
    }

    // =========================================================================
    // SCENARIO 20: ListObject.QueryTable Access on Non-External Data Tables
    // =========================================================================
    // CRITICAL TEST: Does accessing QueryTable property on a ListObject
    // created from regular data (not external connection) throw an exception?

    [Fact]
    public void Scenario20_ListObjectQueryTableAccess_OnRegularTable()
    {
        _output.WriteLine("=== SCENARIO 20: ListObject.QueryTable Access on Regular Tables ===");
        _output.WriteLine("PURPOSE: Determine if accessing QueryTable on a non-query ListObject throws\n");

        dynamic? sheets = null;
        dynamic? sheet = null;
        dynamic? range = null;
        dynamic? listObjects = null;
        dynamic? listObject = null;

        try
        {
            // Create a regular Excel table (not from external data)
            sheets = _workbook.Worksheets;
            sheet = sheets.Add();
            string sheetName = sheet.Name;

            _output.WriteLine($"Created test sheet: {sheetName}");

            // Add some data
            sheet.Range["A1"].Value2 = "Header1";
            sheet.Range["B1"].Value2 = "Header2";
            sheet.Range["A2"].Value2 = "Value1";
            sheet.Range["B2"].Value2 = "Value2";

            // Create a ListObject from range (regular table, NOT from external data)
            range = sheet.Range["A1:B2"];
            listObjects = sheet.ListObjects;

            // xlSrcRange = 1 (create from range data)
            listObject = listObjects.Add(1, range, Type.Missing, 1, Type.Missing);
            string tableName = listObject.Name;
            _output.WriteLine($"Created regular table: {tableName}");

            // NOW: Try to access QueryTable on this regular table
            _output.WriteLine("\n--- Testing ListObject.QueryTable access ---");
            _output.WriteLine("Attempting to access listObject.QueryTable on regular table...");

            dynamic? queryTable = null;
            try
            {
                queryTable = listObject.QueryTable;
                if (queryTable == null)
                {
                    _output.WriteLine("RESULT: QueryTable property returned NULL (no exception)");
                }
                else
                {
                    _output.WriteLine($"RESULT: QueryTable property returned an object (unexpected for regular table)");
                }
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                _output.WriteLine($"RESULT: COMException thrown!");
                _output.WriteLine($"  HResult: 0x{ex.HResult:X8}");
                _output.WriteLine($"  Message: {ex.Message}");
                _output.WriteLine("\n*** FINDING: View/Update MUST use try-catch when accessing ListObject.QueryTable ***");
            }
            finally
            {
                if (queryTable != null) ComUtilities.Release(ref queryTable!);
            }

            // Clean up the test table
            listObject.Delete();
            sheet.Delete();

            _output.WriteLine("\nCleanup complete");
        }
        finally
        {
            if (listObject != null) ComUtilities.Release(ref listObject!);
            if (listObjects != null) ComUtilities.Release(ref listObjects!);
            if (range != null) ComUtilities.Release(ref range!);
            if (sheet != null) ComUtilities.Release(ref sheet!);
            if (sheets != null) ComUtilities.Release(ref sheets!);
        }

        _output.WriteLine("\n=== SCENARIO 20 COMPLETE ===\n");
    }

    /// <summary>
    /// Gets the column count from the first table.
    /// </summary>
    private int GetFirstTableColumnCount()
    {
        dynamic? sheet = null;
        dynamic? listObjects = null;
        dynamic? listObject = null;
        dynamic? columns = null;

        try
        {
            sheet = _workbook.Worksheets.Item(1);
            listObjects = sheet.ListObjects;

            if (listObjects.Count == 0)
            {
                return 0;
            }

            listObject = listObjects.Item(1);
            columns = listObject.ListColumns;
            return columns.Count;
        }
        finally
        {
            ComUtilities.Release(ref columns);
            ComUtilities.Release(ref listObject);
            ComUtilities.Release(ref listObjects);
            ComUtilities.Release(ref sheet);
        }
    }
}
