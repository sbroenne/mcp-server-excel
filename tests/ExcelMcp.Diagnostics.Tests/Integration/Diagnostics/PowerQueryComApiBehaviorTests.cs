// =============================================================================
// DIAGNOSTIC TESTS - Direct Excel COM API Behavior
// =============================================================================
// Purpose: Understand what Excel COM API actually does, without our abstractions
// These tests document the REAL behavior of Excel's Power Query COM API
// =============================================================================

using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
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
[Collection("Sequential")]
[Trait("RequiresExcel", "true")]
[Trait("RunType", "OnDemand")]
public class PowerQueryComApiBehaviorTests : IClassFixture<TempDirectoryFixture>, IDisposable
{
    private readonly string _tempDir;
    private readonly ITestOutputHelper _output;
    private Excel.Application? _excel;
    private Excel.Workbook? _workbook;
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

    public PowerQueryComApiBehaviorTests(TempDirectoryFixture fixture, ITestOutputHelper output)
    {
        _tempDir = fixture.TempDir;
        _output = output;
        _testFile = Path.Combine(_tempDir, $"PQDiag_{Guid.NewGuid():N}.xlsx");

        // Create Excel instance directly via COM
        var excelType = Type.GetTypeFromProgID("Excel.Application");
        Assert.NotNull(excelType);
        try
        {
            _excel = (Excel.Application?)Activator.CreateInstance(excelType);
            Assert.NotNull(_excel);
            _excel.Visible = false;
            _excel.DisplayAlerts = false;
            _workbook = DiagnosticExcelObjects.AddWorkbook(_excel);
            _workbook.SaveAs(_testFile);
            _output.WriteLine($"Test file: {_testFile}");
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

    public void Dispose()
    {
        GC.SuppressFinalize(this);
        ExcelShutdownService.CloseAndQuit(_workbook, _excel, save: false);
        _workbook = null;
        _excel = null;
    }

    // =========================================================================
    // SCENARIO 1: Basic Query Creation - Load to Table
    // =========================================================================

    [Fact]
    public void Scenario1_CreateQuery_LoadToTable()
    {
        Assert.NotNull(_workbook);
        _output.WriteLine("=== SCENARIO 1: Create Query → Load to Table ===");

        // Step 1: Add query to Queries collection
        Excel.Queries? queries = null;
        Excel.WorkbookQuery? query = null;
        Excel.Sheets? sheets = null;
        Excel.Worksheet? sheet = null;
        Excel.ListObjects? listObjects = null;

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
            sheet = (Excel.Worksheet)sheets[1];
            listObjects = sheet.ListObjects;
            _output.WriteLine($"ListObjects count after query add: {listObjects.Count}");

            Assert.Equal(initialCount + 1, (int)queries.Count);
            Assert.Equal(0, (int)listObjects.Count); // Query alone doesn't create table

            // Step 2: To load to table, we need to create a QueryTable
            _output.WriteLine("\n--- Creating QueryTable to load data ---");

            Excel.Range? range = null;
            Excel.QueryTables? queryTables = null;
            Excel.QueryTable? qt = null;

            try
            {
                range = sheet.Range["A1"];
                queryTables = sheet.QueryTables;

                string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=TestQuery";
                qt = queryTables.Add(connectionString, range);
                qt.CommandType = Excel.XlCmdType.xlCmdSql; // xlCmdSql
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
                    Excel.ListObject? lo = listObjects[1];
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
    // SCENARIO 15: Update Data Model-Only Query
    // =========================================================================

    [Fact]
    public void Scenario15_UpdateDataModelOnlyQuery()
    {
        Assert.NotNull(_workbook);
        _output.WriteLine("=== SCENARIO 15: Update Query Loaded ONLY to Data Model ===");
        _output.WriteLine("PURPOSE: Test if Update works for Data Model-only queries (no worksheet table)\n");

        Excel.Queries? queries = null;
        Excel.WorkbookQuery? query = null;
        Excel.Connections? connections = null;
        Excel.Model? model = null;

        try
        {
            queries = _workbook.Queries;
            query = queries.Add("DataModelUpdateTest", SimpleQuery);
            _output.WriteLine("Query created: DataModelUpdateTest");

            // Load to Data Model ONLY (no worksheet table)
            _output.WriteLine("\n--- Loading to Data Model only ---");
            connections = _workbook.Connections;
            string connString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=DataModelUpdateTest";

            Excel.WorkbookConnection? modelConn = null;
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
            Excel.ModelTables? modelTables = model.ModelTables;
            int initialModelTableCount = modelTables.Count;
            _output.WriteLine($"Model tables after initial load: {initialModelTableCount}");
            ComUtilities.Release(ref modelTables);

            // Verify NO worksheet tables
            Excel.Worksheet? sheet = DiagnosticExcelObjects.GetWorksheet(_workbook, 1);
            Excel.ListObjects? listObjects = sheet.ListObjects;
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
                Excel.WorkbookConnection? conn = connections.Item(i);
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
                    Excel.WorkbookConnection? conn = connections.Item(i);
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
                Excel.ModelTable? mt = null;
                Excel.ModelTableColumns? cols = null;
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
                        Excel.ModelTableColumn? col = cols.Item(c);
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
    // SCENARIO 17: Use Add2 for Worksheet Loading (avoid orphaned connections)
    // =========================================================================

    [Fact]
    public void Scenario17_UseAdd2ForWorksheetLoading()
    {
        Assert.NotNull(_workbook);
        _output.WriteLine("=== SCENARIO 17: Use Add2 for Worksheet Loading ===");
        _output.WriteLine("PURPOSE: Test if we can use Connections.Add2 with CreateModelConnection=false");
        _output.WriteLine("         and then create a ListObject that uses that named connection\n");

        Excel.Queries? queries = null;
        Excel.WorkbookQuery? query = null;
        Excel.Connections? connections = null;
        Excel.WorkbookConnection? connection = null;
        Excel.Worksheet? sheet = null;
        Excel.ListObjects? listObjects = null;
        Excel.ListObject? listObject = null;
        Excel.QueryTable? queryTable = null;
        Excel.Range? range = null;

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
            sheet = DiagnosticExcelObjects.GetWorksheet(_workbook, 1);
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
                    Excel.XlYesNoGuess.xlYes,                              // XlListObjectHasHeaders: xlYes
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
                        Excel.XlYesNoGuess.xlYes,                          // XlListObjectHasHeaders: xlYes
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
                queryTable.CommandType = Excel.XlCmdType.xlCmdSql;
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
                    Excel.WorkbookConnection? conn = connections.Item(i);
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
        Assert.NotNull(_workbook);
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

        Excel.Queries? queries = null;
        Excel.WorkbookQuery? query = null;
        Excel.Sheets? worksheets = null;

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
                Excel.WorkbookQuery? q = null;
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
            Excel.WorkbookQuery? foundQuery = null;
            try
            {
                // Find query by name (same as View)
                _output.WriteLine("Finding query by name 'ComplexConnectionOnly'...");
                for (int i = 1; i <= queries.Count; i++)
                {
                    Excel.WorkbookQuery? q = null;
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
                        Excel.Worksheet? worksheet = null;
                        Excel.QueryTables? queryTables = null;
                        Excel.ListObjects? listObjects = null;

                        try
                        {
                            _output.WriteLine($"\n  Worksheet {ws}:");
                            worksheet = (Excel.Worksheet)worksheets[ws];
                            _output.WriteLine($"    Name: {worksheet.Name}");

                            // Check QueryTables
                            _output.WriteLine($"    Accessing QueryTables...");
                            try
                            {
                                queryTables = worksheet.QueryTables;
                                _output.WriteLine($"    SUCCESS: QueryTables.Count = {queryTables.Count}");

                                for (int qt = 1; qt <= queryTables.Count; qt++)
                                {
                                    Excel.QueryTable? qTable = null;
                                    Excel.WorkbookConnection? wbConn = null;
                                    Excel.OLEDBConnection? oledbConn = null;
                                    try
                                    {
                                        _output.WriteLine($"      QueryTable {qt}:");
                                        qTable = queryTables[qt];

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
                                        string connString = ((object?)oledbConn.Connection)?.ToString() ?? "";
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
                                    Excel.ListObject? listObj = null;
                                    Excel.QueryTable? loQueryTable = null;

                                    try
                                    {
                                        _output.WriteLine($"      ListObject {lo}:");
                                        listObj = listObjects[lo];

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
    // SCENARIO 20: ListObject.QueryTable Access on Non-External Data Tables
    // =========================================================================
    // CRITICAL TEST: Does accessing QueryTable property on a ListObject
    // created from regular data (not external connection) throw an exception?

    [Fact]
    public void Scenario20_ListObjectQueryTableAccess_OnRegularTable()
    {
        Assert.NotNull(_workbook);
        _output.WriteLine("=== SCENARIO 20: ListObject.QueryTable Access on Regular Tables ===");
        _output.WriteLine("PURPOSE: Determine if accessing QueryTable on a non-query ListObject throws\n");

        Excel.Sheets? sheets = null;
        Excel.Worksheet? sheet = null;
        Excel.Range? range = null;
        Excel.ListObjects? listObjects = null;
        Excel.ListObject? listObject = null;

        try
        {
            // Create a regular Excel table (not from external data)
            sheets = _workbook.Worksheets;
            sheet = sheets.Add();
            string sheetName = sheet.Name;

            _output.WriteLine($"Created test sheet: {sheetName}");

            // Add some data
            range = sheet.Range["A1:B2"];
            range.Value2 = new object[,]
            {
                { "Header1", "Header2" },
                { "Value1", "Value2" }
            };

            // Create a ListObject from range (regular table, NOT from external data)
            listObjects = sheet.ListObjects;

            // xlSrcRange = 1 (create from range data)
            listObject = listObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, range,
                Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing);
            string tableName = listObject.Name;
            _output.WriteLine($"Created regular table: {tableName}");

            // NOW: Try to access QueryTable on this regular table
            _output.WriteLine("\n--- Testing ListObject.QueryTable access ---");
            _output.WriteLine("Attempting to access listObject.QueryTable on regular table...");

            Excel.QueryTable? queryTable = null;
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

}
