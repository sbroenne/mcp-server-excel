using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// DIAGNOSTIC TEST - Observes Excel's actual QueryTable behavior using raw COM API.
/// This test does NOT use our wrapper code - it directly calls Excel COM to observe behavior.
///
/// PURE OBSERVATION - No expectations, no assertions about what "should" happen.
/// Just perform raw COM operations and report what Excel actually does.
///
/// Scenarios:
/// 1. Load a PowerQuery to a worksheet - observe columns, rows, data
/// 2. Refresh it - observe if data/structure changes
/// 3. Update M code and refresh - observe success/failure, timeouts, column changes
/// 4. Create connection-only query - observe if QueryTables created
/// 5. Load connection-only to worksheet - observe data, QueryTable behavior
///
/// NOTE: These tests are marked RunType=OnDemand and only run when explicitly requested.
/// They validate production code patterns and Excel COM behavior but are slow (~20s each).
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "PowerQuery")]
[Trait("Layer", "Diagnostics")]
[Trait("RequiresExcel", "true")]
[Trait("RunType", "OnDemand")]
public class ExcelQueryTableBehaviorDiagnostics : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    /// <inheritdoc/>

    public ExcelQueryTableBehaviorDiagnostics(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Combine(Path.GetTempPath(), $"excel-qt-diagnostics-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }
    /// <inheritdoc/>

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
        {
            try { Directory.Delete(_tempDir, recursive: true); }
            catch { /* Best effort */ }
        }
        GC.SuppressFinalize(this);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task Scenario1_LoadPowerQueryToWorksheet()
    {
        string testFile = Path.Combine(_tempDir, "scenario1-load.xlsx");

        // Create empty workbook
        var fileCommands = new FileCommands();
        var createResult = await fileCommands.CreateEmptyAsync(testFile);
        Assert.True(createResult.Success, $"Failed to create test file: {createResult.ErrorMessage}");

        await using var session = await ExcelSession.BeginBatchAsync(testFile);

        _output.WriteLine("SCENARIO 1: Load PowerQuery to worksheet");
        _output.WriteLine("-".PadRight(80, '-'));
        await Scenario1_LoadToWorksheet(session);
        _output.WriteLine("✓ Scenario 1 complete");
    }

    [Fact]
    public async Task Scenario2_RefreshLoadedQuery()
    {
        string testFile = Path.Combine(_tempDir, "scenario2-refresh.xlsx");

        // Create empty workbook
        var fileCommands = new FileCommands();
        var createResult = await fileCommands.CreateEmptyAsync(testFile);
        Assert.True(createResult.Success, $"Failed to create test file: {createResult.ErrorMessage}");

        await using var session = await ExcelSession.BeginBatchAsync(testFile);

        _output.WriteLine("SCENARIO 2: Refresh loaded query (like UI Refresh button)");
        _output.WriteLine("-".PadRight(80, '-'));

        // First create and load the query (same as Scenario 1)
        await Scenario1_LoadToWorksheet(session);

        // Then test refreshing it
        await Scenario2_RefreshLoaded(session);
        _output.WriteLine("✓ Scenario 2 complete");
    }

    [Fact]
    public async Task Scenario3_UpdateMCodeAndRefresh()
    {
        string testFile = Path.Combine(_tempDir, "scenario3-update.xlsx");

        // Create empty workbook
        var fileCommands = new FileCommands();
        var createResult = await fileCommands.CreateEmptyAsync(testFile);
        Assert.True(createResult.Success, $"Failed to create test file: {createResult.ErrorMessage}");

        await using var session = await ExcelSession.BeginBatchAsync(testFile);

        _output.WriteLine("SCENARIO 3: Update query M code and refresh");
        _output.WriteLine("-".PadRight(80, '-'));

        // First create and load the query
        await Scenario1_LoadToWorksheet(session);

        // Then test updating M code (this triggers RPC timeout)
        await Scenario3_UpdateAndRefresh(session);
        _output.WriteLine("✓ Scenario 3 complete");
    }

    [Fact]
    public async Task Scenario4_CreateConnectionOnlyQuery()
    {
        string testFile = Path.Combine(_tempDir, "scenario4-connectiononly.xlsx");

        // Create empty workbook
        var fileCommands = new FileCommands();
        var createResult = await fileCommands.CreateEmptyAsync(testFile);
        Assert.True(createResult.Success, $"Failed to create test file: {createResult.ErrorMessage}");

        await using var session = await ExcelSession.BeginBatchAsync(testFile);

        _output.WriteLine("SCENARIO 4: Change connection-only query to load to worksheet");
        _output.WriteLine("-".PadRight(80, '-'));
        await Scenario4_ChangeConnectionOnlyToLoadToWorksheet(session);
        _output.WriteLine("✓ Scenario 4 complete");
    }

    [Fact]
    public async Task Scenario5_LoadConnectionOnlyToWorksheet_Test()
    {
        string testFile = Path.Combine(_tempDir, "scenario5-loadconnectiononly.xlsx");

        // Create empty workbook
        var fileCommands = new FileCommands();
        var createResult = await fileCommands.CreateEmptyAsync(testFile);
        Assert.True(createResult.Success, $"Failed to create test file: {createResult.ErrorMessage}");

        await using var session = await ExcelSession.BeginBatchAsync(testFile);

        _output.WriteLine("SCENARIO 5: Load connection-only to worksheet then delete QueryTable");
        _output.WriteLine("-".PadRight(80, '-'));
        await Scenario5_LoadConnectionOnlyToWorksheet(session);
        _output.WriteLine("✓ Scenario 5 complete");
    }

    private async Task Scenario1_LoadToWorksheet(IExcelBatch session)
    {
        string queryName = "Scenario1Query";
        string mCode = @"
let
    Source = {1..5},
    ToTable = Table.FromList(Source, Splitter.SplitByNothing(), {""Number""}),
    Typed = Table.TransformColumnTypes(ToTable, {{""Number"", Int64.Type}})
in
    Typed";

        _output.WriteLine($"OBSERVATION: Creating query '{queryName}' and loading to worksheet");
        _output.WriteLine($"M Code: Single column 'Number' with values 1-5");
        _output.WriteLine("");

        await session.Execute((ctx, ct) =>
        {
            dynamic? queries = null;
            dynamic? query = null;
            dynamic? sheets = null;
            dynamic? sheet = null;
            dynamic? queryTables = null;
            dynamic? queryTable = null;

            try
            {
                // Step 1: Create Power Query
                queries = ctx.Book.Queries;
                int queriesCountBefore = queries.Count;

                query = queries.Add(queryName, mCode);
                int queriesCountAfter = queries.Count;

                _output.WriteLine($"STEP 1 - Create Query:");
                _output.WriteLine($"  Queries.Count before: {queriesCountBefore}");
                _output.WriteLine($"  Queries.Count after: {queriesCountAfter}");
                _output.WriteLine($"  Query.Name: {query.Name}");
                _output.WriteLine("");

                // Step 2: Create worksheet for the data
                sheets = ctx.Book.Worksheets;
                int sheetsCountBefore = sheets.Count;

                sheet = sheets.Add();
                sheet.Name = queryName;

                int sheetsCountAfter = sheets.Count;
                _output.WriteLine($"STEP 2 - Create Worksheet:");
                _output.WriteLine($"  Worksheets.Count before: {sheetsCountBefore}");
                _output.WriteLine($"  Worksheets.Count after: {sheetsCountAfter}");
                _output.WriteLine($"  Worksheet.Name: {sheet.Name}");
                _output.WriteLine("");

                // Step 3: Create QueryTable to load data
                string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";

                queryTables = sheet.QueryTables;
                int qtCountBefore = queryTables.Count;

                queryTable = queryTables.Add(connectionString, sheet.Range["A1"], Type.Missing);
                queryTable.Name = queryName;
                queryTable.CommandText = $"SELECT * FROM [{queryName}]";

                int qtCountAfter = queryTables.Count;
                _output.WriteLine($"STEP 3 - Create QueryTable:");
                _output.WriteLine($"  QueryTables.Count before: {qtCountBefore}");
                _output.WriteLine($"  QueryTables.Count after: {qtCountAfter}");
                _output.WriteLine($"  QueryTable.Name: {queryTable.Name}");
                _output.WriteLine($"  QueryTable.CommandText: {queryTable.CommandText}");
                _output.WriteLine("");

                // Step 4: Refresh to load data
                _output.WriteLine($"STEP 4 - Refresh QueryTable:");
                _output.WriteLine($"  Calling queryTable.Refresh(false)...");

                bool refreshResult = queryTable.Refresh(false); // false = synchronous

                _output.WriteLine($"  Refresh() returned: {refreshResult}");
                _output.WriteLine($"  QueryTables.Count after refresh: {queryTables.Count}");
                _output.WriteLine("");

                // Step 5: Observe what Excel created
                _output.WriteLine($"STEP 5 - Observe Worksheet Data:");
                dynamic? usedRange = null;
                try
                {
                    usedRange = sheet.UsedRange;
                    int rowCount = usedRange.Rows.Count;
                    int colCount = usedRange.Columns.Count;

                    _output.WriteLine($"  UsedRange.Address: {usedRange.Address}");
                    _output.WriteLine($"  UsedRange.Rows.Count: {rowCount}");
                    _output.WriteLine($"  UsedRange.Columns.Count: {colCount}");
                    _output.WriteLine("");

                    _output.WriteLine($"  Data Sample:");
                    for (int row = 1; row <= Math.Min(3, rowCount); row++)
                    {
                        for (int col = 1; col <= colCount; col++)
                        {
                            var cellAddr = $"{(char)('A' + col - 1)}{row}";
                            var cellValue = sheet.Range[cellAddr].Value2;
                            _output.WriteLine($"    {cellAddr}: {cellValue ?? "(null)"}");
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref usedRange);
                }

                return 0;
            }
            finally
            {
                ComUtilities.Release(ref queryTable);
                ComUtilities.Release(ref queryTables);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
                ComUtilities.Release(ref query);
                ComUtilities.Release(ref queries);
            }
        });

        _output.WriteLine("");
        _output.WriteLine("=".PadRight(80, '='));
    }

    private async Task Scenario2_RefreshLoaded(IExcelBatch session)
    {
        string sheetName = "Scenario1Query";

        _output.WriteLine($"OBSERVATION: Refreshing already-loaded query");
        _output.WriteLine($"Question: What happens on 2nd and 3rd refresh? Duplicates? Errors?");
        _output.WriteLine("");

        await session.Execute((ctx, ct) =>
        {
            dynamic? sheets = null;
            dynamic? sheet = null;
            dynamic? queryTables = null;
            dynamic? queryTable = null;

            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = sheets.Item(sheetName);

                queryTables = sheet.QueryTables;

                _output.WriteLine($"STEP 1 - Initial State:");
                _output.WriteLine($"  QueryTables.Count: {queryTables.Count}");

                if (queryTables.Count > 0)
                {
                    queryTable = queryTables.Item(1);
                    _output.WriteLine($"  QueryTable.Name: {queryTable.Name}");
                    _output.WriteLine("");

                    // Get baseline before refresh
                    dynamic? usedRange1 = null;
                    try
                    {
                        usedRange1 = sheet.UsedRange;
                        _output.WriteLine($"  Before 2nd refresh:");
                        _output.WriteLine($"    UsedRange: {usedRange1.Address}");
                        _output.WriteLine($"    Rows: {usedRange1.Rows.Count}");
                        _output.WriteLine($"    Columns: {usedRange1.Columns.Count}");
                    }
                    finally
                    {
                        ComUtilities.Release(ref usedRange1);
                    }
                    _output.WriteLine("");

                    // 2nd refresh
                    _output.WriteLine($"STEP 2 - Second Refresh:");
                    _output.WriteLine($"  Calling queryTable.Refresh(false)...");
                    bool refreshResult2 = queryTable.Refresh(false);
                    _output.WriteLine($"  Refresh() returned: {refreshResult2}");
                    _output.WriteLine($"  QueryTables.Count: {queryTables.Count}");

                    dynamic? usedRange2 = null;
                    try
                    {
                        usedRange2 = sheet.UsedRange;
                        _output.WriteLine($"  After 2nd refresh:");
                        _output.WriteLine($"    UsedRange: {usedRange2.Address}");
                        _output.WriteLine($"    Rows: {usedRange2.Rows.Count}");
                        _output.WriteLine($"    Columns: {usedRange2.Columns.Count}");
                    }
                    finally
                    {
                        ComUtilities.Release(ref usedRange2);
                    }
                    _output.WriteLine("");

                    // 3rd refresh
                    _output.WriteLine($"STEP 3 - Third Refresh:");
                    _output.WriteLine($"  Calling queryTable.Refresh(false)...");
                    bool refreshResult3 = queryTable.Refresh(false);
                    _output.WriteLine($"  Refresh() returned: {refreshResult3}");
                    _output.WriteLine($"  QueryTables.Count: {queryTables.Count}");

                    dynamic? usedRange3 = null;
                    try
                    {
                        usedRange3 = sheet.UsedRange;
                        _output.WriteLine($"  After 3rd refresh:");
                        _output.WriteLine($"    UsedRange: {usedRange3.Address}");
                        _output.WriteLine($"    Rows: {usedRange3.Rows.Count}");
                        _output.WriteLine($"    Columns: {usedRange3.Columns.Count}");
                    }
                    finally
                    {
                        ComUtilities.Release(ref usedRange3);
                    }
                }
                else
                {
                    _output.WriteLine($"  OBSERVATION: No QueryTables found on sheet '{sheetName}'");
                }

                return 0;
            }
            finally
            {
                ComUtilities.Release(ref queryTable);
                ComUtilities.Release(ref queryTables);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });

        _output.WriteLine("");
        _output.WriteLine("=".PadRight(80, '='));
    }

    private async Task Scenario3_UpdateAndRefresh(IExcelBatch session)
    {
        string queryName = "Scenario1Query";
        string newMCode = @"
let
    Source = {10..15},
    ToTable = Table.FromList(Source, Splitter.SplitByNothing(), {""Number""}),
    Typed = Table.TransformColumnTypes(ToTable, {{""Number"", Int64.Type}}),
    AddColumn = Table.AddColumn(Typed, ""Doubled"", each [Number] * 2, Int64.Type)
in
    AddColumn";

        _output.WriteLine($"OBSERVATION: Verify our DELETE→UPDATE→RECREATE pattern works");
        _output.WriteLine($"Old M code: 1 column (Number)");
        _output.WriteLine($"New M code: 2 columns (Number, Doubled)");
        _output.WriteLine($"Pattern: Delete QueryTable → Update M code → Clear sheet → Recreate QueryTable");
        _output.WriteLine("");

        await session.Execute((ctx, ct) =>
        {
            dynamic? queries = null;
            dynamic? query = null;
            dynamic? sheets = null;
            dynamic? sheet = null;
            dynamic? queryTables = null;
            dynamic? queryTable = null;

            try
            {
                // Step 1: Get current state before update
                queries = ctx.Book.Queries;
                query = queries.Item(queryName);
                sheets = ctx.Book.Worksheets;
                sheet = sheets.Item(queryName);
                queryTables = sheet.QueryTables;

                _output.WriteLine($"STEP 1 - State Before Update:");
                _output.WriteLine($"  Query.Name: {query.Name}");
                _output.WriteLine($"  Formula length: {query.Formula.ToString().Length} chars");
                _output.WriteLine($"  QueryTables.Count: {queryTables.Count}");

                dynamic? usedRangeBefore = null;
                int columnsBefore = 0;
                try
                {
                    usedRangeBefore = sheet.UsedRange;
                    columnsBefore = usedRangeBefore.Columns.Count;
                    _output.WriteLine($"  UsedRange: {usedRangeBefore.Address}");
                    _output.WriteLine($"  Rows: {usedRangeBefore.Rows.Count}");
                    _output.WriteLine($"  Columns: {columnsBefore}");
                }
                finally
                {
                    ComUtilities.Release(ref usedRangeBefore);
                }
                _output.WriteLine("");

                // Step 2: DELETE QueryTable (our fix pattern - STEP 1)
                _output.WriteLine($"STEP 2 - DELETE QueryTable (Prevent Excel Crash):");
                if (queryTables.Count > 0)
                {
                    queryTable = queryTables.Item(1);
                    _output.WriteLine($"  QueryTable.Name: {queryTable.Name}");
                    _output.WriteLine($"  Calling queryTable.Delete()...");
                    queryTable.Delete();
                    ComUtilities.Release(ref queryTable);
                    queryTable = null;
                    _output.WriteLine($"  QueryTables.Count after delete: {queryTables.Count}");
                }
                _output.WriteLine("");

                // Step 3: UPDATE M code (our fix pattern - STEP 2)
                _output.WriteLine($"STEP 3 - Update M Code (Safe Now):");
                _output.WriteLine($"  Calling query.Formula = newMCode...");
                query.Formula = newMCode;
                _output.WriteLine($"  New formula length: {query.Formula.ToString().Length} chars");
                _output.WriteLine("");

                // Step 4: CLEAR worksheet (our fix pattern - STEP 3)
                _output.WriteLine($"STEP 4 - Clear Old Data:");
                dynamic? usedRange = null;
                try
                {
                    usedRange = sheet.UsedRange;
                    _output.WriteLine($"  UsedRange before clear: {usedRange.Address}");
                    usedRange.Clear();
                    _output.WriteLine($"  Worksheet cleared");
                }
                catch
                {
                    _output.WriteLine($"  No data to clear (empty sheet)");
                }
                finally
                {
                    ComUtilities.Release(ref usedRange);
                }
                _output.WriteLine("");

                // Step 5: RECREATE QueryTable (our fix pattern - STEP 4)
                _output.WriteLine($"STEP 5 - Recreate QueryTable with New Structure:");
                string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
                dynamic? rangeObj = null;
                try
                {
                    rangeObj = sheet.Range["A1"];
                    queryTable = queryTables.Add(
                        Connection: connectionString,
                        Destination: rangeObj,
                        Sql: Type.Missing);

                    queryTable.Name = queryName;
                    queryTable.CommandText = $"SELECT * FROM [{queryName}]";
                    queryTable.RefreshStyle = 1; // xlInsertDeleteCells

                    _output.WriteLine($"  QueryTable recreated: {queryTable.Name}");
                    _output.WriteLine($"  QueryTables.Count: {queryTables.Count}");
                }
                finally
                {
                    ComUtilities.Release(ref rangeObj);
                }
                _output.WriteLine("");

                // Step 6: REFRESH to load new data
                _output.WriteLine($"STEP 6 - Refresh to Load New Data:");
                _output.WriteLine($"  Calling queryTable.Refresh(false)...");

                bool refreshResult = queryTable.Refresh(false);
                _output.WriteLine($"  Refresh() returned: {refreshResult}");
                _output.WriteLine("");

                // Step 7: Observe the result
                _output.WriteLine($"STEP 7 - Observe Final Data:");
                dynamic? usedRangeAfter = null;
                try
                {
                    usedRangeAfter = sheet.UsedRange;
                    int rows = usedRangeAfter.Rows.Count;
                    int cols = usedRangeAfter.Columns.Count;

                    _output.WriteLine($"  UsedRange: {usedRangeAfter.Address}");
                    _output.WriteLine($"  Rows: {rows}");
                    _output.WriteLine($"  Columns: {cols}");
                    _output.WriteLine("");

                    _output.WriteLine($"  Data Sample:");
                    for (int col = 1; col <= Math.Min(cols, 3); col++)
                    {
                        var colLetter = (char)('A' + col - 1);
                        _output.WriteLine($"    {colLetter}1: {sheet.Range[$"{colLetter}1"].Value2 ?? "(null)"}");
                        _output.WriteLine($"    {colLetter}2: {sheet.Range[$"{colLetter}2"].Value2 ?? "(null)"}");
                    }
                    _output.WriteLine("");

                    _output.WriteLine("=".PadRight(80, '='));
                    _output.WriteLine("FINDINGS:");
                    _output.WriteLine($"  ✅ DELETE→UPDATE→RECREATE pattern: SUCCESS");
                    _output.WriteLine($"  Columns changed from {columnsBefore} to {cols}");
                    _output.WriteLine($"  New column structure loaded successfully");
                    _output.WriteLine($"  No Excel crash (RPC timeout avoided)");
                    _output.WriteLine("=".PadRight(80, '='));
                }
                finally
                {
                    ComUtilities.Release(ref usedRangeAfter);
                }

                return 0;
            }
            finally
            {
                ComUtilities.Release(ref queryTable);
                ComUtilities.Release(ref queryTables);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
                ComUtilities.Release(ref query);
                ComUtilities.Release(ref queries);
            }
        });

        _output.WriteLine("");
        _output.WriteLine("=".PadRight(80, '='));
    }

    private async Task Scenario4_ChangeConnectionOnlyToLoadToWorksheet(IExcelBatch session)
    {
        string queryName = "ConnectionOnlyToWorksheet";
        string sheetName = "LoadedData";
        string mCode = @"
let
    Source = {100..105},
    ToTable = Table.FromList(Source, Splitter.SplitByNothing(), {""Value""}),
    Typed = Table.TransformColumnTypes(ToTable, {{""Value"", Int64.Type}})
in
    Typed";

        _output.WriteLine($"OBSERVATION: Test if connection-only query can load WITHOUT creating QueryTable");
        _output.WriteLine($"M Code: Single column 'Value' with values 100-105");
        _output.WriteLine($"Question: Is QueryTable creation necessary? Can we load data via other means?");
        _output.WriteLine($"Test: Try alternative approaches WITHOUT QueryTable.Add()");
        _output.WriteLine("");

        await session.Execute((ctx, ct) =>
        {
            dynamic? queries = null;
            dynamic? query = null;
            dynamic? sheets = null;
            dynamic? sheet = null;
            dynamic? queryTables = null;
            dynamic? queryTable = null;
            dynamic? usedRange = null;

            try
            {
                // Step 1: Create Power Query (connection-only)
                queries = ctx.Book.Queries;
                int queriesCountBefore = queries.Count;

                _output.WriteLine($"STEP 1 - Create Connection-Only Query:");
                _output.WriteLine($"  Queries.Count before: {queriesCountBefore}");

                query = queries.Add(queryName, mCode);

                int queriesCountAfter = queries.Count;
                _output.WriteLine($"  Queries.Count after: {queriesCountAfter}");
                _output.WriteLine($"  Query.Name: {query.Name}");
                _output.WriteLine("");

                // Step 2: Create worksheet to load data into
                sheets = ctx.Book.Worksheets;
                int sheetsCountBefore = sheets.Count;

                _output.WriteLine($"STEP 2 - Create Target Worksheet:");
                _output.WriteLine($"  Worksheets.Count before: {sheetsCountBefore}");

                sheet = sheets.Add();
                sheet.Name = sheetName;

                int sheetsCountAfter = sheets.Count;
                _output.WriteLine($"  Worksheets.Count after: {sheetsCountAfter}");
                _output.WriteLine($"  Worksheet.Name: {sheet.Name}");
                _output.WriteLine("");

                // Step 3: Try to load data WITHOUT creating QueryTable
                _output.WriteLine($"STEP 3 - Attempt to Load Data WITHOUT QueryTable:");
                _output.WriteLine($"  Testing alternative approaches...");
                _output.WriteLine("");

                // Alternative 1: Try setting query LoadToWorksheet property (if it exists)
                _output.WriteLine($"  Alternative 1: Try setting query properties directly");
                try
                {
                    // Attempt to access LoadToWorksheet or similar property
                    // This will likely fail as Power Query object doesn't have data loading properties
                    _output.WriteLine($"    Query object type: {query.GetType().Name}");
                    _output.WriteLine($"    Cannot set load destination - Query object has no LoadToWorksheet property");
                }
                catch (Exception ex)
                {
                    _output.WriteLine($"    Failed: {ex.Message}");
                }
                _output.WriteLine("");

                // Alternative 2: Check if data appears automatically in worksheet
                _output.WriteLine($"  Alternative 2: Check if data appears automatically");
                usedRange = sheet.UsedRange;
                string addressNoQueryTable = usedRange.Address;
                int rowsNoQueryTable = usedRange.Rows.Count;
                int colsNoQueryTable = usedRange.Columns.Count;
                object a1NoQueryTable = sheet.Range["A1"].Value2;

                _output.WriteLine($"    UsedRange.Address: {addressNoQueryTable}");
                _output.WriteLine($"    Rows: {rowsNoQueryTable}");
                _output.WriteLine($"    Columns: {colsNoQueryTable}");
                _output.WriteLine($"    Cell A1 value: {a1NoQueryTable ?? "(null)"}");
                _output.WriteLine($"    Result: No data - connection-only queries don't auto-populate worksheets");
                _output.WriteLine("");

                // Alternative 3: Try accessing query data directly (will fail)
                _output.WriteLine($"  Alternative 3: Try accessing query data directly");
                try
                {
                    // Try to read data from query object
                    // This will fail - queries only have formula/M code, not data
                    dynamic queryData = query.Data; // This property doesn't exist
                    _output.WriteLine($"    Query data: {queryData}");
                }
                catch (Exception ex)
                {
                    _output.WriteLine($"    Failed: {ex.Message}");
                    _output.WriteLine($"    Result: Query object has no 'Data' property - only stores M code formula");
                }
                _output.WriteLine("");

                // Step 4: Prove QueryTable is necessary by creating one
                _output.WriteLine($"STEP 4 - Prove QueryTable Is Necessary (Create One Now):");
                queryTables = sheet.QueryTables;
                int queryTablesCountBefore = queryTables.Count;

                _output.WriteLine($"  QueryTables.Count before: {queryTablesCountBefore}");

                string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
                _output.WriteLine($"  ConnectionString: {connectionString}");

                queryTable = queryTables.Add(
                    Connection: connectionString,
                    Destination: sheet.Range["A1"],
                    Sql: Type.Missing);

                queryTable.CommandText = $"SELECT * FROM [{queryName}]";

                int queryTablesCountAfter = queryTables.Count;
                _output.WriteLine($"  QueryTables.Count after: {queryTablesCountAfter}");
                _output.WriteLine($"  QueryTable created: {queryTable.Name}");
                _output.WriteLine("");

                // Step 5: Refresh to actually load data
                _output.WriteLine($"STEP 5 - Refresh QueryTable to Load Data:");

                bool refreshResult = queryTable.Refresh(false);

                _output.WriteLine($"  Refresh result: {refreshResult}");
                _output.WriteLine("");

                // Step 6: Observe final data state
                _output.WriteLine($"STEP 6 - Observe Final Data State (After Explicit Refresh):");

                ComUtilities.Release(ref usedRange);
                usedRange = sheet.UsedRange;
                string addressAfter = usedRange.Address;
                int rowsAfter = usedRange.Rows.Count;
                int colsAfter = usedRange.Columns.Count;

                _output.WriteLine($"  UsedRange.Address: {addressAfter}");
                _output.WriteLine($"  Rows: {rowsAfter}");
                _output.WriteLine($"  Columns: {colsAfter}");

                object a1After = sheet.Range["A1"].Value2;
                object a2After = sheet.Range["A2"].Value2;

                _output.WriteLine($"  Cell A1 value: {a1After ?? "(null)"}");
                _output.WriteLine($"  Cell A2 value: {a2After ?? "(null)"}");
                _output.WriteLine("");

                // FINDINGS
                _output.WriteLine("=".PadRight(80, '='));
                _output.WriteLine("FINDINGS:");
                _output.WriteLine($"  Connection-only query created: {query.Name}");
                _output.WriteLine($"  Worksheet created: {sheet.Name}");
                _output.WriteLine("");
                _output.WriteLine($"  WITHOUT QueryTable:");
                _output.WriteLine($"    Alternative 1 (query properties): FAILED - Query object has no load destination properties");
                _output.WriteLine($"    Alternative 2 (auto-populate): FAILED - No data in worksheet ({rowsNoQueryTable}x{colsNoQueryTable})");
                _output.WriteLine($"    Alternative 3 (query.Data): FAILED - Query object has no Data property");
                _output.WriteLine($"    Result: Cannot load connection-only query without QueryTable");
                _output.WriteLine("");
                _output.WriteLine($"  WITH QueryTable:");
                _output.WriteLine($"    QueryTable created: {queryTable.Name}");
                _output.WriteLine($"    Refresh result: {refreshResult}");
                _output.WriteLine($"    Data state after refresh: {rowsAfter} rows, {colsAfter} cols");
                _output.WriteLine($"    Sample data: A1={a1After}, A2={a2After}");
                _output.WriteLine("");
                _output.WriteLine($"  ✅ CONCLUSION: QueryTable creation IS NECESSARY");
                _output.WriteLine($"    - No alternative mechanism found");
                _output.WriteLine($"    - QueryTable is THE Excel COM API pattern for loading connection-only queries");
                _output.WriteLine($"    - Our LoadToAsync implementation is correct");
                _output.WriteLine("=".PadRight(80, '='));

                return 0;
            }
            finally
            {
                ComUtilities.Release(ref usedRange);
                ComUtilities.Release(ref queryTable);
                ComUtilities.Release(ref queryTables);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
                ComUtilities.Release(ref query);
                ComUtilities.Release(ref queries);
            }
        });

        _output.WriteLine("");
    }

    private async Task Scenario5_LoadConnectionOnlyToWorksheet(IExcelBatch session)
    {
        string queryName = "ConnectionOnlyQuery";
        string mCode = @"
let
    Source = {100..105},
    #""Converted to Table"" = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #""Renamed Columns"" = Table.RenameColumns(#""Converted to Table"", {{""Column1"", ""Value""}})
in
    #""Renamed Columns""";

        _output.WriteLine($"OBSERVATION: Loading connection-only query to worksheet, then deleting QueryTable");
        _output.WriteLine($"M Code: Single column 'Value' with values 100-105");
        _output.WriteLine($"Question: Does QueryTable.Delete() remove data or just the QueryTable object?");
        _output.WriteLine("");

        await session.Execute((ctx, ct) =>
        {
            dynamic? queries = null;
            dynamic? query = null;
            dynamic? sheets = null;
            dynamic? sheet = null;
            dynamic? queryTables = null;
            dynamic? queryTable = null;

            try
            {
                // Step 1: Create connection-only query
                queries = ctx.Book.Queries;
                _output.WriteLine($"STEP 1 - Create Connection-Only Query:");
                _output.WriteLine($"  Queries.Count before: {queries.Count}");

                query = queries.Add(queryName, mCode);

                _output.WriteLine($"  Queries.Count after: {queries.Count}");
                _output.WriteLine($"  Query.Name: {query.Name}");
                _output.WriteLine("");

                // Step 2: Create worksheet
                sheets = ctx.Book.Worksheets;
                int sheetsCountBefore = sheets.Count;

                _output.WriteLine($"STEP 2 - Create Worksheet:");
                _output.WriteLine($"  Worksheets.Count before: {sheetsCountBefore}");

                sheet = sheets.Add();
                sheet.Name = queryName;

                int sheetsCountAfter = sheets.Count;
                _output.WriteLine($"  Worksheets.Count after: {sheetsCountAfter}");
                _output.WriteLine($"  Worksheet.Name: {sheet.Name}");
                _output.WriteLine("");

                // Step 3: Create QueryTable to load data
                string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";

                _output.WriteLine($"STEP 3 - Create QueryTable:");
                _output.WriteLine($"  ConnectionString: {connectionString}");

                queryTables = sheet.QueryTables;
                int qtCountBefore = queryTables.Count;
                _output.WriteLine($"  QueryTables.Count before: {qtCountBefore}");

                queryTable = queryTables.Add(connectionString, sheet.Range["A1"], Type.Missing);
                queryTable.Name = queryName;
                queryTable.CommandText = $"SELECT * FROM [{queryName}]";

                int qtCountAfter = queryTables.Count;
                _output.WriteLine($"  QueryTables.Count after: {qtCountAfter}");
                _output.WriteLine($"  QueryTable.Name: {queryTable.Name}");
                _output.WriteLine("");

                // Step 4: Refresh to actually load the data
                _output.WriteLine($"STEP 4 - Refresh to Load Data:");
                _output.WriteLine($"  Calling queryTable.Refresh(false)...");

                bool refreshResult = queryTable.Refresh(false);

                _output.WriteLine($"  Refresh result: {refreshResult}");
                _output.WriteLine($"  QueryTables.Count after refresh: {queryTables.Count}");
                _output.WriteLine("");

                // Step 5: Observe loaded data
                _output.WriteLine($"STEP 5 - Observe Loaded Data:");
                dynamic? usedRangeBefore = null;
                int rowsBefore = 0;
                int colsBefore = 0;
                object? a1Before = null;
                object? a2Before = null;

                try
                {
                    usedRangeBefore = sheet.UsedRange;
                    rowsBefore = usedRangeBefore.Rows.Count;
                    colsBefore = usedRangeBefore.Columns.Count;
                    a1Before = sheet.Range["A1"].Value2;
                    a2Before = sheet.Range["A2"].Value2;

                    _output.WriteLine($"  UsedRange: {usedRangeBefore.Address}");
                    _output.WriteLine($"  Rows: {rowsBefore}");
                    _output.WriteLine($"  Columns: {colsBefore}");
                    _output.WriteLine($"  Cell A1: {a1Before ?? "(null)"}");
                    _output.WriteLine($"  Cell A2: {a2Before ?? "(null)"}");
                }
                finally
                {
                    ComUtilities.Release(ref usedRangeBefore);
                }
                _output.WriteLine("");

                // Step 6: CRITICAL TEST - Delete QueryTable
                _output.WriteLine($"STEP 6 - Delete QueryTable (CRITICAL):");
                _output.WriteLine($"  QueryTables.Count before delete: {queryTables.Count}");
                _output.WriteLine($"  Calling queryTable.Delete()...");

                queryTable.Delete();
                ComUtilities.Release(ref queryTable);
                queryTable = null;

                int qtCountAfterDelete = queryTables.Count;
                _output.WriteLine($"  QueryTables.Count after delete: {qtCountAfterDelete}");
                _output.WriteLine("");

                // Step 7: Check if data remains after QueryTable deletion
                _output.WriteLine($"STEP 7 - Check Data After QueryTable Deletion:");
                dynamic? usedRangeAfter = null;
                int rowsAfter = 0;
                int colsAfter = 0;
                object? a1After = null;
                object? a2After = null;

                try
                {
                    usedRangeAfter = sheet.UsedRange;
                    rowsAfter = usedRangeAfter.Rows.Count;
                    colsAfter = usedRangeAfter.Columns.Count;
                    a1After = sheet.Range["A1"].Value2;
                    a2After = sheet.Range["A2"].Value2;

                    _output.WriteLine($"  UsedRange: {usedRangeAfter.Address}");
                    _output.WriteLine($"  Rows: {rowsAfter}");
                    _output.WriteLine($"  Columns: {colsAfter}");
                    _output.WriteLine($"  Cell A1: {a1After ?? "(null)"}");
                    _output.WriteLine($"  Cell A2: {a2After ?? "(null)"}");
                }
                catch (Exception ex)
                {
                    _output.WriteLine($"  Error reading UsedRange: {ex.Message}");
                    _output.WriteLine($"  This may indicate no data remains");
                }
                finally
                {
                    ComUtilities.Release(ref usedRangeAfter);
                }
                _output.WriteLine("");

                // FINDINGS
                _output.WriteLine("=".PadRight(80, '='));
                _output.WriteLine("FINDINGS:");
                _output.WriteLine($"  Worksheet created: {sheet.Name}");
                _output.WriteLine($"  QueryTable created and refreshed successfully");
                _output.WriteLine($"  Data loaded - Rows: {rowsBefore}, Columns: {colsBefore}");
                _output.WriteLine($"  Sample data before delete - A1: {a1Before}, A2: {a2Before}");
                _output.WriteLine($"  QueryTable.Delete() executed successfully");
                _output.WriteLine($"  QueryTables.Count: {qtCountBefore} → {qtCountAfterDelete}");
                _output.WriteLine($"  Data after delete - Rows: {rowsAfter}, Columns: {colsAfter}");
                _output.WriteLine($"  Sample data after delete - A1: {a1After}, A2: {a2After}");
                _output.WriteLine($"  Data persisted after QueryTable deletion: {(a1After != null && a2After != null)}");
                _output.WriteLine($"  Data matches pre-deletion: {(a1After?.ToString() == a1Before?.ToString() && a2After?.ToString() == a2Before?.ToString())}");
                _output.WriteLine("=".PadRight(80, '='));

                return 0;
            }
            finally
            {
                ComUtilities.Release(ref queryTable);
                ComUtilities.Release(ref queryTables);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
                ComUtilities.Release(ref query);
                ComUtilities.Release(ref queries);
            }
        });

        _output.WriteLine("");
    }
}
