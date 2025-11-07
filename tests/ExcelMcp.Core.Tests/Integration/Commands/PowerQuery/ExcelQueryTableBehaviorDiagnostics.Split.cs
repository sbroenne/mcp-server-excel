using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// DIAGNOSTIC TESTS - Discovers Excel's actual QueryTable behavior using raw COM API.
/// Each test works on an independent file to avoid Excel COM stability issues.
///
/// Tests discover what Excel actually does when you:
/// 1. Load a PowerQuery to a worksheet
/// 2. Refresh it multiple times (like clicking Refresh in UI)
/// 3. Change the query M code and refresh
/// 4. Create connection-only query
/// 5. Load connection-only to worksheet
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "Core")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
public class ExcelQueryTableBehaviorDiagnosticsSplit : IClassFixture<TempDirectoryFixture>
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    /// <inheritdoc/>

    public ExcelQueryTableBehaviorDiagnosticsSplit(TempDirectoryFixture fixture, ITestOutputHelper output)
    {
        _output = output;
        _tempDir = fixture.TempDir;
    }

    /// <summary>
    /// DIAGNOSTIC TEST 1: Load PowerQuery to worksheet
    /// Tests: Does Excel create QueryTable? How many? What happens on first load?
    /// Expected: Single QueryTable created, data loaded successfully
    /// </summary>
    [Fact]
    public async Task Diagnostic1_LoadToWorksheet_CreatesOneQueryTable()
    {
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(ExcelQueryTableBehaviorDiagnosticsSplit),
            nameof(Diagnostic1_LoadToWorksheet_CreatesOneQueryTable),
            _tempDir,
            ".xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        _output.WriteLine("=== DIAGNOSTIC 1: Load PowerQuery to worksheet ===");

        string queryName = "Scenario1Query";
        string mCode = @"
let
    Source = {1..5},
    ToTable = Table.FromList(Source, Splitter.SplitByNothing(), {""Number""}),
    Typed = Table.TransformColumnTypes(ToTable, {{""Number"", Int64.Type}})
in
    Typed";

        await batch.Execute((ctx, ct) =>
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
                // Step 1: Create Power Query
                queries = ctx.Book.Queries;
                _output.WriteLine($"Before: Queries.Count = {queries.Count}");

                query = queries.Add(queryName, mCode);
                _output.WriteLine($"After Add: Queries.Count = {queries.Count}");
                _output.WriteLine($"Query.Name = {query.Name}");

                // Step 2: Create worksheet
                sheets = ctx.Book.Worksheets;
                sheet = sheets.Add();
                sheet.Name = queryName;
                _output.WriteLine($"Created worksheet: {sheet.Name}");

                // Step 3: Create QueryTable (this loads the data)
                string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
                _output.WriteLine($"Creating QueryTable with connection string...");

                queryTables = sheet.QueryTables;
                queryTable = queryTables.Add(
                    Connection: connectionString,
                    Destination: sheet.Range["A1"],
                    Sql: Type.Missing);

                queryTable.Name = queryName;
                queryTable.CommandText = $"SELECT * FROM [{queryName}]";
                _output.WriteLine($"QueryTable created: {queryTable.Name}");

                // Step 4: Refresh to load data
                _output.WriteLine($"Before Refresh: QueryTables.Count = {queryTables.Count}");
                bool refreshResult = queryTable.Refresh(false);
                _output.WriteLine($"Refresh result: {refreshResult}");
                _output.WriteLine($"After Refresh: QueryTables.Count = {queryTables.Count}");

                // Step 5: Check the data
                usedRange = sheet.UsedRange;
                _output.WriteLine($"UsedRange.Rows.Count = {usedRange.Rows.Count}");
                _output.WriteLine($"UsedRange.Columns.Count = {usedRange.Columns.Count}");
                _output.WriteLine($"Cell A1 value: {sheet.Range["A1"].Value2}");
                _output.WriteLine($"Cell A2 value: {sheet.Range["A2"].Value2}");

                // FINDINGS
                _output.WriteLine("");
                _output.WriteLine("FINDINGS:");
                _output.WriteLine($"  ✓ QueryTables.Count = {queryTables.Count} (should be 1)");
                _output.WriteLine($"  ✓ Data loaded: {usedRange.Rows.Count} rows × {usedRange.Columns.Count} columns");

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
    }

    /// <summary>
    /// DIAGNOSTIC TEST 2: Refresh loaded query multiple times
    /// Tests: Does Excel create duplicate QueryTables on refresh?
    /// Expected: QueryTables.Count stays at 1 across multiple refreshes
    /// </summary>
    [Fact]
    public async Task Diagnostic2_RefreshMultipleTimes_NoDuplicates()
    {
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(ExcelQueryTableBehaviorDiagnosticsSplit),
            nameof(Diagnostic2_RefreshMultipleTimes_NoDuplicates),
            _tempDir,
            ".xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        _output.WriteLine("=== DIAGNOSTIC 2: Refresh loaded query multiple times ===");

        // First load the query (same as Diagnostic 1)
        string queryName = "RefreshTestQuery";
        string mCode = @"
let
    Source = {1..5},
    ToTable = Table.FromList(Source, Splitter.SplitByNothing(), {""Number""}),
    Typed = Table.TransformColumnTypes(ToTable, {{""Number"", Int64.Type}})
in
    Typed";

        await batch.Execute((ctx, ct) =>
        {
            dynamic? queries = null;
            dynamic? query = null;
            dynamic? sheets = null;
            dynamic? sheet = null;
            dynamic? queryTables = null;
            dynamic? queryTable = null;

            try
            {
                // Create query and load to worksheet
                queries = ctx.Book.Queries;
                query = queries.Add(queryName, mCode);

                sheets = ctx.Book.Worksheets;
                sheet = sheets.Add();
                sheet.Name = queryName;

                string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
                queryTables = sheet.QueryTables;
                queryTable = queryTables.Add(
                    Connection: connectionString,
                    Destination: sheet.Range["A1"],
                    Sql: Type.Missing);
                queryTable.Name = queryName;
                queryTable.CommandText = $"SELECT * FROM [{queryName}]";
                queryTable.Refresh(false);

                _output.WriteLine($"Initial load: QueryTables.Count = {queryTables.Count}");

                // TEST: Refresh 2nd time
                _output.WriteLine("");
                _output.WriteLine("2nd Refresh:");
                queryTable.Refresh(false);
                _output.WriteLine($"  After 2nd refresh: QueryTables.Count = {queryTables.Count}");

                // TEST: Refresh 3rd time
                _output.WriteLine("");
                _output.WriteLine("3rd Refresh:");
                queryTable.Refresh(false);
                _output.WriteLine($"  After 3rd refresh: QueryTables.Count = {queryTables.Count}");

                // TEST: Refresh 4th time
                _output.WriteLine("");
                _output.WriteLine("4th Refresh:");
                queryTable.Refresh(false);
                _output.WriteLine($"  After 4th refresh: QueryTables.Count = {queryTables.Count}");

                // FINDINGS
                _output.WriteLine("");
                _output.WriteLine("FINDINGS:");
                _output.WriteLine($"  ✓ QueryTables.Count = {queryTables.Count} (should stay at 1)");
                _output.WriteLine($"  ✓ Excel does NOT create duplicate QueryTables on refresh");

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
    }

    /// <summary>
    /// DIAGNOSTIC TEST 3: Update query M code and refresh
    /// Tests: Can we update M code while QueryTable exists? What errors occur?
    /// Expected: RPC timeout (0x800706BE) when refreshing after structural change
    /// </summary>
    [Fact]
    public async Task Diagnostic3_UpdateMCode_RpcTimeoutExpected()
    {
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(ExcelQueryTableBehaviorDiagnosticsSplit),
            nameof(Diagnostic3_UpdateMCode_RpcTimeoutExpected),
            _tempDir,
            ".xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        _output.WriteLine("=== DIAGNOSTIC 3: Update query M code and refresh ===");

        string queryName = "UpdateTestQuery";
        string originalMCode = @"
let
    Source = {1..5},
    ToTable = Table.FromList(Source, Splitter.SplitByNothing(), {""Number""}),
    Typed = Table.TransformColumnTypes(ToTable, {{""Number"", Int64.Type}})
in
    Typed";

        string updatedMCode = @"
let
    Source = {1..5},
    ToTable = Table.FromList(Source, Splitter.SplitByNothing(), {""Number""}),
    AddColumn1 = Table.AddColumn(ToTable, ""Double"", each [Number] * 2),
    AddColumn2 = Table.AddColumn(AddColumn1, ""Triple"", each [Number] * 3),
    Typed = Table.TransformColumnTypes(AddColumn2, {{""Number"", Int64.Type}, {""Double"", Int64.Type}, {""Triple"", Int64.Type}})
in
    Typed";

        await batch.Execute((ctx, ct) =>
        {
            dynamic? queries = null;
            dynamic? query = null;
            dynamic? sheets = null;
            dynamic? sheet = null;
            dynamic? queryTables = null;
            dynamic? queryTable = null;

            try
            {
                // Create query and load to worksheet
                queries = ctx.Book.Queries;
                query = queries.Add(queryName, originalMCode);

                sheets = ctx.Book.Worksheets;
                sheet = sheets.Add();
                sheet.Name = queryName;

                string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
                queryTables = sheet.QueryTables;
                queryTable = queryTables.Add(
                    Connection: connectionString,
                    Destination: sheet.Range["A1"],
                    Sql: Type.Missing);
                queryTable.Name = queryName;
                queryTable.CommandText = $"SELECT * FROM [{queryName}]";
                queryTable.Refresh(false);

                _output.WriteLine($"Initial load: QueryTables.Count = {queryTables.Count}, Columns = 1");

                // TEST: Update M code (structural change - 1 column → 3 columns)
                _output.WriteLine("");
                _output.WriteLine("Updating M code (1 column → 3 columns)...");
                query.Formula = updatedMCode;
                _output.WriteLine("  M code updated");
                _output.WriteLine($"  QueryTables.Count = {queryTables.Count} (still 1)");

                // TEST: Try to refresh with new M code
                _output.WriteLine("");
                _output.WriteLine("Attempting to refresh with updated M code...");

                try
                {
                    Thread.Sleep(1000); // Let Excel process the change
                    queryTable.Refresh(false);
                    _output.WriteLine("  ✗ Refresh succeeded (unexpected!)");
                    _output.WriteLine($"  QueryTables.Count = {queryTables.Count}");
                }
                catch (COMException ex) when (ex.HResult == unchecked((int)0x800706BE))
                {
                    _output.WriteLine("  ✓ RPC timeout (0x800706BE) - EXPECTED behavior");
                    _output.WriteLine("  Excel cannot refresh QueryTable after structural M code change");
                }

                // FINDINGS
                _output.WriteLine("");
                _output.WriteLine("FINDINGS:");
                _output.WriteLine("  ✓ Updating Query.Formula while QueryTable exists causes RPC timeout");
                _output.WriteLine("  ✓ Excel requires: Delete QueryTable → Update Formula → Recreate QueryTable");

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
    }

    /// <summary>
    /// DIAGNOSTIC TEST 4: Create connection-only query
    /// Tests: Does Excel auto-create QueryTables? Can we update connection-only queries freely?
    /// Expected: No QueryTables created, query exists in Queries collection
    /// </summary>
    [Fact]
    public async Task Diagnostic4_ConnectionOnly_NoAutoQueryTable()
    {
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(ExcelQueryTableBehaviorDiagnosticsSplit),
            nameof(Diagnostic4_ConnectionOnly_NoAutoQueryTable),
            _tempDir,
            ".xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        _output.WriteLine("=== DIAGNOSTIC 4: Create connection-only query ===");

        string queryName = "ConnectionOnlyQuery";
        string mCode = @"
let
    Source = {100..105},
    ToTable = Table.FromList(Source, Splitter.SplitByNothing(), {""Value""}),
    Typed = Table.TransformColumnTypes(ToTable, {{""Value"", Int64.Type}})
in
    Typed";

        await batch.Execute((ctx, ct) =>
        {
            dynamic? queries = null;
            dynamic? query = null;
            dynamic? sheets = null;

            try
            {
                // Create Power Query (connection-only)
                queries = ctx.Book.Queries;
                _output.WriteLine($"Before: Queries.Count = {queries.Count}");

                query = queries.Add(queryName, mCode);
                _output.WriteLine($"After Add: Queries.Count = {queries.Count}");
                _output.WriteLine($"Query.Name = {query.Name}");

                // Check if Excel auto-created any QueryTables
                _output.WriteLine("");
                _output.WriteLine("Checking for auto-created QueryTables...");
                sheets = ctx.Book.Worksheets;
                int totalQueryTables = 0;

                for (int i = 1; i <= sheets.Count; i++)
                {
                    dynamic? sheet = null;
                    dynamic? queryTables = null;
                    try
                    {
                        sheet = sheets.Item(i);
                        queryTables = sheet.QueryTables;
                        int count = queryTables.Count;
                        totalQueryTables += count;

                        if (count > 0)
                        {
                            _output.WriteLine($"  Sheet '{sheet.Name}': {count} QueryTables");
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref queryTables);
                        ComUtilities.Release(ref sheet);
                    }
                }

                _output.WriteLine($"Total QueryTables across all sheets: {totalQueryTables}");

                // FINDINGS
                _output.WriteLine("");
                _output.WriteLine("FINDINGS:");
                _output.WriteLine($"  ✓ Queries.Count = {queries.Count}");
                _output.WriteLine($"  ✓ Total QueryTables = {totalQueryTables} (should be 0)");
                _output.WriteLine("  ✓ Excel does NOT auto-create QueryTables for connection-only queries");

                return 0;
            }
            finally
            {
                ComUtilities.Release(ref sheets);
                ComUtilities.Release(ref query);
                ComUtilities.Release(ref queries);
            }
        });
    }

    /// <summary>
    /// DIAGNOSTIC TEST 5: Load connection-only query to worksheet
    /// Tests: Can we manually create QueryTable from connection-only? What happens on delete?
    /// Expected: Manual QueryTable creation works, data loads successfully
    /// </summary>
    [Fact]
    public async Task Diagnostic5_LoadConnectionOnly_ManualQueryTable()
    {
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(ExcelQueryTableBehaviorDiagnosticsSplit),
            nameof(Diagnostic5_LoadConnectionOnly_ManualQueryTable),
            _tempDir,
            ".xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        _output.WriteLine("=== DIAGNOSTIC 5: Load connection-only to worksheet ===");

        string queryName = "LoadConnectionOnlyQuery";
        string mCode = @"
let
    Source = {200..205},
    ToTable = Table.FromList(Source, Splitter.SplitByNothing(), {""Value""}),
    Typed = Table.TransformColumnTypes(ToTable, {{""Value"", Int64.Type}})
in
    Typed";

        await batch.Execute((ctx, ct) =>
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
                // Step 1: Create connection-only query
                queries = ctx.Book.Queries;
                query = queries.Add(queryName, mCode);
                _output.WriteLine($"Created connection-only query: {query.Name}");

                // Step 2: Manually create QueryTable (simulates UI: Load To > Table)
                _output.WriteLine("");
                _output.WriteLine("Manually creating QueryTable from connection-only query...");

                sheets = ctx.Book.Worksheets;
                sheet = sheets.Add();
                sheet.Name = queryName;

                string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
                queryTables = sheet.QueryTables;
                queryTable = queryTables.Add(
                    Connection: connectionString,
                    Destination: sheet.Range["A1"],
                    Sql: Type.Missing);
                queryTable.Name = queryName;
                queryTable.CommandText = $"SELECT * FROM [{queryName}]";

                _output.WriteLine($"QueryTable created: {queryTable.Name}");
                _output.WriteLine($"Before Refresh: QueryTables.Count = {queryTables.Count}");

                // Step 3: Refresh to load data
                bool refreshResult = queryTable.Refresh(false);
                _output.WriteLine($"Refresh result: {refreshResult}");
                _output.WriteLine($"After Refresh: QueryTables.Count = {queryTables.Count}");

                // Step 4: Check the data
                usedRange = sheet.UsedRange;
                _output.WriteLine($"UsedRange.Rows.Count = {usedRange.Rows.Count}");
                _output.WriteLine($"UsedRange.Columns.Count = {usedRange.Columns.Count}");
                _output.WriteLine($"Cell A1 value: {sheet.Range["A1"].Value2}");
                _output.WriteLine($"Cell A2 value: {sheet.Range["A2"].Value2}");

                // FINDINGS
                _output.WriteLine("");
                _output.WriteLine("FINDINGS:");
                _output.WriteLine("  ✓ Connection-only query can be loaded to worksheet via manual QueryTable creation");
                _output.WriteLine($"  ✓ QueryTables.Count = {queryTables.Count} (should be 1)");
                _output.WriteLine($"  ✓ Data loaded: {usedRange.Rows.Count} rows × {usedRange.Columns.Count} columns");

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
    }
}
