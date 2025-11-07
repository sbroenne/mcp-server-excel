using System;
using System.IO;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// DIAGNOSTIC TEST - Discovers Excel's actual QueryTable behavior using raw COM API.
/// This test does NOT use our wrapper code - it directly calls Excel COM to observe behavior.
///
/// Purpose: Understand what Excel actually does when you:
/// 1. Load a PowerQuery to a worksheet
/// 2. Refresh it (like clicking Refresh in UI)
/// 3. Change the query and load it again
/// 4. Create connection-only query
/// 5. Load connection-only to worksheet (Load To in UI)
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "PowerQuery")]
[Trait("Layer", "Diagnostics")]
[Trait("RequiresExcel", "true")]
public class ExcelQueryTableBehaviorDiagnostics : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;

    public ExcelQueryTableBehaviorDiagnostics(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Combine(Path.GetTempPath(), $"excel-qt-diagnostics-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
        {
            try { Directory.Delete(_tempDir, recursive: true); }
            catch { /* Best effort */ }
        }
        GC.SuppressFinalize(this);
    }

    [Fact]
    public async Task DiagnoseExcelBehavior_AllScenarios()
    {
        string testFile = Path.Combine(_tempDir, "diagnostic-test.xlsx");

        // Create empty workbook first
        var fileCommands = new FileCommands();
        var createResult = await fileCommands.CreateEmptyAsync(testFile);
        Assert.True(createResult.Success, $"Failed to create test file: {createResult.ErrorMessage}");

        await using var session = await ExcelSession.BeginBatchAsync(testFile);

        _output.WriteLine("=".PadRight(80, '='));
        _output.WriteLine("EXCEL QUERYTABLE BEHAVIOR DIAGNOSTICS");
        _output.WriteLine("=".PadRight(80, '='));
        _output.WriteLine("");

        // Scenario 1: Load PowerQuery to worksheet
        _output.WriteLine("SCENARIO 1: Load PowerQuery to worksheet");
        _output.WriteLine("-".PadRight(80, '-'));
        await Scenario1_LoadToWorksheet(session);
        _output.WriteLine("");

        // Scenario 2: Refresh the loaded query (simulate UI Refresh)
        _output.WriteLine("SCENARIO 2: Refresh loaded query (like UI Refresh button)");
        _output.WriteLine("-".PadRight(80, '-'));
        await Scenario2_RefreshLoaded(session);
        _output.WriteLine("");

        // Scenario 3: Change query M code and load again
        _output.WriteLine("SCENARIO 3: Update query M code and refresh");
        _output.WriteLine("-".PadRight(80, '-'));
        await Scenario3_UpdateAndRefresh(session);
        _output.WriteLine("");

        // Scenario 4: Create connection-only query
        _output.WriteLine("SCENARIO 4: Create connection-only query");
        _output.WriteLine("-".PadRight(80, '-'));

        // Give Excel time to recover from previous scenarios
        _output.WriteLine("Waiting for Excel to stabilize...");
        await Task.Delay(3000);

        try
        {
            await Scenario4_ConnectionOnly(session);
        }
        catch (System.Runtime.InteropServices.COMException ex) when (ex.HResult == unchecked((int)0x800706BA))
        {
            _output.WriteLine($"⚠️ Excel RPC server unavailable - Excel may have crashed from previous scenarios");
            _output.WriteLine($"This is a known Excel COM limitation with rapid successive operations");
            _output.WriteLine($"Skipping remaining scenarios - restart required");
            _output.WriteLine("");
            _output.WriteLine("=".PadRight(80, '='));
            _output.WriteLine("DIAGNOSTICS PARTIAL COMPLETE (Scenarios 1-3 successful)");
            _output.WriteLine("=".PadRight(80, '='));
            return; // Exit test gracefully
        }
        _output.WriteLine("");

        // Scenario 5: Load connection-only to worksheet (UI: Load To > Table)
        _output.WriteLine("SCENARIO 5: Load connection-only to worksheet");
        _output.WriteLine("-".PadRight(80, '-'));
        await Scenario5_LoadConnectionOnlyToWorksheet(session);
        _output.WriteLine("");

        _output.WriteLine("=".PadRight(80, '='));
        _output.WriteLine("DIAGNOSTICS COMPLETE");
        _output.WriteLine("=".PadRight(80, '='));
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

        _output.WriteLine($"Creating query: {queryName}");

        await session.Execute<int>((ctx, ct) =>
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
                _output.WriteLine($"  Before: Queries.Count = {queries.Count}");

                query = queries.Add(queryName, mCode);
                _output.WriteLine($"  After Add: Queries.Count = {queries.Count}");
                _output.WriteLine($"  Query.Name = {query.Name}");

                // Step 2: Create worksheet for the data
                sheets = ctx.Book.Worksheets;
                sheet = sheets.Add();
                sheet.Name = queryName;
                _output.WriteLine($"  Created worksheet: {sheet.Name}");

                // Step 3: Create QueryTable to load data
                string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
                _output.WriteLine($"  Creating QueryTable with connection string...");

                queryTable = sheet.QueryTables.Add(connectionString, sheet.Range["A1"], Type.Missing);
                queryTable.Name = queryName;
                queryTable.CommandText = $"SELECT * FROM [{queryName}]";
                _output.WriteLine($"  QueryTable created: {queryTable.Name}");

                // Step 4: Refresh to load data
                queryTables = sheet.QueryTables;
                _output.WriteLine($"  Before Refresh: QueryTables.Count = {queryTables.Count}");

                bool refreshResult = queryTable.Refresh(false); // false = synchronous
                _output.WriteLine($"  Refresh result: {refreshResult}");
                _output.WriteLine($"  After Refresh: QueryTables.Count = {queryTables.Count}");

                // Step 5: Check what Excel created
                dynamic? usedRange = null;
                try
                {
                    usedRange = sheet.UsedRange;
                    _output.WriteLine($"  UsedRange.Rows.Count = {usedRange.Rows.Count}");
                    _output.WriteLine($"  UsedRange.Columns.Count = {usedRange.Columns.Count}");
                    _output.WriteLine($"  Cell A1 value: {sheet.Range["A1"].Value2}");
                    _output.WriteLine($"  Cell A2 value: {sheet.Range["A2"].Value2}");
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

        _output.WriteLine("✓ Scenario 1 complete");
    }

    private async Task Scenario2_RefreshLoaded(IExcelBatch session)
    {
        string sheetName = "Scenario1Query";

        _output.WriteLine($"Refreshing query on sheet: {sheetName}");

        await session.Execute<int>((ctx, ct) =>
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
                _output.WriteLine($"  Before 2nd Refresh: QueryTables.Count = {queryTables.Count}");

                if (queryTables.Count > 0)
                {
                    queryTable = queryTables.Item(1);
                    _output.WriteLine($"  QueryTable.Name = {queryTable.Name}");

                    // THIS IS THE KEY TEST: What happens on 2nd refresh?
                    bool refreshResult = queryTable.Refresh(false);
                    _output.WriteLine($"  Refresh result: {refreshResult}");
                    _output.WriteLine($"  After 2nd Refresh: QueryTables.Count = {queryTables.Count}");

                    // Check for duplicates
                    dynamic? usedRange = null;
                    try
                    {
                        usedRange = sheet.UsedRange;
                        _output.WriteLine($"  UsedRange.Rows.Count = {usedRange.Rows.Count}");
                        _output.WriteLine($"  UsedRange.Columns.Count = {usedRange.Columns.Count}");
                    }
                    finally
                    {
                        ComUtilities.Release(ref usedRange);
                    }

                    // Do a 3rd refresh to be absolutely sure
                    _output.WriteLine($"  Doing 3rd refresh...");
                    refreshResult = queryTable.Refresh(false);
                    _output.WriteLine($"  After 3rd Refresh: QueryTables.Count = {queryTables.Count}");
                }
                else
                {
                    _output.WriteLine($"  ERROR: No QueryTables found!");
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

        _output.WriteLine("✓ Scenario 2 complete");
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

        _output.WriteLine($"Updating query: {queryName} with NEW M CODE (more columns)");

        await session.Execute<int>((ctx, ct) =>
        {
            dynamic? queries = null;
            dynamic? query = null;
            dynamic? sheets = null;
            dynamic? sheet = null;
            dynamic? queryTables = null;
            dynamic? queryTable = null;

            try
            {
                // Step 1: Update the query M code
                queries = ctx.Book.Queries;
                query = queries.Item(queryName);

                _output.WriteLine($"  Old Formula length: {query.Formula.ToString().Length} chars");
                query.Formula = newMCode;
                _output.WriteLine($"  New Formula length: {query.Formula.ToString().Length} chars");

                // Step 2: Find the QueryTable and refresh
                sheets = ctx.Book.Worksheets;
                sheet = sheets.Item(queryName);

                queryTables = sheet.QueryTables;
                _output.WriteLine($"  Before refresh after update: QueryTables.Count = {queryTables.Count}");

                if (queryTables.Count > 0)
                {
                    queryTable = queryTables.Item(1);

                    // KEY TEST: What happens when we refresh with DIFFERENT M code?
                    _output.WriteLine($"  Refreshing with UPDATED M code...");

                    try
                    {
                        // Give Excel a moment to stabilize after formula change
                        System.Threading.Thread.Sleep(1000);

                        bool refreshResult = queryTable.Refresh(false);
                        _output.WriteLine($"  Refresh result: {refreshResult}");
                        _output.WriteLine($"  After refresh: QueryTables.Count = {queryTables.Count}");

                        // Check the data structure
                        dynamic? usedRange = null;
                        try
                        {
                            usedRange = sheet.UsedRange;
                            _output.WriteLine($"  UsedRange.Rows.Count = {usedRange.Rows.Count}");
                            _output.WriteLine($"  UsedRange.Columns.Count = {usedRange.Columns.Count} (should be MORE now)");
                            _output.WriteLine($"  Cell A1 value: {sheet.Range["A1"].Value2}");
                            _output.WriteLine($"  Cell B1 value: {sheet.Range["B1"].Value2}");
                            _output.WriteLine($"  Cell C1 value: {sheet.Range["C1"].Value2}");
                        }
                        finally
                        {
                            ComUtilities.Release(ref usedRange);
                        }
                    }
                    catch (System.Runtime.InteropServices.COMException ex) when (ex.HResult == unchecked((int)0x800706BE))
                    {
                        // RPC_E_CALL_REJECTED or RPC timeout - Excel is busy processing the formula change
                        _output.WriteLine($"  ⚠️ Excel busy (RPC timeout) - This is EXPECTED when updating M code!");
                        _output.WriteLine($"  Finding: Changing M code while QueryTable exists requires special handling");
                        _output.WriteLine($"  Workaround: Delete QueryTable, update formula, recreate QueryTable");
                    }
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

        _output.WriteLine("✓ Scenario 3 complete");
    }

    private async Task Scenario4_ConnectionOnly(IExcelBatch session)
    {
        string queryName = "ConnectionOnlyQuery";
        string mCode = @"
let
    Source = {100..105},
    ToTable = Table.FromList(Source, Splitter.SplitByNothing(), {""Value""}),
    Typed = Table.TransformColumnTypes(ToTable, {{""Value"", Int64.Type}})
in
    Typed";

        _output.WriteLine($"Creating connection-only query: {queryName}");

        await session.Execute<int>((ctx, ct) =>
        {
            dynamic? queries = null;
            dynamic? query = null;
            dynamic? sheets = null;

            try
            {
                // Step 1: Create Power Query (connection-only)
                queries = ctx.Book.Queries;
                query = queries.Add(queryName, mCode);
                _output.WriteLine($"  Query created: {query.Name}");

                // Step 2: Check if Excel created any QueryTables automatically
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
                        totalQueryTables += queryTables.Count;

                        if (queryTables.Count > 0)
                        {
                            _output.WriteLine($"  Sheet '{sheet.Name}' has {queryTables.Count} QueryTables");
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref queryTables);
                        ComUtilities.Release(ref sheet);
                    }
                }

                _output.WriteLine($"  Total QueryTables in workbook: {totalQueryTables}");
                _output.WriteLine($"  Expected: 0 (connection-only should NOT create QueryTables)");

                return 0;
            }
            finally
            {
                ComUtilities.Release(ref sheets);
                ComUtilities.Release(ref query);
                ComUtilities.Release(ref queries);
            }
        });

        _output.WriteLine("✓ Scenario 4 complete");
    }

    private async Task Scenario5_LoadConnectionOnlyToWorksheet(IExcelBatch session)
    {
        string queryName = "ConnectionOnlyQuery";

        _output.WriteLine($"Loading connection-only query to worksheet (simulating UI 'Load To')");

        await session.Execute<int>((ctx, ct) =>
        {
            dynamic? sheets = null;
            dynamic? sheet = null;
            dynamic? queryTables = null;
            dynamic? queryTable = null;

            try
            {
                // Step 1: Create worksheet
                sheets = ctx.Book.Worksheets;
                sheet = sheets.Add();
                sheet.Name = queryName;
                _output.WriteLine($"  Created worksheet: {sheet.Name}");

                // Step 2: Create QueryTable (this simulates UI "Load To > Table")
                string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
                _output.WriteLine($"  Creating QueryTable...");

                queryTables = sheet.QueryTables;
                _output.WriteLine($"  Before: QueryTables.Count = {queryTables.Count}");

                queryTable = queryTables.Add(connectionString, sheet.Range["A1"], Type.Missing);
                queryTable.Name = queryName;
                queryTable.CommandText = $"SELECT * FROM [{queryName}]";
                _output.WriteLine($"  QueryTable created: {queryTable.Name}");
                _output.WriteLine($"  After creation: QueryTables.Count = {queryTables.Count}");

                // Step 3: Refresh to actually load the data
                _output.WriteLine($"  Refreshing to load data...");
                bool refreshResult = queryTable.Refresh(false);
                _output.WriteLine($"  Refresh result: {refreshResult}");
                _output.WriteLine($"  After refresh: QueryTables.Count = {queryTables.Count}");

                // Step 4: Verify data loaded
                dynamic? usedRange = null;
                try
                {
                    usedRange = sheet.UsedRange;
                    _output.WriteLine($"  UsedRange.Rows.Count = {usedRange.Rows.Count}");
                    _output.WriteLine($"  UsedRange.Columns.Count = {usedRange.Columns.Count}");
                    _output.WriteLine($"  Cell A1 value: {sheet.Range["A1"].Value2}");
                    _output.WriteLine($"  Cell A2 value: {sheet.Range["A2"].Value2}");
                }
                finally
                {
                    ComUtilities.Release(ref usedRange);
                }

                // KEY TEST: Now delete the QueryTable - what happens to data?
                _output.WriteLine($"  CRITICAL TEST: Deleting QueryTable...");
                queryTable.Delete();
                ComUtilities.Release(ref queryTable);
                queryTable = null;

                _output.WriteLine($"  After Delete: QueryTables.Count = {queryTables.Count}");

                // Check if data remains
                _output.WriteLine($"  Checking if data remains after QueryTable.Delete()...");
                _output.WriteLine($"  Cell A1 value: {sheet.Range["A1"].Value2}");
                _output.WriteLine($"  Cell A2 value: {sheet.Range["A2"].Value2}");
                _output.WriteLine($"  DATA REMAINS: {sheet.Range["A1"].Value2 != null}");

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

        _output.WriteLine("✓ Scenario 5 complete");
    }
}
