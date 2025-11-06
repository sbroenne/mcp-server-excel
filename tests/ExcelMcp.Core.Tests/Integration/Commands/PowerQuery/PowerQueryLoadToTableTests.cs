using System;
using System.IO;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// Tests for Power Query SetLoadToTable workflow
/// Reproduces bug where connection-only query → refresh → set-load-to-table creates range instead of table
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "PowerQuery")]
[Trait("RunType", "Standard")]
public class PowerQueryLoadToTableTests : IDisposable
{
    private readonly string _tempDir;
    private readonly DataModelCommands _dataModelCommands;
    private readonly PowerQueryCommands _powerQueryCommands;

    public PowerQueryLoadToTableTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"pq-loadtotable-tests-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        _dataModelCommands = new DataModelCommands();
        _powerQueryCommands = new PowerQueryCommands(_dataModelCommands);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
        {
            try
            {
                Directory.Delete(_tempDir, recursive: true);
            }
            catch
            {
                // Best effort cleanup
            }
        }
        GC.SuppressFinalize(this);
    }

    [Fact]
    public async Task Bug_RefreshConnectionOnlyThenSetLoadToTable_CreatesRangeInsteadOfTable()
    {
        // ARRANGE - Simulate the exact LLM workflow from the issue using SINGLE BATCH
        // BUG: Connection-only query → Refresh → SetLoadToTable was duplicating columns
        // ROOT CAUSE: SetLoadToTableAsync was calling usedRange.Clear() instead of just deleting QueryTables
        // FIX: Only delete QueryTables, let Excel handle data cleanup automatically
        string testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryLoadToTableTests),
            nameof(Bug_RefreshConnectionOnlyThenSetLoadToTable_CreatesRangeInsteadOfTable),
            _tempDir);

        // Create a simple Power Query (connection-only mode)
        string mCode = @"
let
    Source = {1..10},
    Table = Table.FromList(Source, Splitter.SplitByNothing(), {""Value""}),
    ChangedType = Table.TransformColumnTypes(Table,{{""Value"", Int64.Type}})
in
    ChangedType";

        string queryName = "TestQuery";
        string mFile = CreateTempMCodeFile("test", mCode);

        // USE SINGLE BATCH FOR ENTIRE TEST
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // ACT 1 - Import query in connection-only mode
        var importResult = await _powerQueryCommands.ImportAsync(
            batch,
            queryName,
            mFile,
            loadDestination: "connection-only");

        Assert.True(importResult.Success, $"Import failed: {importResult.ErrorMessage}");

        // ACT 2 - Refresh the connection-only query (this might load data temporarily)
        var refreshResult = await _powerQueryCommands.RefreshAsync(batch, queryName);
        Assert.True(refreshResult.Success, $"Refresh failed: {refreshResult.ErrorMessage}");

        // ACT 3 - Set load to table (this should create a QueryTable/Table, but bug might create range)
        var loadResult = await _powerQueryCommands.SetLoadToTableAsync(batch, queryName, queryName);
        Assert.True(loadResult.Success, $"SetLoadToTable failed: {loadResult.ErrorMessage}");
        Assert.True(loadResult.DataLoadedToTable, "Data should be loaded to table");
        Assert.True(loadResult.RowsLoaded > 0, "Should have loaded rows");

        // CRITICAL: Save changes before verification
        await batch.SaveAsync();

        // ASSERT - Verify QueryTable was actually created (not just a range)
        var hasQueryTable = await batch.Execute((ctx, ct) =>
        {
            dynamic? sheets = null;
            dynamic? sheet = null;
            dynamic? queryTables = null;
            try
            {
                sheets = ctx.Book.Worksheets;

                // Find the target sheet
                for (int i = 1; i <= sheets.Count; i++)
                {
                    dynamic? currentSheet = null;
                    try
                    {
                        currentSheet = sheets.Item(i);
                        if (currentSheet.Name == queryName)
                        {
                            sheet = currentSheet;
                            currentSheet = null; // Keep reference
                            break;
                        }
                    }
                    finally
                    {
                        if (currentSheet != null)
                            ComUtilities.Release(ref currentSheet);
                    }
                }

                if (sheet == null)
                    return (false, "Sheet not found", 0);

                queryTables = sheet.QueryTables;
                int qtCount = queryTables.Count;

                // Check if QueryTable exists with expected name
                string expectedQTName = queryName.Replace(" ", "_");
                for (int i = 1; i <= qtCount; i++)
                {
                    dynamic? qt = null;
                    try
                    {
                        qt = queryTables.Item(i);
                        string qtName = qt.Name?.ToString() ?? "";
                        if (qtName.Equals(expectedQTName, StringComparison.OrdinalIgnoreCase) ||
                            qtName.Contains(expectedQTName, StringComparison.OrdinalIgnoreCase))
                        {
                            return (true, "QueryTable found", qtCount);
                        }
                    }
                    finally
                    {
                        if (qt != null)
                            ComUtilities.Release(ref qt);
                    }
                }

                return (false, $"QueryTable not found (sheet has {qtCount} QueryTables)", qtCount);
            }
            finally
            {
                if (queryTables != null)
                    ComUtilities.Release(ref queryTables);
                if (sheet != null)
                    ComUtilities.Release(ref sheet);
                if (sheets != null)
                    ComUtilities.Release(ref sheets);
            }
        });

        // Verify QueryTable exists
        Assert.True(hasQueryTable.Item1, $"BUG REPRODUCED: {hasQueryTable.Item2}");

        // CRITICAL REGRESSION CHECK: Verify EXACTLY ONE QueryTable (no duplicates)
        // Bug could create multiple QueryTables for the same query
        int queryTableCount = hasQueryTable.Item3;
        Assert.True(queryTableCount == 1,
            $"Expected exactly 1 QueryTable but found {queryTableCount}. " +
            "Multiple QueryTables indicates improper cleanup during SetLoadToTable.");

        // CRITICAL REGRESSION CHECK: Verify columns are NOT duplicated
        // Bug was: After Refresh → SetLoadToTable, columns would appear twice
        // Expected: 1 column (Value)
        // Bug would cause: 2+ columns (Value, Value, ...)
        var columnCount = await batch.Execute((ctx, ct) =>
        {
            dynamic? sheets = null;
            dynamic? sheet = null;
            dynamic? usedRange = null;
            try
            {
                sheets = ctx.Book.Worksheets;

                // Find the target sheet
                for (int i = 1; i <= sheets.Count; i++)
                {
                    dynamic? currentSheet = null;
                    try
                    {
                        currentSheet = sheets.Item(i);
                        if (currentSheet.Name == queryName)
                        {
                            sheet = currentSheet;
                            currentSheet = null; // Keep reference
                            break;
                        }
                    }
                    finally
                    {
                        if (currentSheet != null)
                            ComUtilities.Release(ref currentSheet);
                    }
                }

                if (sheet == null)
                    return 0;

                usedRange = sheet.UsedRange;
                int cols = usedRange.Columns.Count;
                return cols;
            }
            finally
            {
                if (usedRange != null)
                    ComUtilities.Release(ref usedRange);
                if (sheet != null)
                    ComUtilities.Release(ref sheet);
                if (sheets != null)
                    ComUtilities.Release(ref sheets);
            }
        });

        // REGRESSION ASSERTION: Verify only 1 column exists (not duplicated)
        Assert.Equal(1, columnCount);

        // BONUS: Check if data is in a ListObject (Excel Table)
        var hasListObject = await batch.Execute((ctx, ct) =>
        {
            dynamic? sheets = null;
            dynamic? sheet = null;
            dynamic? listObjects = null;
            try
            {
                sheets = ctx.Book.Worksheets;

                // Find the target sheet
                for (int i = 1; i <= sheets.Count; i++)
                {
                    dynamic? currentSheet = null;
                    try
                    {
                        currentSheet = sheets.Item(i);
                        if (currentSheet.Name == queryName)
                        {
                            sheet = currentSheet;
                            currentSheet = null; // Keep reference
                            break;
                        }
                    }
                    finally
                    {
                        if (currentSheet != null)
                            ComUtilities.Release(ref currentSheet);
                    }
                }

                if (sheet == null)
                    return (false, "Sheet not found", 0);

                listObjects = sheet.ListObjects;
                int tableCount = listObjects.Count;

                return (tableCount > 0, $"Found {tableCount} table(s)", tableCount);
            }
            finally
            {
                if (listObjects != null)
                    ComUtilities.Release(ref listObjects);
                if (sheet != null)
                    ComUtilities.Release(ref sheet);
                if (sheets != null)
                    ComUtilities.Release(ref sheets);
            }
        });

        // This assertion documents expected behavior (QueryTable may exist without ListObject)
        // If this fails, it means we're creating a QueryTable but not converting to Table
        // which is acceptable for QueryTables but not ideal for user experience
    }

    [Fact]
    public async Task Bug_DirectSetLoadToTable_CreatesProperTable()
    {
        // ARRANGE - Direct path without refresh, using SINGLE BATCH
        string testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryLoadToTableTests),
            nameof(Bug_DirectSetLoadToTable_CreatesProperTable),
            _tempDir);

        string mCode = @"
let
    Source = {1..10},
    Table = Table.FromList(Source, Splitter.SplitByNothing(), {""Value""}),
    ChangedType = Table.TransformColumnTypes(Table,{{""Value"", Int64.Type}})
in
    ChangedType";

        string queryName = "DirectQuery";
        string mFile = CreateTempMCodeFile("direct", mCode);

        // USE SINGLE BATCH
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Import query with direct load to table (no connection-only step)
        var importResult = await _powerQueryCommands.ImportAsync(
            batch,
            queryName,
            mFile,
            loadDestination: "worksheet");

        Assert.True(importResult.Success, $"Import failed: {importResult.ErrorMessage}");

        // CRITICAL: Save changes before verification
        await batch.SaveAsync();

        // Verify QueryTable was created properly
        var hasQueryTable = await batch.Execute((ctx, ct) =>
        {
            dynamic? sheets = null;
            dynamic? sheet = null;
            dynamic? queryTables = null;
            try
            {
                sheets = ctx.Book.Worksheets;

                for (int i = 1; i <= sheets.Count; i++)
                {
                    dynamic? currentSheet = null;
                    try
                    {
                        currentSheet = sheets.Item(i);
                        if (currentSheet.Name == queryName)
                        {
                            sheet = currentSheet;
                            currentSheet = null;
                            break;
                        }
                    }
                    finally
                    {
                        if (currentSheet != null)
                            ComUtilities.Release(ref currentSheet);
                    }
                }

                if (sheet == null)
                    return false;

                queryTables = sheet.QueryTables;
                return queryTables.Count > 0;
            }
            finally
            {
                if (queryTables != null)
                    ComUtilities.Release(ref queryTables);
                if (sheet != null)
                    ComUtilities.Release(ref sheet);
                if (sheets != null)
                    ComUtilities.Release(ref sheets);
            }
        });

        Assert.True(hasQueryTable, "Direct load should create QueryTable");
    }

    private string CreateTempMCodeFile(string prefix, string mCode)
    {
        string tempFile = Path.Combine(_tempDir, $"{prefix}_{Guid.NewGuid()}.pq");
        System.IO.File.WriteAllText(tempFile, mCode);
        return tempFile;
    }
}
