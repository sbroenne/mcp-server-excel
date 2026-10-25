using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query LoadTo operations - STANDALONE implementation.
/// Uses ListObjects.Add() pattern (same as Create) for consistency.
/// Based on Microsoft WorkbookQuery and ListObject API.
/// </summary>
public partial class PowerQueryCommands
{
    /// <summary>
    /// Applies load destination to an existing Power Query.
    /// Uses ListObjects.Add() for worksheet loading (consistent with Create).
    /// </summary>
    /// <remarks>
    /// Microsoft Docs Reference:
    /// - ListObjects.Add method - Creates Excel Table with external data source
    /// - QueryTable properties - Configure refresh behavior and formatting
    /// - Connections.Add2 method - Load to Data Model with CreateModelConnection=true
    ///
    /// IMPORTANT: Uses ListObjects.Add() (not QueryTables.Add()) for worksheet loading.
    /// This is the CORRECT approach per Microsoft docs and matches Create() behavior.
    /// </remarks>
    public PowerQueryLoadResult LoadTo(
        IExcelBatch batch,
        string queryName,
        PowerQueryLoadMode loadMode,
        string? targetSheet = null,
        string? targetCellAddress = null)
    {
        // Validate inputs
        bool requiresWorksheet = loadMode == PowerQueryLoadMode.LoadToTable || loadMode == PowerQueryLoadMode.LoadToBoth;

        if (requiresWorksheet && string.IsNullOrWhiteSpace(targetSheet))
        {
            targetSheet = queryName; // Default to query name
        }

        if (!string.IsNullOrWhiteSpace(targetCellAddress) && !requiresWorksheet)
        {
            throw new ArgumentException("targetCellAddress is only supported when loadMode is 'LoadToTable' or 'LoadToBoth'.", nameof(targetCellAddress));
        }

        targetCellAddress ??= "A1"; // Default cell address

        var timeout = ComInteropConstants.DataOperationTimeout;

        var started = System.Diagnostics.Stopwatch.GetTimestamp();
        return ExcelBatchWait.Execute(batch, timeout, (ctx, ct) =>
        {
            Excel.Queries? queries = null;
            Excel.WorkbookQuery? query = null;
            bool destinationsChanging = false;
            var result = new PowerQueryLoadResult
            {
                FilePath = batch.WorkbookPath,
                QueryName = queryName,
                LoadDestination = loadMode,
                WorksheetName = requiresWorksheet ? targetSheet : null,
                TargetCellAddress = requiresWorksheet ? targetCellAddress : null
            };

            try
            {
                // STEP 1: Find the Power Query
                queries = ctx.Book.Queries;
                query = null;
                for (int i = 1; i <= queries.Count; i++)
                {
                    Excel.WorkbookQuery? q = null;
                    try
                    {
                        q = queries.Item(i);
                        string qName = q.Name?.ToString() ?? "";
                        if (qName.Equals(queryName, StringComparison.OrdinalIgnoreCase))
                        {
                            query = q;
                            q = null; // Don't release - keeping reference
                            break;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref q!);
                    }
                }

                if (query == null)
                {
                    throw new InvalidOperationException($"Query '{queryName}' not found.");
                }

                // STEP 2: Always clean up current destinations before applying the new load mode.
                // This ensures clean state transitions in all directions:
                // LoadToTable → DataModel (removes old ListObject + worksheet connection)
                // DataModel → LoadToTable (removes old Data Model connection)
                // LoadToBoth → ConnectionOnly (removes both destinations)
                // etc.
                destinationsChanging = true;
                UnloadFromDestinations(ctx.Book, queryName);

                // STEP 3: Apply load destination based on mode
                switch (loadMode)
                {
                    case PowerQueryLoadMode.LoadToTable:
                        LoadQueryToWorksheet(ctx.Book, queryName, targetSheet!, targetCellAddress, result, ct);
                        break;

                    case PowerQueryLoadMode.LoadToDataModel:
                        LoadQueryToDataModel(ctx.Book, queryName, result, cancellationToken: ct);
                        break;

                    case PowerQueryLoadMode.LoadToBoth:
                        // For LoadToBoth, create TWO separate properly-named connections:
                        // 1. Worksheet connection: "Query - {name}" (created by LoadQueryToWorksheet)
                        // 2. Data Model connection: "Query - {name} (Data Model)" (with suffix to avoid conflict)
                        LoadQueryToWorksheet(ctx.Book, queryName, targetSheet!, targetCellAddress, result, ct);
                        LoadQueryToDataModel(ctx.Book, queryName, result, " (Data Model)", ct);
                        break;

                    case PowerQueryLoadMode.ConnectionOnly:
                        // UnloadFromDestinations already called above; just set result properties.
                        result.DataRefreshed = false;
                        result.RowsLoaded = 0;
                        result.TargetCellAddress = null;
                        result.Success = true;
                        break;
                }

                // Set additional result properties
                if (result.Success)
                {
                    result.ConfigurationApplied = true;
                    result.DataRefreshed = (loadMode != PowerQueryLoadMode.ConnectionOnly);
                }

                if (loadMode != PowerQueryLoadMode.ConnectionOnly)
                    ReadLoadMetrics(ctx.Book, queryName, result, ct);
                result.DurationMs = System.Diagnostics.Stopwatch.GetElapsedTime(started).TotalMilliseconds;
                return result;
            }
            catch (Exception ex) when (destinationsChanging && !ct.IsCancellationRequested && CanReportMutationFailure(ex))
            {
                throw CreateMutationFailure(ex, queryName, queryCreated: false,
                    state => InspectMutationState(ctx.Book, state, ct), loadDestinationChanged: true);
            }
            finally
            {
                ComUtilities.Release(ref query!);
                ComUtilities.Release(ref queries!);
            }
        });
    }

    /// <summary>
    /// Loads query data to a worksheet using ListObjects.Add (correct approach for Power Query).
    /// SHARED IMPLEMENTATION - Used by both Create and LoadTo.
    /// </summary>
    /// <remarks>
    /// This is extracted from Create.cs for reuse. Both Create and LoadTo should use
    /// the same ListObjects.Add() pattern for consistency.
    /// Matches Excel UI behavior: Creates worksheet if it doesn't exist, or loads to existing worksheet.
    /// </remarks>
    private static bool LoadQueryToWorksheet(
        Excel.Workbook workbook,
        string queryName,
        string sheetName,
        string targetCellAddress,
        PowerQueryOperationResult result,
        CancellationToken cancellationToken)
    {
        Excel.Sheets? worksheets = null;
        Excel.Worksheet? sheet = null;
        Excel.Range? destination = null;
        Excel.Connections? connections = null;
        Excel.WorkbookConnection? connection = null;
        Excel.ListObjects? listObjects = null;
        Excel.ListObject? listObject = null;
        Excel.QueryTable? queryTable = null;
        Excel.Range? listObjectRange = null;

        try
        {
            worksheets = workbook.Worksheets;

            // Check if worksheet exists (Excel UI behavior: validate occupied cells on existing sheets)
            bool worksheetExists = false;
            for (int i = 1; i <= worksheets.Count; i++)
            {
                Excel.Worksheet? ws = null;
                try
                {
                    ws = (Excel.Worksheet)worksheets[i];
                    string wsName = ws.Name?.ToString() ?? "";
                    if (wsName.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        worksheetExists = true;
                        sheet = ws;
                        ws = null; // Keep reference, don't release
                        break;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref ws!);
                }
            }

            // Create new worksheet if doesn't exist
            if (!worksheetExists)
            {
                sheet = (Excel.Worksheet)worksheets.Add();
                sheet.Name = sheetName;
            }

            if (sheet == null)
            {
                throw new InvalidOperationException($"Cannot access worksheet '{sheetName}'");
            }

            // Get destination range
            destination = sheet.Range[targetCellAddress];

            // For existing worksheets, check if target area would overlap with existing tables
            // Excel allows loading over cell data, but NOT over existing tables/PivotTables
            if (worksheetExists)
            {
                // Check if any ListObjects (tables) would overlap with this location
                // Excel error: "A table cannot overlap a range that contains a pivot table report, query results, protected cells or another table."
                Excel.ListObjects? existingTables = null;
                try
                {
                    existingTables = sheet.ListObjects;
                    int tableCount = existingTables.Count;

                    if (tableCount > 0)
                    {
                        // Get destination cell row/column for comparison
                        int destRow = Convert.ToInt32(destination.Row);
                        int destCol = Convert.ToInt32(destination.Column);

                        for (int i = 1; i <= tableCount; i++)
                        {
                            Excel.ListObject? table = null;
                            Excel.Range? tableRange = null;
                            try
                            {
                                table = existingTables[i];
                                tableRange = table.Range;

                                int tableStartRow = Convert.ToInt32(tableRange.Row);
                                int tableStartCol = Convert.ToInt32(tableRange.Column);
                                (int tableRows, int tableColumns) = Range.RangeHelpers.GetDimensions(tableRange);
                                int tableEndRow = tableStartRow + tableRows - 1;
                                int tableEndCol = tableStartCol + tableColumns - 1;

                                // Check if destination cell would overlap with existing table
                                if (destRow >= tableStartRow && destRow <= tableEndRow &&
                                    destCol >= tableStartCol && destCol <= tableEndCol)
                                {
                                    throw new InvalidOperationException($"Cell {targetCellAddress} on sheet '{sheetName}' overlaps with existing table.");
                                }
                            }
                            finally
                            {
                                ComUtilities.Release(ref tableRange);
                                ComUtilities.Release(ref table);
                            }
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref existingTables);
                }

                // Also check if target cell contains data (Excel UI validation)
                object? cellValue = destination.Value2;
                bool cellHasData = cellValue is not null && !string.IsNullOrWhiteSpace(cellValue.ToString());
                ComUtilities.Release(ref cellValue!);

                if (cellHasData)
                {
                    throw new InvalidOperationException($"Target cell '{targetCellAddress}' on worksheet '{sheetName}' already contains data. Choose a different targetCellAddress or clear the existing data first.");
                }
            }

            // Step 1: Create connection with Connections.Add2() using proper naming
            // This ensures the connection is named "Query - {queryName}" instead of generic "Connection", "Connection1", etc.
            connections = workbook.Connections;
            string connectionName = $"Query - {queryName}";
            string connectionDescription = $"Connection to the '{queryName}' query in the workbook.";
            string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
            string commandText = $"SELECT * FROM [{queryName}]";

            connection = connections.Add2(
                Name: connectionName,
                Description: connectionDescription,
                ConnectionString: connectionString,
                CommandText: commandText,
                lCmdtype: 2,                    // xlCmdSql
                CreateModelConnection: false,   // Worksheet loading, NOT Data Model
                ImportRelationships: false
            );

            // Step 2: Add ListObject using the connection object (not connection string)
            // This reuses the properly-named connection instead of creating a new generic one
            listObjects = sheet.ListObjects;
            listObject = listObjects.Add(
                Excel.XlListObjectSourceType.xlSrcExternal,
                connection,         // Source: connection object (reuses existing named connection)
                Type.Missing,       // LinkSource
                Excel.XlYesNoGuess.xlYes,
                destination         // Destination: starting cell
            );

            // Configure the QueryTable behind the ListObject
            queryTable = listObject.QueryTable;
            queryTable.CommandType = Excel.XlCmdType.xlCmdSql;
            queryTable.CommandText = $"SELECT * FROM [{queryName}]";
            queryTable.AdjustColumnWidth = true;
            queryTable.PreserveFormatting = true;
            queryTable.BackgroundQuery = false; // Synchronous
            queryTable.RefreshStyle = Excel.XlCellInsertionMode.xlInsertDeleteCells;
            // Excel needs column bindings to remove fields when the source schema shrinks.
            queryTable.PreserveColumnInfo = true;

            // Refresh to materialize the table.
            // Do NOT use EnterLongOperation here: synchronous QueryTable refresh depends on inbound
            // Excel callbacks to complete. Rejecting those callbacks can deadlock the load.
            QueryTableRefreshHelper.RefreshSynchronously(
                queryTable,
                cancellationToken,
                $"Power Query load for '{queryName}'");

            // Name the table after the query for predictable M-code referencing.
            // Without this, Excel auto-assigns a generic name ("Table1", "Table2", etc.)
            // which makes Excel.CurrentWorkbook(){[Name=queryName]}[Content] lookups unreliable.
            string createdTableName = listObject.Name;
            try
            {
                listObject.Name = queryName;
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException(
                    $"Query '{queryName}' loaded into table '{createdTableName}' on worksheet '{sheetName}', " +
                    $"but renaming the table failed: {ex.Message}. The loaded table remains; no changes were rolled back.", ex);
            }

            // Capture results - use ListObject Range for total rows, subtract header
            listObjectRange = listObject.Range;
            int totalRows = listObjectRange is not null ? Range.RangeHelpers.GetDimensions(listObjectRange).Rows : 0;
            result.RowsLoaded = totalRows > 0 ? totalRows - 1 : 0; // Subtract header row
            result.Success = true;

            return true;
        }
        finally
        {
            ComUtilities.Release(ref listObjectRange);
            ComUtilities.Release(ref queryTable);
            ComUtilities.Release(ref listObject);
            ComUtilities.Release(ref listObjects);
            ComUtilities.Release(ref connection);
            ComUtilities.Release(ref connections);
            ComUtilities.Release(ref destination);
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref worksheets);
        }
    }

    /// <summary>
    /// Loads query data to the Data Model using Connections.Add2.
    /// SHARED IMPLEMENTATION - Used by both Create and LoadTo.
    /// </summary>
    /// <param name="workbook">The workbook to load into.</param>
    /// <param name="queryName">The Power Query name.</param>
    /// <param name="result">Result object to populate.</param>
    /// <param name="connectionNameSuffix">Optional suffix for connection name to avoid conflicts (e.g., " (Data Model)").</param>
    private static bool LoadQueryToDataModel(
        Excel.Workbook workbook,
        string queryName,
        PowerQueryOperationResult result,
        string? connectionNameSuffix = null,
        CancellationToken cancellationToken = default)
    {
        Excel.Connections? connections = null;
        Excel.WorkbookConnection? connection = null;

        try
        {
            connections = workbook.Connections;

            // Use suffix if provided (for LoadToBoth to avoid conflict with worksheet connection)
            string connectionName = string.IsNullOrEmpty(connectionNameSuffix)
                ? $"Query - {queryName}"
                : $"Query - {queryName}{connectionNameSuffix}";
            string description = $"Connection to the '{queryName}' query in the workbook.";
            string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
            string commandText = $"\"{queryName}\"";

            connection = connections.Add2(
                Name: connectionName,
                Description: description,
                ConnectionString: connectionString,
                CommandText: commandText,
                lCmdtype: 6, // Data Model command type
                CreateModelConnection: true, // CRITICAL: This loads to Data Model
                ImportRelationships: false
            );

            // Refresh the connection to actually load data into the Data Model.
            // Without this call, the connection is registered but no data is materialized —
            // the table never appears in the Data Model even though success is returned.
            // Do NOT use EnterLongOperation here: synchronous connection.Refresh() can require
            // inbound callbacks from Excel/MashupHost to complete.
            OleMessageFilter.SetPendingCancellationToken(cancellationToken);
            try
            {
                connection.Refresh();
            }
            finally
            {
                OleMessageFilter.ClearPendingCancellationToken();
            }

            result.Success = true;
            return true;
        }
        finally
        {
            ComUtilities.Release(ref connection);
            ComUtilities.Release(ref connections);
        }
    }

}
