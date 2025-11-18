using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

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
        var result = new PowerQueryLoadResult
        {
            FilePath = batch.WorkbookPath,
            QueryName = queryName,
            LoadDestination = loadMode,
            WorksheetName = targetSheet,
            TargetCellAddress = targetCellAddress
        };

        // Validate inputs
        bool requiresWorksheet = loadMode == PowerQueryLoadMode.LoadToTable || loadMode == PowerQueryLoadMode.LoadToBoth;

        if (requiresWorksheet && string.IsNullOrWhiteSpace(targetSheet))
        {
            targetSheet = queryName; // Default to query name
            result.WorksheetName = targetSheet;
        }

        if (!string.IsNullOrWhiteSpace(targetCellAddress) && !requiresWorksheet)
        {
            result.Success = false;
            result.ErrorMessage = "targetCellAddress is only supported when loadMode is 'LoadToTable' or 'LoadToBoth'.";
            return result;
        }

        targetCellAddress ??= "A1"; // Default cell address

        try
        {
            return batch.Execute((ctx, ct) =>
            {
                dynamic? queries = null;
                dynamic? query = null;

                try
                {
                    // STEP 1: Find the Power Query
                    queries = ctx.Book.Queries;
                    query = null;
                    for (int i = 1; i <= queries.Count; i++)
                    {
                        dynamic? q = null;
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
                            if (q != null) ComUtilities.Release(ref q!);
                        }
                    }

                    if (query == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Query '{queryName}' not found";
                        return result;
                    }

                    // STEP 2: Apply load destination based on mode
                    switch (loadMode)
                    {
                        case PowerQueryLoadMode.LoadToTable:
                            LoadQueryToWorksheet(ctx.Book, queryName, targetSheet!, targetCellAddress, result);
                            break;

                        case PowerQueryLoadMode.LoadToDataModel:
                            LoadQueryToDataModel(ctx.Book, queryName, result);
                            break;

                        case PowerQueryLoadMode.LoadToBoth:
                            // Load to worksheet first
                            LoadQueryToWorksheet(ctx.Book, queryName, targetSheet!, targetCellAddress, result);

                            if (result.Success)
                            {
                                // Preserve worksheet properties before loading to Data Model
                                int worksheetRows = result.RowsLoaded;
                                string? worksheetCell = result.TargetCellAddress;

                                // Then also load to Data Model
                                LoadQueryToDataModel(ctx.Book, queryName, result);

                                // Restore worksheet properties (Data Model sets them to null/-1)
                                if (result.Success)
                                {
                                    result.RowsLoaded = worksheetRows;
                                    result.TargetCellAddress = worksheetCell;
                                }
                            }
                            break;

                        case PowerQueryLoadMode.ConnectionOnly:
                            // No loading needed - query already exists as connection-only
                            result.ConfigurationApplied = true;
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

                    return result;
                }
                finally
                {
                    if (query != null) ComUtilities.Release(ref query!);
                    if (queries != null) ComUtilities.Release(ref queries!);
                }
            });
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Unexpected error: {ex.Message}";
            return result;
        }
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
    private static void LoadQueryToWorksheet(
        dynamic workbook,
        string queryName,
        string sheetName,
        string targetCellAddress,
        dynamic result)
    {
        dynamic? worksheets = null;
        dynamic? sheet = null;
        dynamic? destination = null;
        dynamic? listObjects = null;
        dynamic? listObject = null;
        dynamic? queryTable = null;

        try
        {
            worksheets = workbook.Worksheets;

            // Check if worksheet exists (Excel UI behavior: validate occupied cells on existing sheets)
            bool worksheetExists = false;
            for (int i = 1; i <= worksheets.Count; i++)
            {
                dynamic? ws = null;
                try
                {
                    ws = worksheets.Item(i);
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
                    if (ws != null) ComUtilities.Release(ref ws!);
                }
            }

            // Create new worksheet if doesn't exist
            if (!worksheetExists)
            {
                sheet = worksheets.Add();
                sheet.Name = sheetName;
            }

            if (sheet == null)
            {
                result.Success = false;
                result.ErrorMessage = $"Cannot access worksheet '{sheetName}'";
                return;
            }

            // Get destination range
            destination = sheet.Range[targetCellAddress];

            // For existing worksheets, check if target area would overlap with existing tables
            // Excel allows loading over cell data, but NOT over existing tables/PivotTables
            if (worksheetExists)
            {
                // Check if any ListObjects (tables) would overlap with this location
                // Excel error: "A table cannot overlap a range that contains a pivot table report, query results, protected cells or another table."
                dynamic? existingTables = null;
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
                            dynamic? table = null;
                            dynamic? tableRange = null;
                            try
                            {
                                table = existingTables.Item(i);
                                tableRange = table.Range;

                                int tableStartRow = Convert.ToInt32(tableRange.Row);
                                int tableStartCol = Convert.ToInt32(tableRange.Column);
                                int tableEndRow = tableStartRow + Convert.ToInt32(tableRange.Rows.Count) - 1;
                                int tableEndCol = tableStartCol + Convert.ToInt32(tableRange.Columns.Count) - 1;

                                // Check if destination cell would overlap with existing table
                                if (destRow >= tableStartRow && destRow <= tableEndRow &&
                                    destCol >= tableStartCol && destCol <= tableEndCol)
                                {
                                    result.Success = false;
                                    result.ErrorMessage = $"Cell {targetCellAddress} on sheet '{sheetName}' overlaps with existing table.";
                                    return;
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
                dynamic? cellValue = destination.Value2;
                bool cellHasData = cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString());
                ComUtilities.Release(ref cellValue!);

                if (cellHasData)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Target cell '{targetCellAddress}' on worksheet '{sheetName}' already contains data. Choose a different targetCellAddress or clear the existing data first.";
                    return;
                }
            }

            // Build OLE DB connection string for Power Query
            string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName};Extended Properties=\"\"";

            // Add ListObject (Excel Table) with external data source
            // This is the CORRECT way to load Power Query to worksheet
            listObjects = sheet.ListObjects;
            listObject = listObjects.Add(
                0,                  // SourceType: 0 = xlSrcExternal
                connectionString,   // Source: connection string
                Type.Missing,       // LinkSource
                1,                  // XlListObjectHasHeaders: xlYes
                destination         // Destination: starting cell
            );

            // Configure the QueryTable behind the ListObject
            queryTable = listObject.QueryTable;
            queryTable.CommandType = 2; // xlCmdSql
            queryTable.CommandText = $"SELECT * FROM [{queryName}]";
            queryTable.AdjustColumnWidth = true;
            queryTable.PreserveFormatting = true;
            queryTable.BackgroundQuery = false; // Synchronous
            queryTable.RefreshStyle = 1; // xlInsertDeleteCells
            queryTable.PreserveColumnInfo = false; // Allow schema changes on refresh

            // Refresh to materialize the table
            queryTable.Refresh(false); // Synchronous refresh

            // Capture results - use ListObject Range for total rows, subtract header
            dynamic? listObjectRange = listObject.Range;
            int totalRows = listObjectRange != null ? Convert.ToInt32(listObjectRange.Rows.Count) : 0;
            result.TargetCellAddress = targetCellAddress;
            result.RowsLoaded = totalRows > 0 ? totalRows - 1 : 0; // Subtract header row
            result.Success = true;

            ComUtilities.Release(ref listObjectRange!);
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error loading to worksheet: {ex.Message}";
        }
        finally
        {
            ComUtilities.Release(ref queryTable);
            ComUtilities.Release(ref listObject);
            ComUtilities.Release(ref listObjects);
            ComUtilities.Release(ref destination);
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref worksheets);
        }
    }

    /// <summary>
    /// Loads query data to the Data Model using Connections.Add2.
    /// SHARED IMPLEMENTATION - Used by both Create and LoadTo.
    /// </summary>
    private static void LoadQueryToDataModel(
        dynamic workbook,
        string queryName,
        dynamic result)
    {
        dynamic? connections = null;
        dynamic? connection = null;

        try
        {
            connections = workbook.Connections;

            string connectionName = $"Query - {queryName}";
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

            result.RowsLoaded = -1; // Data Model doesn't expose row count
            result.TargetCellAddress = null;
            result.Success = true;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error loading to Data Model: {ex.Message}";
        }
        finally
        {
            ComUtilities.Release(ref connection);
            ComUtilities.Release(ref connections);
        }
    }

    /// <summary>
    /// Gets an existing worksheet or creates a new one.
    /// SHARED HELPER - Used by Create and LoadTo.
    /// </summary>
    private static dynamic? GetOrCreateWorksheet(dynamic worksheets, string sheetName)
    {
        dynamic? sheet = null;

        try
        {
            // Try to find existing worksheet
            int count = worksheets.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic? candidate = null;
                try
                {
                    candidate = worksheets.Item(i);
                    string name = candidate.Name ?? "";

                    if (name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        sheet = candidate;
                        candidate = null; // Prevent release in finally
                        return sheet;
                    }
                }
                finally
                {
                    if (candidate != null)
                    {
                        ComUtilities.Release(ref candidate);
                    }
                }
            }

            // Sheet not found, create new one
            sheet = worksheets.Add();
            sheet.Name = sheetName;
            return sheet;
        }
        catch
        {
            // Error accessing or creating worksheet
            if (sheet != null)
            {
                ComUtilities.Release(ref sheet);
            }
            return null;
        }
    }
}
