using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Table DAX operations (create-from-dax, update-dax, get-dax)
/// </summary>
public partial class TableCommands
{
    // Excel constants for DAX operations
    private const int xlSrcModel = 4;  // PowerPivot Data Model source type
    private const int xlCmdDAX = 8;    // DAX command type

    /// <summary>
    /// Finds the WorkbookConnection for a table by trying the TableObject path first,
    /// then falling back to the QueryTable path. Returns null if neither works.
    /// Caller is responsible for releasing the out parameters.
    /// </summary>
    private static Excel.WorkbookConnection? FindTableWorkbookConnection(
        Excel.ListObject table,
        out Excel.TableObject? tableObject,
        out Excel.QueryTable? queryTable)
    {
        queryTable = null;
        tableObject = null;
        int sourceType = (int)table.SourceType;

        // Try the TableObject path first (for xlSrcModel tables created with ListObjects.Add)
        if (sourceType == xlSrcModel)
        {
            tableObject = table.TableObject;
            if (tableObject != null)
            {
                Excel.WorkbookConnection? conn = tableObject.WorkbookConnection;
                if (conn != null) return conn;
            }
        }

        // Fall back to QueryTable path (for QueryTables.Add based tables)
        if (sourceType is 0 or 3)
        {
            queryTable = table.QueryTable;
            if (queryTable != null)
            {
                Excel.WorkbookConnection? conn = queryTable.WorkbookConnection;
                if (conn != null) return conn;
            }
        }

        return null;
    }

    /// <inheritdoc />
    public OperationResult CreateFromDax(IExcelBatch batch, string sheetName, string tableName, string daxQuery, string? targetCell = null)
    {
        // Validate parameters
        if (string.IsNullOrWhiteSpace(sheetName))
        {
            throw new ArgumentException("sheetName is required for create-from-dax action", nameof(sheetName));
        }

        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ArgumentException("tableName is required for create-from-dax action", nameof(tableName));
        }
        if (string.IsNullOrWhiteSpace(daxQuery))
        {
            throw new ArgumentException("daxQuery is required for create-from-dax action", nameof(daxQuery));
        }

        // Default target cell
        targetCell ??= "A1";

        return batch.Execute((ctx, ct) =>
        {
            Excel.Model? model = null;
            Excel.WorkbookConnection? modelWbConn = null;
            Excel.ModelConnection? modelConnection = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? destRange = null;
            Excel.ListObjects? listObjects = null;
            Excel.ListObject? listObject = null;

            try
            {
                // Check if table name already exists
                if (TableExists(ctx.Book, tableName))
                {
                    throw new InvalidOperationException($"Table '{tableName}' already exists");
                }

                // Get the sheet
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found");
                }

                // Check if workbook has Data Model and get first table name
                // CreateModelWorkbookConnection requires a ModelTable name to create the connection
                model = ctx.Book.Model;
                Excel.ModelTables? modelTables = null;
                string? baseModelTableName = null;
                try
                {
                    modelTables = model.ModelTables;
                    if (modelTables is null || modelTables.Count == 0)
                    {
                        throw new InvalidOperationException("Workbook has no Data Model tables. Add data to the Data Model first using powerquery or table add-to-datamodel.");
                    }

                    // Get the first ModelTable name to use as base for connection
                    Excel.ModelTable? firstTable = modelTables!.Item(1);
                    try
                    {
                        baseModelTableName = firstTable.Name?.ToString();
                    }
                    finally
                    {
                        ComUtilities.Release(ref firstTable);
                    }
                }
                finally
                {
                    ComUtilities.Release(ref modelTables);
                }

                if (string.IsNullOrEmpty(baseModelTableName))
                {
                    throw new InvalidOperationException("Could not get table name from Data Model.");
                }

                // Create a model workbook connection using an existing ModelTable name
                // This creates a connection that we can then configure for DAX queries
                modelWbConn = model.CreateModelWorkbookConnection(baseModelTableName);
                modelConnection = modelWbConn.ModelConnection;

                // Configure the connection for DAX EVALUATE query
                modelConnection.CommandType = Excel.XlCmdType.xlCmdDAX;
                modelConnection.CommandText = daxQuery;

                // Keep synchronous DAX refresh cancellable without blocking inbound COM callbacks.
                OleMessageFilter.SetPendingCancellationToken(ct);
                try
                {
                    modelWbConn.Refresh();
                }
                finally
                {
                    OleMessageFilter.ClearPendingCancellationToken();
                }
                ComUtilities.Release(ref modelConnection);

                // Get target range for the table
                destRange = sheet.Range[targetCell];

                // Create Excel Table (ListObject) backed by the DAX query
                listObjects = sheet.ListObjects;
                listObject = listObjects.Add(
                    Excel.XlListObjectSourceType.xlSrcModel,
                    modelWbConn,    // The ModelWorkbookConnection with DAX
                    Type.Missing,   // LinkSource (not used)
                    Excel.XlYesNoGuess.xlYes,
                    destRange       // Target range
                );

                // Set the table name
                listObject.Name = tableName;

                OleMessageFilter.SetPendingCancellationToken(ct);
                try
                {
                    // Refresh the table to populate data.
                    listObject.Refresh();
                }
                finally
                {
                    OleMessageFilter.ClearPendingCancellationToken();
                }

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref listObject);
                ComUtilities.Release(ref listObjects);
                ComUtilities.Release(ref destRange);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref modelConnection);
                ComUtilities.Release(ref modelWbConn);
                ComUtilities.Release(ref model);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult UpdateDax(IExcelBatch batch, string tableName, string daxQuery)
    {
        // Validate parameters
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ArgumentException("tableName is required for update-dax action", nameof(tableName));
        }
        if (string.IsNullOrWhiteSpace(daxQuery))
        {
            throw new ArgumentException("daxQuery is required for update-dax action", nameof(daxQuery));
        }

        return batch.Execute((ctx, ct) =>
        {
            Excel.ListObject? table = null;
            Excel.TableObject? tableObject = null;
            Excel.QueryTable? queryTable = null;
            Excel.WorkbookConnection? workbookConnection = null;
            Excel.ModelConnection? modelConnection = null;

            try
            {
                // Find the table
                table = FindTable(ctx.Book, tableName);

                workbookConnection = FindTableWorkbookConnection(table, out tableObject, out queryTable);

                if (workbookConnection == null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' is not connected to a data source. Only DAX-backed tables can be updated.");
                }

                // Get the model connection
                try
                {
                    modelConnection = workbookConnection.ModelConnection;
                }
                catch (COMException)
                {
                    throw new InvalidOperationException($"Table '{tableName}' does not have a ModelConnection. Use update-dax only with DAX-backed tables.");
                }

                if (modelConnection == null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' is not backed by a Model connection. Use update-dax only with DAX-backed tables.");
                }

                // Check if current command type is DAX
                int currentCmdType = (int)modelConnection.CommandType;
                if (currentCmdType != xlCmdDAX)
                {
                    throw new InvalidOperationException($"Table '{tableName}' has command type {currentCmdType}, not xlCmdDAX (8). Use update-dax only with DAX-backed tables.");
                }

                // Update the DAX query
                modelConnection.CommandText = daxQuery;

                OleMessageFilter.SetPendingCancellationToken(ct);
                try
                {
                    // Refresh to execute the new query.
                    workbookConnection.Refresh();
                }
                finally
                {
                    OleMessageFilter.ClearPendingCancellationToken();
                }

                OleMessageFilter.SetPendingCancellationToken(ct);
                try
                {
                    table.Refresh();
                }
                finally
                {
                    OleMessageFilter.ClearPendingCancellationToken();
                }

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref modelConnection);
                ComUtilities.Release(ref workbookConnection);
                ComUtilities.Release(ref queryTable);
                ComUtilities.Release(ref tableObject);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
    public TableDaxInfoResult GetDax(IExcelBatch batch, string tableName)
    {
        // Validate parameters
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ArgumentException("tableName is required for get-dax action", nameof(tableName));
        }
        var result = new TableDaxInfoResult
        {
            FilePath = batch.WorkbookPath,
            TableName = tableName
        };

        return batch.Execute((ctx, ct) =>
        {
            Excel.ListObject? table = null;
            Excel.TableObject? tableObject = null;
            Excel.QueryTable? queryTable = null;
            Excel.WorkbookConnection? workbookConnection = null;
            Excel.ModelConnection? modelConnection = null;

            try
            {
                // Find the table
                table = FindTable(ctx.Book, tableName);

                workbookConnection = FindTableWorkbookConnection(table, out tableObject, out queryTable);

                if (workbookConnection == null)
                {
                    result.HasDaxConnection = false;
                    result.Success = true;
                    return result;
                }

                // ModelConnection is only defined for the Data Model connection type.
                if ((int)workbookConnection.Type != 7)
                {
                    result.HasDaxConnection = false;
                    result.Success = true;
                    return result;
                }
                modelConnection = workbookConnection.ModelConnection;

                if (modelConnection == null)
                {
                    result.HasDaxConnection = false;
                    result.Success = true;
                    return result;
                }

                // Check if command type is DAX
                int cmdType = (int)modelConnection.CommandType;
                if (cmdType == xlCmdDAX)
                {
                    result.HasDaxConnection = true;
                    result.DaxQuery = ((object?)modelConnection.CommandText)?.ToString();
                    result.ModelConnectionName = workbookConnection.Name?.ToString();
                }
                else
                {
                    result.HasDaxConnection = false;
                }

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref modelConnection);
                ComUtilities.Release(ref workbookConnection);
                ComUtilities.Release(ref queryTable);
                ComUtilities.Release(ref tableObject);
                ComUtilities.Release(ref table);
            }
        });
    }
}
