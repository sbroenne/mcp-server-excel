using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Table DAX operations (create-from-dax, update-dax, get-dax)
/// </summary>
public partial class TableCommands
{
    // Excel constants for DAX operations
    private const int xlSrcModel = 4;  // PowerPivot Data Model source type
    private const int xlCmdDAX = 8;    // DAX command type
    // xlYes is defined in TableCommands.Sort.cs

    /// <summary>
    /// Finds the WorkbookConnection for a table by trying the TableObject path first,
    /// then falling back to the QueryTable path. Returns null if neither works.
    /// Caller is responsible for releasing the out parameters.
    /// </summary>
    private static dynamic? FindTableWorkbookConnection(
        dynamic table,
        out dynamic? tableObject,
        out dynamic? queryTable)
    {
        queryTable = null;

        // Try the TableObject path first (for xlSrcModel tables created with ListObjects.Add)
        try
        {
            tableObject = table.TableObject;
            if (tableObject != null)
            {
                dynamic? conn = tableObject.WorkbookConnection;
                if (conn != null) return conn;
            }
        }
        catch (COMException)
        {
            tableObject = null;
        }

        // Fall back to QueryTable path (for QueryTables.Add based tables)
        try
        {
            queryTable = table.QueryTable;
            if (queryTable != null)
            {
                dynamic? conn = queryTable.WorkbookConnection;
                if (conn != null) return conn;
            }
        }
        catch (COMException)
        {
            queryTable = null;
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

        ValidateTableName(tableName);

        if (string.IsNullOrWhiteSpace(daxQuery))
        {
            throw new ArgumentException("daxQuery is required for create-from-dax action", nameof(daxQuery));
        }

        // Default target cell
        targetCell ??= "A1";

        return batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            dynamic? modelWbConn = null;
            dynamic? modelConnection = null;
            dynamic? sheet = null;
            dynamic? destRange = null;
            dynamic? listObjects = null;
            dynamic? listObject = null;

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
                dynamic? modelTables = null;
                string? baseModelTableName = null;
                try
                {
                    modelTables = model.ModelTables;
                    if (modelTables == null || modelTables.Count == 0)
                    {
                        throw new InvalidOperationException("Workbook has no Data Model tables. Add data to the Data Model first using powerquery or table add-to-datamodel.");
                    }

                    // Get the first ModelTable name to use as base for connection
                    dynamic? firstTable = modelTables.Item(1);
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
                modelConnection.CommandType = xlCmdDAX;  // 8 = xlCmdDAX
                modelConnection.CommandText = daxQuery;

                // Refresh to execute the DAX query
                modelWbConn.Refresh();

                // Get target range for the table
                destRange = sheet.Range[targetCell];

                // Create Excel Table (ListObject) backed by the DAX query
                listObjects = sheet.ListObjects;
                listObject = listObjects.Add(
                    xlSrcModel,     // Source type: PowerPivot Data Model
                    modelWbConn,    // The ModelWorkbookConnection with DAX
                    Type.Missing,   // LinkSource (not used)
                    xlYes,          // HasHeaders: Yes
                    destRange       // Target range
                );

                // Set the table name
                listObject.Name = tableName;

                // Refresh the table to populate data
                listObject.Refresh();

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

        ValidateTableName(tableName);

        if (string.IsNullOrWhiteSpace(daxQuery))
        {
            throw new ArgumentException("daxQuery is required for update-dax action", nameof(daxQuery));
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? tableObject = null;
            dynamic? queryTable = null;
            dynamic? workbookConnection = null;
            dynamic? modelConnection = null;

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
                int currentCmdType = Convert.ToInt32(modelConnection.CommandType);
                if (currentCmdType != xlCmdDAX)
                {
                    throw new InvalidOperationException($"Table '{tableName}' has command type {currentCmdType}, not xlCmdDAX (8). Use update-dax only with DAX-backed tables.");
                }

                // Update the DAX query
                modelConnection.CommandText = daxQuery;

                // Refresh to execute the new query
                workbookConnection.Refresh();
                table.Refresh();

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

        ValidateTableName(tableName);

        var result = new TableDaxInfoResult
        {
            FilePath = batch.WorkbookPath,
            TableName = tableName
        };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? tableObject = null;
            dynamic? queryTable = null;
            dynamic? workbookConnection = null;
            dynamic? modelConnection = null;

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

                // Try to get the model connection
                try
                {
                    modelConnection = workbookConnection.ModelConnection;
                }
                catch (COMException)
                {
                    result.HasDaxConnection = false;
                    result.Success = true;
                    return result;
                }

                if (modelConnection == null)
                {
                    result.HasDaxConnection = false;
                    result.Success = true;
                    return result;
                }

                // Check if command type is DAX
                int cmdType = Convert.ToInt32(modelConnection.CommandType);
                if (cmdType == xlCmdDAX)
                {
                    result.HasDaxConnection = true;
                    result.DaxQuery = modelConnection.CommandText?.ToString();
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


