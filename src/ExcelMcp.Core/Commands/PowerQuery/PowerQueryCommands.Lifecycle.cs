using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query lifecycle operations (List, View, Import, Export, Update, Delete)
/// </summary>
public partial class PowerQueryCommands
{
    /// <inheritdoc />
    public PowerQueryListResult List(IExcelBatch batch)
    {
        var result = new PowerQueryListResult { FilePath = batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
        {
            Excel.Queries? queriesCollection = null;
            try
            {
                queriesCollection = ctx.Book.Queries;
                int count = queriesCollection.Count;

                for (int i = 1; i <= count; i++)
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.WorkbookQuery? query = null;
                    try
                    {
                        query = queriesCollection.Item(i);
                        string name = query.Name ??
                            throw new InvalidOperationException(
                                $"Power Query at index {i} did not expose a name.");
                        string formula = query.Formula?.ToString() ?? string.Empty;
                        string preview = formula.Length > 80 ? formula[..77] + "..." : formula;
                        var loadState = DetectLoadState(ctx.Book, name, ct);

                        result.Queries.Add(new PowerQueryInfo
                        {
                            Name = name,
                            FormulaPreview = preview,
                            CharacterCount = formula.Length,
                            LoadMode = loadState.LoadMode,
                            TargetSheet = loadState.TargetSheet,
                            IsConnectionOnly = loadState.IsConnectionOnly,
                            IsLoadedToDataModel = loadState.IsLoadedToDataModel
                        });
                    }
                    finally
                    {
                        ComUtilities.Release(ref query);
                    }
                }

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref queriesCollection);
            }
        });
    }

    // View method moved to PowerQueryCommands.View.cs (standalone implementation)

    /// <inheritdoc />
    public PowerQueryLoadConfigResult GetLoadConfig(IExcelBatch batch, string queryName)
    {
        var result = new PowerQueryLoadConfigResult
        {
            FilePath = batch.WorkbookPath,
            QueryName = queryName
        };

        // Validate query name
        if (!ValidateQueryName(queryName, out string? validationError))
        {
            throw new ArgumentException(validationError, nameof(queryName));
        }

        return batch.Execute((ctx, ct) =>
        {
            Excel.WorkbookQuery? query = null;
            try
            {
                query = PowerQuery.PowerQueryHelpers.FindQueryByExactName(ctx.Book, queryName);
                if (query == null)
                {
                    throw new InvalidOperationException($"Query '{queryName}' not found.");
                }

                var loadState = DetectLoadState(ctx.Book, queryName, ct);
                result.HasConnection = loadState.HasConnection;
                result.LoadMode = loadState.LoadMode;
                result.TargetSheet = loadState.TargetSheet;
                result.IsLoadedToDataModel = loadState.IsLoadedToDataModel;
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref query);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Delete(IExcelBatch batch, string queryName)
    {
        // Validate query name
        if (!ValidateQueryName(queryName, out string? validationError))
        {
            throw new ArgumentException(validationError, nameof(queryName));
        }

        return batch.Execute((ctx, ct) =>
        {
            Excel.WorkbookQuery? query = null;
            Excel.Sheets? worksheets = null;

            try
            {
                query = PowerQuery.PowerQueryHelpers.FindQueryByExactName(ctx.Book, queryName);
                if (query == null)
                {
                    throw new InvalidOperationException($"Query '{queryName}' not found.");
                }

                // STEP 1: Clean up any ListObjects (tables) that reference this query
                // When a query is loaded to a worksheet, Excel creates a ListObject with QueryTable
                // Delete must remove these to prevent orphaned tables
                worksheets = ctx.Book.Worksheets;
                int worksheetCount = worksheets.Count;

                for (int i = 1; i <= worksheetCount; i++)
                {
                    Excel.Worksheet? sheet = null;
                    Excel.ListObjects? listObjects = null;

                    try
                    {
                        sheet = (Excel.Worksheet)worksheets[i];
                        listObjects = sheet.ListObjects;
                        int tableCount = listObjects.Count;

                        // Iterate backwards to safely delete while iterating
                        for (int j = tableCount; j >= 1; j--)
                        {
                            Excel.ListObject? table = null;
                            Excel.QueryTable? queryTable = null;
                            Excel.WorkbookConnection? workbookConnection = null;
                            Excel.OLEDBConnection? oleDbConnection = null;

                            try
                            {
                                table = listObjects[j];
                                int sourceType = (int)table.SourceType;
                                if (sourceType is not (0 or 3))
                                    continue;

                                // Check if this table has a QueryTable with our query
                                queryTable = table.QueryTable;
                                if (queryTable != null)
                                {
                                    workbookConnection = queryTable.WorkbookConnection;
                                    if (workbookConnection is null ||
                                        workbookConnection.Type != Excel.XlConnectionType.xlConnectionTypeOLEDB)
                                        continue;
                                    oleDbConnection = workbookConnection.OLEDBConnection;
                                    if (oleDbConnection != null)
                                    {
                                        object? connectionValue = oleDbConnection.Connection;
                                        string? connString = connectionValue?.ToString() ?? "";
                                        if (PowerQuery.PowerQueryHelpers.MatchesMashupLocation(
                                            connString,
                                            queryName))
                                        {
                                            // This table is associated with our query - delete it
                                            table.Delete();
                                        }
                                    }
                                }
                            }
                            finally
                            {
                                ComUtilities.Release(ref oleDbConnection);
                                ComUtilities.Release(ref workbookConnection);
                                ComUtilities.Release(ref queryTable);
                                ComUtilities.Release(ref table);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref listObjects);
                        ComUtilities.Release(ref sheet);
                    }
                }

                PowerQuery.PowerQueryHelpers.RemoveConnectionsByMashupLocation(
                    ctx.Book,
                    queryName);

                query.Delete();

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref worksheets);
                ComUtilities.Release(ref query);
            }
        });
    }


    /// <summary>
    /// Converts query to connection-only (removes data load)
    /// Uses ListObjects pattern (matches Delete cleanup logic)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query</param>
    /// <returns>Operation result</returns>
    public OperationResult Unload(IExcelBatch batch, string queryName)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "unload"
        };

        // Validate query name
        if (!ValidateQueryName(queryName, out string? validationError))
        {
            throw new ArgumentException(validationError, nameof(queryName));
        }

        return batch.Execute((ctx, ct) =>
        {
            Excel.WorkbookQuery? query = null;

            try
            {
                query = PowerQuery.PowerQueryHelpers.FindQueryByExactName(ctx.Book, queryName);
                if (query == null)
                {
                    throw new InvalidOperationException($"Query '{queryName}' not found.");
                }

                UnloadFromDestinations(ctx.Book, queryName);

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref query);
            }
        }, cancellationToken: default);
    }

    /// <summary>
    /// Removes all load destinations for a query (ListObjects and Data Model connections).
    /// Shared logic used by both <see cref="Unload"/> and <see cref="LoadTo"/> ConnectionOnly mode.
    /// The query definition itself is preserved.
    /// </summary>
    private static void UnloadFromDestinations(Excel.Workbook workbook, string queryName)
    {
        Excel.Sheets? worksheets = null;

        try
        {
            // STEP 1: Remove ListObjects (tables) that reference this query
            worksheets = workbook.Worksheets;
            int worksheetCount = worksheets.Count;

            for (int i = 1; i <= worksheetCount; i++)
            {
                Excel.Worksheet? sheet = null;
                Excel.ListObjects? listObjects = null;

                try
                {
                    sheet = (Excel.Worksheet)worksheets[i];
                    listObjects = sheet.ListObjects;
                    int tableCount = listObjects.Count;

                    // Iterate backwards to safely delete while iterating
                    for (int j = tableCount; j >= 1; j--)
                    {
                        Excel.ListObject? table = null;
                        Excel.QueryTable? queryTable = null;
                        Excel.WorkbookConnection? workbookConnection = null;
                        Excel.OLEDBConnection? oleDbConnection = null;

                        try
                        {
                            table = listObjects[j];
                            int sourceType = (int)table.SourceType;
                            if (sourceType is not (0 or 3))
                                continue;

                            // Check if this table has a QueryTable with our query
                            queryTable = table.QueryTable;
                            if (queryTable != null)
                            {
                                workbookConnection = queryTable.WorkbookConnection;
                                if (workbookConnection is null ||
                                    workbookConnection.Type != Excel.XlConnectionType.xlConnectionTypeOLEDB)
                                    continue;
                                oleDbConnection = workbookConnection.OLEDBConnection;
                                if (oleDbConnection != null)
                                {
                                    object? connectionValue = oleDbConnection.Connection;
                                    string? connString = connectionValue?.ToString() ?? "";
                                    if (PowerQuery.PowerQueryHelpers.MatchesMashupLocation(
                                        connString,
                                        queryName))
                                    {
                                        // This table is associated with our query - delete it
                                        table.Delete();
                                    }
                                }
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref oleDbConnection);
                            ComUtilities.Release(ref workbookConnection);
                            ComUtilities.Release(ref queryTable);
                            ComUtilities.Release(ref table);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref listObjects);
                    ComUtilities.Release(ref sheet);
                }
            }

            PowerQuery.PowerQueryHelpers.RemoveConnectionsByMashupLocation(
                workbook,
                queryName);
        }
        finally
        {
            ComUtilities.Release(ref worksheets);
        }
    }
}
