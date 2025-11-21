using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query Create operation
/// </summary>
public partial class PowerQueryCommands
{
    /// <summary>
    /// Creates new Power Query from M code with specified load destination.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name for the new query</param>
    /// <param name="mCode">Power Query M code</param>
    /// <param name="loadMode">Where to load the data (default: LoadToTable)</param>
    /// <param name="targetSheet">Target worksheet name (defaults to queryName if not specified)</param>
    /// <param name="targetCellAddress">Target cell address (e.g., "A1", "B5")</param>
    /// <returns>Result with query creation and data load status</returns>
    public PowerQueryCreateResult Create(
        IExcelBatch batch,
        string queryName,
        string mCode,
        PowerQueryLoadMode loadMode = PowerQueryLoadMode.LoadToTable,
        string? targetSheet = null,
        string? targetCellAddress = null)
    {
        var result = new PowerQueryCreateResult
        {
            FilePath = batch.WorkbookPath,
            QueryName = queryName,
            LoadDestination = loadMode,
            WorksheetName = targetSheet,
            TargetCellAddress = targetCellAddress
        };

        // Validate inputs
        if (string.IsNullOrWhiteSpace(queryName))
        {
            throw new ArgumentException("Query name cannot be empty", nameof(queryName));
        }

        if (string.IsNullOrWhiteSpace(mCode))
        {
            throw new ArgumentException("M code cannot be empty", nameof(mCode));
        }

        // Resolve target sheet name (default to query name)
        if (loadMode == PowerQueryLoadMode.LoadToTable || loadMode == PowerQueryLoadMode.LoadToBoth)
        {
            targetSheet ??= queryName;
            result.WorksheetName = targetSheet;
        }

        // Resolve target cell address (default to A1)
        targetCellAddress ??= "A1";

        return batch.Execute((ctx, ct) =>
        {
            dynamic? queries = null;
            dynamic? query = null;

            try
            {
                queries = ctx.Book.Queries;

                // Check if query already exists
                dynamic? existingQuery = FindQueryByName(queries, queryName);
                if (existingQuery != null)
                {
                    ComUtilities.Release(ref existingQuery);
                    throw new InvalidOperationException($"Query '{queryName}' already exists");
                }

                // Step 1: Create the query (always creates in ConnectionOnly mode initially)
                query = queries.Add(queryName, mCode);
                result.QueryCreated = true;

                // Step 2: Apply load destination based on mode
                switch (loadMode)
                {
                    case PowerQueryLoadMode.ConnectionOnly:
                        // Query created, no data loading needed
                        result.DataLoaded = false;
                        result.RowsLoaded = 0;
                        result.TargetCellAddress = null;
                        result.Success = true;
                        break;

                    case PowerQueryLoadMode.LoadToTable:
                        LoadQueryToWorksheet(ctx.Book, queryName, targetSheet!, targetCellAddress!, result);
                        break;

                    case PowerQueryLoadMode.LoadToDataModel:
                        LoadQueryToDataModel(ctx.Book, queryName, result);
                        break;

                    case PowerQueryLoadMode.LoadToBoth:
                        // Load to worksheet first
                        if (LoadQueryToWorksheet(ctx.Book, queryName, targetSheet!, targetCellAddress!, result))
                        {
                            // Preserve worksheet properties before loading to Data Model
                            int worksheetRows = result.RowsLoaded;
                            string? worksheetCell = result.TargetCellAddress;
                            string? worksheetName = result.WorksheetName;

                            // Then also load to Data Model
                            if (LoadQueryToDataModel(ctx.Book, queryName, result))
                            {
                                // Restore worksheet properties (Data Model sets them to null/-1)
                                result.RowsLoaded = worksheetRows;
                                result.TargetCellAddress = worksheetCell;
                                result.WorksheetName = worksheetName;
                            }
                        }
                        break;
                }

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref query);
                ComUtilities.Release(ref queries);
            }
        }, cancellationToken: default);
    }

    /// <summary>
    /// Finds a query by name in the queries collection.
    /// Returns null if not found.
    /// </summary>
    private static dynamic? FindQueryByName(dynamic queriesCollection, string queryName)
    {
        try
        {
            int count = queriesCollection.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic? query = null;
                try
                {
                    query = queriesCollection.Item(i);
                    string name = query.Name ?? "";

                    if (name.Equals(queryName, StringComparison.OrdinalIgnoreCase))
                    {
                        return query; // Caller must release
                    }
                }
                finally
                {
                    if (query != null)
                    {
                        ComUtilities.Release(ref query);
                    }
                }
            }
        }
        catch
        {
            // Query not found or error accessing collection
        }

        return null;
    }
}
