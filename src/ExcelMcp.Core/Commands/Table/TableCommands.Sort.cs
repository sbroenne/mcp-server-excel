using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// TableCommands partial class - Sort operations
/// </summary>
public partial class TableCommands
{
    // Excel constants for sorting
    private const int xlYes = 1;
    private const int xlAscending = 1;
    private const int xlDescending = 2;

    /// <summary>
    /// Sorts a table by a single column
    /// </summary>
    public OperationResult Sort(
        IExcelBatch batch,
        string tableName,
        string columnName,
        bool ascending = true)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "sort-table" };
        return batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? columns = null;
            dynamic? column = null;
            dynamic? sortRange = null;
            dynamic? columnRange = null;

            try
            {
                table = FindTable(ctx.Book, tableName);
                if (table == null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' not found");
                }

                // Find column
                columns = table.ListColumns;
                column = columns.Item(columnName);
                if (column == null)
                {
                    throw new InvalidOperationException($"Column '{columnName}' not found in table '{tableName}'");
                }

                // Get ranges for sorting
                sortRange = table.Range;
                columnRange = column.Range;

                // Perform sort
                sortRange.Sort(
                    Key1: columnRange,
                    Order1: ascending ? xlAscending : xlDescending,
                    Header: xlYes
                );

                result.Success = true;

                return result;
            }
            finally
            {
                ComUtilities.Release(ref columnRange);
                ComUtilities.Release(ref sortRange);
                ComUtilities.Release(ref column);
                ComUtilities.Release(ref columns);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <summary>
    /// Sorts a table by multiple columns (up to 3 levels)
    /// </summary>
    public OperationResult Sort(
        IExcelBatch batch,
        string tableName,
        List<TableSortColumn> sortColumns)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "sort-table-multi" };
        return batch.Execute((ctx, ct) =>
        {
            if (sortColumns == null || sortColumns.Count == 0)
            {
                throw new ArgumentException("At least one sort column must be specified", nameof(sortColumns));
            }

            if (sortColumns.Count > 3)
            {
                throw new ArgumentException("Excel supports a maximum of 3 sort levels", nameof(sortColumns));
            }

            dynamic? table = null;
            dynamic? sortRange = null;
            dynamic? key1 = null, key2 = null, key3 = null;

            try
            {
                table = FindTable(ctx.Book, tableName);
                if (table == null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' not found");
                }

                sortRange = table.Range;
                dynamic? columns = null;
                try
                {
                    columns = table.ListColumns;

                    // Get column ranges
                    for (int i = 0; i < sortColumns.Count; i++)
                    {
                        dynamic? col = null;
                        try
                        {
                            col = columns.Item(sortColumns[i].ColumnName);
                            if (col == null)
                            {
                                throw new InvalidOperationException($"Column '{sortColumns[i].ColumnName}' not found in table '{tableName}'");
                            }

                            if (i == 0) key1 = col.Range;
                            else if (i == 1) key2 = col.Range;
                            else if (i == 2) key3 = col.Range;
                        }
                        finally
                        {
                            ComUtilities.Release(ref col);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref columns);
                }

                // Perform sort based on number of columns
                if (sortColumns.Count == 1)
                {
                    sortRange.Sort(
                        Key1: key1,
                        Order1: sortColumns[0].Ascending ? xlAscending : xlDescending,
                        Header: xlYes
                    );
                }
                else if (sortColumns.Count == 2)
                {
                    sortRange.Sort(
                        Key1: key1,
                        Order1: sortColumns[0].Ascending ? xlAscending : xlDescending,
                        Key2: key2,
                        Order2: sortColumns[1].Ascending ? xlAscending : xlDescending,
                        Header: xlYes
                    );
                }
                else if (sortColumns.Count == 3)
                {
                    sortRange.Sort(
                        Key1: key1,
                        Order1: sortColumns[0].Ascending ? xlAscending : xlDescending,
                        Key2: key2,
                        Order2: sortColumns[1].Ascending ? xlAscending : xlDescending,
                        Key3: key3,
                        Order3: sortColumns[2].Ascending ? xlAscending : xlDescending,
                        Header: xlYes
                    );
                }

                // Build description
                var sortDesc = string.Join(", ", sortColumns.Select(sc => $"{sc.ColumnName} ({(sc.Ascending ? "asc" : "desc")})"));
                result.Success = true;

                return result;
            }
            finally
            {
                ComUtilities.Release(ref key3);
                ComUtilities.Release(ref key2);
                ComUtilities.Release(ref key1);
                ComUtilities.Release(ref sortRange);
                ComUtilities.Release(ref table);
            }
        });
    }
}

