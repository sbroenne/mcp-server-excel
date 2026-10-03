using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Filtering;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

public partial class TableCommands
{
    /// <inheritdoc />
    public OperationResult ApplyFilter(IExcelBatch batch, string tableName, string columnName, FilterOptions options)
    {
        ValidateRequiredTableName(tableName);
        ArgumentException.ThrowIfNullOrWhiteSpace(columnName);
        NativeFilterHelpers.Validate(options);
        return batch.Execute((ctx, token) =>
        {
            Excel.ListObject? table = null;
            Excel.ListColumn? column = null;
            Excel.Range? range = null;
            try
            {
                table = FindTable(ctx.Book, tableName);
                column = FindColumn(table, columnName);
                if (column is null)
                    throw new InvalidOperationException($"Column '{columnName}' not found in table '{tableName}'.");
                range = table.Range;
                NativeFilterHelpers.Apply(ctx.Book, range, column.Index, options, token);
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath, Action = "apply-filter" };
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref column);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult ClearFilters(IExcelBatch batch, string tableName)
    {
        ValidateRequiredTableName(tableName);
        return batch.Execute((ctx, token) =>
        {
            Excel.ListObject? table = null;
            Excel.AutoFilter? filter = null;
            try
            {
                token.ThrowIfCancellationRequested();
                table = FindTable(ctx.Book, tableName);
                filter = table.AutoFilter;
                if (filter is not null && filter.FilterMode)
                    filter.ShowAllData();
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath, Action = "clear-filters" };
            }
            finally
            {
                ComUtilities.Release(ref filter);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
    public TableFilterResult GetFilters(IExcelBatch batch, string tableName)
    {
        ValidateRequiredTableName(tableName);
        return batch.Execute((ctx, token) =>
        {
            Excel.ListObject? table = null;
            Excel.ListColumns? columns = null;
            Excel.AutoFilter? filter = null;
            try
            {
                token.ThrowIfCancellationRequested();
                table = FindTable(ctx.Book, tableName);
                columns = table.ListColumns;
                var names = new List<string>(columns.Count);
                for (int index = 1; index <= columns.Count; index++)
                {
                    token.ThrowIfCancellationRequested();
                    Excel.ListColumn? column = null;
                    try
                    {
                        column = columns[index];
                        names.Add(column.Name);
                    }
                    finally
                    {
                        ComUtilities.Release(ref column);
                    }
                }
                filter = table.AutoFilter;
                var read = filter is null
                    ? names.Select((name, index) => new ColumnFilter { ColumnName = name, ColumnIndex = index + 1 }).ToList()
                    : NativeFilterHelpers.Read(filter, names, token);
                return new TableFilterResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    TableName = tableName,
                    ColumnFilters = read,
                    HasActiveFilters = read.Any(column => column.IsFiltered)
                };
            }
            finally
            {
                ComUtilities.Release(ref filter);
                ComUtilities.Release(ref columns);
                ComUtilities.Release(ref table);
            }
        });
    }
}
