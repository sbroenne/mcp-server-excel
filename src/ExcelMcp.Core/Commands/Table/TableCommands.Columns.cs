using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Table column management operations (NEW)
/// </summary>
public partial class TableCommands
{
    /// <inheritdoc />
    public OperationResult AddColumn(IExcelBatch batch, string tableName, string columnName, int? position = null)
    {
        ValidateRequiredTableName(tableName);

        return batch.Execute((ctx, ct) =>
        {
            Excel.ListObject? table = null;
            Excel.ListColumns? listColumns = null;
            Excel.ListColumn? newColumn = null;
            try
            {
                table = FindTable(ctx.Book, tableName);

                // Check if column already exists
                listColumns = table.ListColumns;
                for (int i = 1; i <= listColumns.Count; i++)
                {
                    Excel.ListColumn? column = null;
                    try
                    {
                        column = listColumns[i];
                        if (column.Name == columnName)
                        {
                            throw new InvalidOperationException($"Column '{columnName}' already exists in table '{tableName}'");
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref column);
                    }
                }

                // Add column at specified position or at the end
                if (position.HasValue)
                {
                    newColumn = listColumns.Add(Position: position.Value);
                }
                else
                {
                    newColumn = listColumns.Add(); // Adds at end
                }

                newColumn.Name = columnName;

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref newColumn);
                ComUtilities.Release(ref listColumns);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult RemoveColumn(IExcelBatch batch, string tableName, string columnName)
    {
        ValidateRequiredTableName(tableName);

        return batch.Execute((ctx, ct) =>
        {
            Excel.ListObject? table = null;
            Excel.ListColumns? listColumns = null;
            Excel.ListColumn? column = null;
            try
            {
                table = FindTable(ctx.Book, tableName);

                // Find column
                listColumns = table.ListColumns;
                for (int i = 1; i <= listColumns.Count; i++)
                {
                    Excel.ListColumn? col = null;
                    try
                    {
                        col = listColumns[i];
                        if (col.Name == columnName)
                        {
                            column = col;
                            col = null;
                            break;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref col);
                    }
                }

                if (column == null)
                {
                    throw new InvalidOperationException($"Column '{columnName}' not found in table '{tableName}'");
                }

                // Delete column
                column.Delete();

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref column);
                ComUtilities.Release(ref listColumns);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult RenameColumn(IExcelBatch batch, string tableName, string oldName, string newName)
    {
        ValidateRequiredTableName(tableName);

        return batch.Execute((ctx, ct) =>
        {
            Excel.ListObject? table = null;
            Excel.ListColumns? listColumns = null;
            Excel.ListColumn? column = null;
            try
            {
                table = FindTable(ctx.Book, tableName);

                // Find column
                listColumns = table.ListColumns;
                for (int i = 1; i <= listColumns.Count; i++)
                {
                    Excel.ListColumn? col = null;
                    try
                    {
                        col = listColumns[i];
                        if (col.Name == oldName)
                        {
                            column = col;
                            col = null;
                            break;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref col);
                    }
                }

                if (column == null)
                {
                    throw new InvalidOperationException($"Column '{oldName}' not found in table '{tableName}'");
                }

                // Check if new name already exists
                for (int i = 1; i <= listColumns.Count; i++)
                {
                    Excel.ListColumn? col = null;
                    try
                    {
                        col = listColumns[i];
                        if (col.Name == newName)
                        {
                            throw new InvalidOperationException($"Column '{newName}' already exists in table '{tableName}'");
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref col);
                    }
                }

                // Rename column
                column.Name = newName;

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref column);
                ComUtilities.Release(ref listColumns);
                ComUtilities.Release(ref table);
            }
        });
    }
}
