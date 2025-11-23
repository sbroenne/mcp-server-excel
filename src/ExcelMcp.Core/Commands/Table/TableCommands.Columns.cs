#pragma warning disable IDE0005 // Using directive is unnecessary (all usings are needed for COM interop)

using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Table column management operations (NEW)
/// </summary>
public partial class TableCommands
{
    /// <inheritdoc />
    public void AddColumn(IExcelBatch batch, string tableName, string columnName, int? position = null)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? listColumns = null;
            dynamic? newColumn = null;
            try
            {
                table = FindTable(ctx.Book, tableName);
                if (table == null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' not found");
                }

                // Check if column already exists
                listColumns = table.ListColumns;
                for (int i = 1; i <= listColumns.Count; i++)
                {
                    dynamic? column = null;
                    try
                    {
                        column = listColumns.Item(i);
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

                return 0;
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
    public void RemoveColumn(IExcelBatch batch, string tableName, string columnName)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? listColumns = null;
            dynamic? column = null;
            try
            {
                table = FindTable(ctx.Book, tableName);
                if (table == null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' not found");
                }

                // Find column
                listColumns = table.ListColumns;
                for (int i = 1; i <= listColumns.Count; i++)
                {
                    dynamic? col = null;
                    try
                    {
                        col = listColumns.Item(i);
                        if (col.Name == columnName)
                        {
                            column = col;
                            break;
                        }
                    }
                    finally
                    {
                        if (col != null && col.Name != columnName)
                        {
                            ComUtilities.Release(ref col);
                        }
                    }
                }

                if (column == null)
                {
                    throw new InvalidOperationException($"Column '{columnName}' not found in table '{tableName}'");
                }

                // Delete column
                column.Delete();

                return 0;
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
    public void RenameColumn(IExcelBatch batch, string tableName, string oldName, string newName)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? listColumns = null;
            dynamic? column = null;
            try
            {
                table = FindTable(ctx.Book, tableName);
                if (table == null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' not found");
                }

                // Find column
                listColumns = table.ListColumns;
                for (int i = 1; i <= listColumns.Count; i++)
                {
                    dynamic? col = null;
                    try
                    {
                        col = listColumns.Item(i);
                        if (col.Name == oldName)
                        {
                            column = col;
                            break;
                        }
                    }
                    finally
                    {
                        if (col != null && col.Name != oldName)
                        {
                            ComUtilities.Release(ref col);
                        }
                    }
                }

                if (column == null)
                {
                    throw new InvalidOperationException($"Column '{oldName}' not found in table '{tableName}'");
                }

                // Check if new name already exists
                for (int i = 1; i <= listColumns.Count; i++)
                {
                    dynamic? col = null;
                    try
                    {
                        col = listColumns.Item(i);
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

                return 0;
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

