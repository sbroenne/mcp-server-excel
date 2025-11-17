using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Table filter operations (NEW)
/// </summary>
public partial class TableCommands
{
    /// <inheritdoc />
    public OperationResult ApplyFilter(IExcelBatch batch, string tableName, string columnName, string criteria)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "apply-filter" };
        return batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? autoFilter = null;
            try
            {
                table = FindTable(ctx.Book, tableName);
                if (table == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Table '{tableName}' not found";
                    return result;
                }

                // Find column index
                int columnIndex = -1;
                dynamic? listColumns = null;
                try
                {
                    listColumns = table.ListColumns;
                    for (int i = 1; i <= listColumns.Count; i++)
                    {
                        dynamic? column = null;
                        try
                        {
                            column = listColumns.Item(i);
                            if (column.Name == columnName)
                            {
                                columnIndex = i;
                                break;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref column);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref listColumns);
                }

                if (columnIndex == -1)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Column '{columnName}' not found in table '{tableName}'";
                    return result;
                }

                // Apply filter
                autoFilter = table.AutoFilter;
                if (autoFilter == null)
                {
                    // AutoFilter not enabled - enable it first
                    dynamic? range = null;
                    try
                    {
                        range = table.Range;
                        range.AutoFilter(Field: 1); // Enable with default
                        autoFilter = table.AutoFilter;
                    }
                    finally
                    {
                        ComUtilities.Release(ref range);
                    }
                }

                // Apply filter to specific field
                // xlFilterValues = 7, xlAnd = 1
                int xlFilterValues = 7;
                autoFilter.Range.AutoFilter(
                    Field: columnIndex,
                    Criteria1: criteria,
                    Operator: xlFilterValues
                );

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref autoFilter);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult ApplyFilter(IExcelBatch batch, string tableName, string columnName, List<string> criteria)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "apply-filter-values" };
        return batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? autoFilter = null;
            try
            {
                table = FindTable(ctx.Book, tableName);
                if (table == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Table '{tableName}' not found";
                    return result;
                }

                // Find column index
                int columnIndex = -1;
                dynamic? listColumns = null;
                try
                {
                    listColumns = table.ListColumns;
                    for (int i = 1; i <= listColumns.Count; i++)
                    {
                        dynamic? column = null;
                        try
                        {
                            column = listColumns.Item(i);
                            if (column.Name == columnName)
                            {
                                columnIndex = i;
                                break;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref column);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref listColumns);
                }

                if (columnIndex == -1)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Column '{columnName}' not found in table '{tableName}'";
                    return result;
                }

                // Apply filter
                autoFilter = table.AutoFilter;
                if (autoFilter == null)
                {
                    // AutoFilter not enabled - enable it first
                    dynamic? range = null;
                    try
                    {
                        range = table.Range;
                        range.AutoFilter(Field: 1); // Enable with default
                        autoFilter = table.AutoFilter;
                    }
                    finally
                    {
                        ComUtilities.Release(ref range);
                    }
                }

                // Apply filter with multiple values
                // Convert List<string> to string array for COM interop
                string[] valuesArray = criteria.ToArray();
                autoFilter.Range.AutoFilter(
                    Field: columnIndex,
                    Criteria1: valuesArray,
                    Operator: 7 // xlFilterValues
                );

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref autoFilter);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult ClearFilters(IExcelBatch batch, string tableName)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "clear-filters" };
        return batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? autoFilter = null;
            try
            {
                table = FindTable(ctx.Book, tableName);
                if (table == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Table '{tableName}' not found";
                    return result;
                }

                autoFilter = table.AutoFilter;
                if (autoFilter != null)
                {
                    autoFilter.ShowAllData();
                }

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref autoFilter);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
    public TableFilterResult GetFilters(IExcelBatch batch, string tableName)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        var result = new TableFilterResult { FilePath = batch.WorkbookPath, TableName = tableName };
        return batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? autoFilter = null;
            try
            {
                table = FindTable(ctx.Book, tableName);
                if (table == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Table '{tableName}' not found";
                    return result;
                }

                autoFilter = table.AutoFilter;
                if (autoFilter == null)
                {
                    result.Success = true;
                    result.HasActiveFilters = false;
                    return result;
                }

                // Check each column for filters
                dynamic? filters = null;
                try
                {
                    filters = autoFilter.Filters;
                    dynamic? listColumns = null;
                    try
                    {
                        listColumns = table.ListColumns;
                        for (int i = 1; i <= listColumns.Count; i++)
                        {
                            dynamic? column = null;
                            dynamic? filter = null;
                            try
                            {
                                column = listColumns.Item(i);
                                string columnName = column.Name;

                                filter = filters.Item(i);
                                bool isFiltered = filter.On;

                                if (isFiltered)
                                {
                                    result.HasActiveFilters = true;
                                    result.ColumnFilters.Add(new ColumnFilter
                                    {
                                        ColumnName = columnName,
                                        ColumnIndex = i,
                                        IsFiltered = true,
                                        Criteria = filter.Criteria1?.ToString() ?? "",
                                        FilterValues = [] // Could extract from Criteria1 if array
                                    });
                                }
                                else
                                {
                                    result.ColumnFilters.Add(new ColumnFilter
                                    {
                                        ColumnName = columnName,
                                        ColumnIndex = i,
                                        IsFiltered = false
                                    });
                                }
                            }
                            finally
                            {
                                ComUtilities.Release(ref filter);
                                ComUtilities.Release(ref column);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref listColumns);
                    }
                }
                finally
                {
                    ComUtilities.Release(ref filters);
                }

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref autoFilter);
                ComUtilities.Release(ref table);
            }
        });
    }
}

