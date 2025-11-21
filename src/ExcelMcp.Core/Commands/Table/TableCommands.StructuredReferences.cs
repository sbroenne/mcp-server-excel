using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// TableCommands partial class - Structured Reference operations
/// </summary>
public partial class TableCommands
{
    /// <summary>
    /// Gets structured reference information for a table region or column
    /// </summary>
    public TableStructuredReferenceResult GetStructuredReference(
        IExcelBatch batch,
        string tableName,
        TableRegion region,
        string? columnName = null)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        var result = new TableStructuredReferenceResult { FilePath = batch.WorkbookPath };
        return batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            try
            {
                table = FindTable(ctx.Book, tableName);
                if (table == null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' not found");
                }

                result.TableName = tableName;
                result.Region = region;
                result.ColumnName = columnName;

                // Get sheet name
                dynamic? range = null;
                dynamic? sheet = null;
                try
                {
                    range = table.Range;
                    sheet = range.Worksheet;
                    result.SheetName = sheet.Name;
                }
                finally
                {
                    ComUtilities.Release(ref sheet);
                    ComUtilities.Release(ref range);
                }

                // Build structured reference
                result.StructuredReference = BuildStructuredReference(tableName, region, columnName);

                // Get region range
                dynamic? regionRange = null;
                try
                {
                    regionRange = GetRegionRange(table, region, columnName);
                    result.RangeAddress = regionRange.Address;
                    result.RowCount = regionRange.Rows.Count;
                    result.ColumnCount = regionRange.Columns.Count;
                }
                finally
                {
                    ComUtilities.Release(ref regionRange);
                }

                result.Success = true;

                return result;
            }
            finally
            {
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <summary>
    /// Builds a structured reference formula string for a table region
    /// </summary>
    private static string BuildStructuredReference(string tableName, TableRegion region, string? columnName)
    {
        if (!string.IsNullOrEmpty(columnName))
        {
            // Column-specific reference
            return region switch
            {
                TableRegion.All => $"{tableName}[[#All],[{columnName}]]",
                TableRegion.Data => $"{tableName}[[{columnName}]]", // Default is data
                TableRegion.Headers => $"{tableName}[[#Headers],[{columnName}]]",
                TableRegion.Totals => $"{tableName}[[#Totals],[{columnName}]]",
                TableRegion.ThisRow => $"{tableName}[[@],[{columnName}]]",
                _ => $"{tableName}[[{columnName}]]"
            };
        }
        else
        {
            // Entire region reference
            return region switch
            {
                TableRegion.All => $"{tableName}[#All]",
                TableRegion.Data => $"{tableName}[#Data]",
                TableRegion.Headers => $"{tableName}[#Headers]",
                TableRegion.Totals => $"{tableName}[#Totals]",
                TableRegion.ThisRow => $"{tableName}[@]",
                _ => $"{tableName}[#All]"
            };
        }
    }

    /// <summary>
    /// Gets the Excel Range object for a specific table region
    /// </summary>
    private static dynamic GetRegionRange(dynamic table, TableRegion region, string? columnName)
    {
        dynamic regionRange = region switch
        {
            TableRegion.All => table.Range,
            TableRegion.Data => table.DataBodyRange,
            TableRegion.Headers => table.HeaderRowRange,
            TableRegion.Totals => table.TotalsRowRange,
            TableRegion.ThisRow => table.Range, // Full range for ThisRow context
            _ => table.Range
        };

        // If column specified, get intersection with column
        if (!string.IsNullOrEmpty(columnName))
        {
            dynamic? columns = null;
            dynamic? column = null;
            try
            {
                columns = table.ListColumns;
                column = columns.Item(columnName);
                dynamic columnRange = column.Range;

                // Intersect with region range
                dynamic? app = null;
                dynamic? intersection = null;
                try
                {
                    app = table.Application;
                    intersection = app.Intersect(regionRange, columnRange);
                    return intersection; // Return intersection
                }
                catch
                {
                    // If intersection fails, return column range
                    return columnRange;
                }
                finally
                {
                    ComUtilities.Release(ref app);
                    // Don't release intersection here - caller will do it
                }
            }
            finally
            {
                ComUtilities.Release(ref column);
                ComUtilities.Release(ref columns);
            }
        }

        return regionRange; // Return region range directly
    }
}

