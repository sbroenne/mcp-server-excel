using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

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
        ValidateRequiredTableName(tableName);

        var result = new TableStructuredReferenceResult { FilePath = batch.WorkbookPath };
        return batch.Execute((ctx, ct) =>
        {
            Excel.ListObject? table = null;
            try
            {
                table = FindTable(ctx.Book, tableName);

                result.TableName = tableName;
                result.Region = region;
                result.ColumnName = columnName;

                // Get sheet name
                Excel.Range? range = null;
                Excel.Worksheet? sheet = null;
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
                Excel.Range? regionRange = null;
                Excel.Range? rows = null;
                Excel.Range? columns = null;
                try
                {
                    regionRange = GetRegionRange(table, region, columnName);
                    result.RangeAddress = regionRange.Address;
                    rows = regionRange.Rows;
                    columns = regionRange.Columns;
                    result.RowCount = rows.Count;
                    result.ColumnCount = columns.Count;
                }
                finally
                {
                    ComUtilities.Release(ref columns);
                    ComUtilities.Release(ref rows);
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
    private static Excel.Range GetRegionRange(Excel.ListObject table, TableRegion region, string? columnName)
    {
        Excel.Range? regionRange = region switch
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
            Excel.ListColumns? columns = null;
            Excel.ListColumn? column = null;
            Excel.Range? columnRange = null;
            Excel.Range? intersection = null;
            try
            {
                columns = table.ListColumns;
                column = columns[columnName];
                columnRange = column.Range;

                // Intersect with region range
                Excel.Application? app = null;
                try
                {
                    app = table.Application;
                    intersection = app.Intersect(regionRange, columnRange);
                    Excel.Range result = intersection;
                    intersection = null;
                    return result;
                }
                finally
                {
                    ComUtilities.Release(ref app);
                }
            }
            finally
            {
                ComUtilities.Release(ref intersection);
                ComUtilities.Release(ref columnRange);
                ComUtilities.Release(ref regionRange);
                ComUtilities.Release(ref column);
                ComUtilities.Release(ref columns);
            }
        }

        Excel.Range transferredRange = regionRange!;
        regionRange = null;
        return transferredRange;
    }
}
