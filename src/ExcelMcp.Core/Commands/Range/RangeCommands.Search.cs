using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;


namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Range search and sort operations (find, replace, sort)
/// </summary>
public partial class RangeCommands
{
    // === FIND/REPLACE OPERATIONS ===

    /// <inheritdoc />
    public RangeFindResult Find(IExcelBatch batch, string sheetName, string rangeAddress, string searchValue, FindOptions? findOptions = null, int maxMatches = 10)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maxMatches);
        findOptions ??= new FindOptions();

        var result = new RangeFindResult
        {
            FilePath = batch.WorkbookPath,
            SheetName = sheetName,
            RangeAddress = rangeAddress,
            SearchValue = searchValue
        };

        return batch.Execute((ctx, ct) =>
        {
            Excel.Range? range = null;
            Excel.Range? foundCell = null;
            Excel.Range? nextCell = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    throw new InvalidOperationException(specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                }

                var lookIn = findOptions.SearchFormulas && !findOptions.SearchValues
                    ? Excel.XlFindLookIn.xlFormulas
                    : Excel.XlFindLookIn.xlValues;
                var lookAt = findOptions.MatchEntireCell
                    ? Excel.XlLookAt.xlWhole
                    : Excel.XlLookAt.xlPart;

                foundCell = range.Find(
                    What: searchValue,
                    LookIn: lookIn,
                    LookAt: lookAt,
                    SearchOrder: Excel.XlSearchOrder.xlByRows,
                    SearchDirection: Excel.XlSearchDirection.xlNext,
                    MatchCase: findOptions.MatchCase
                );

                if (foundCell != null)
                {
                    string firstAddress = foundCell.Address;
                    do
                    {
                        ct.ThrowIfCancellationRequested();
                        result.TotalCount++;
                        if (result.MatchingCells.Count < maxMatches)
                        {
                            result.MatchingCells.Add(new RangeCell
                            {
                                Address = foundCell.Address,
                                Row = foundCell.Row,
                                Column = foundCell.Column,
                                Value = foundCell.Value2
                            });
                        }

                        nextCell = range.FindNext(foundCell);
                        ComUtilities.Release(ref foundCell);
                        foundCell = nextCell;
                        nextCell = null;
                    } while (foundCell != null && foundCell.Address != firstAddress);
                }

                ct.ThrowIfCancellationRequested();
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref nextCell);
                ComUtilities.Release(ref foundCell);
                ComUtilities.Release(ref range);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Replace(IExcelBatch batch, string sheetName, string rangeAddress, string findValue, string replaceValue, ReplaceOptions replaceOptions)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Range? range = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    throw new InvalidOperationException(specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                }

                // Excel COM constants
                int lookAt = replaceOptions.MatchEntireCell ? 1 : 2; // xlWhole : xlPart

                range.Replace(
                    What: findValue,
                    Replacement: replaceValue,
                    LookAt: lookAt,
                    SearchOrder: 1, // xlByRows
                    MatchCase: replaceOptions.MatchCase,
                    MatchByte: false
                );

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref range);
            }
        });
    }

    // === SORT OPERATIONS ===

    /// <inheritdoc />
    public OperationResult Sort(IExcelBatch batch, string sheetName, string rangeAddress, List<SortColumn> sortColumns, bool hasHeaders = true)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Range? range = null;
            Excel.Range? columns = null;
            Excel.Range? key1 = null;
            Excel.Range? key2 = null;
            Excel.Range? key3 = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    throw new InvalidOperationException(specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                }

                if (sortColumns.Count == 0)
                {
                    throw new ArgumentException("At least one sort column must be specified", nameof(sortColumns));
                }

                // Excel COM supports up to 3 sort keys
                if (sortColumns.Count > 3)
                {
                    throw new ArgumentException("Excel COM API supports maximum 3 sort columns", nameof(sortColumns));
                }

                // Get sort key ranges
                columns = range.Columns;
                key1 = (Excel.Range)columns[sortColumns[0].ColumnIndex];
                if (sortColumns.Count >= 2) key2 = (Excel.Range)columns[sortColumns[1].ColumnIndex];
                if (sortColumns.Count >= 3) key3 = (Excel.Range)columns[sortColumns[2].ColumnIndex];

                // Excel COM constants
                int order1 = sortColumns[0].Ascending ? 1 : 2; // xlAscending : xlDescending
                int order2 = sortColumns.Count >= 2 ? (sortColumns[1].Ascending ? 1 : 2) : 1;
                int order3 = sortColumns.Count >= 3 ? (sortColumns[2].Ascending ? 1 : 2) : 1;
                int header = hasHeaders ? 1 : 2; // xlYes : xlNo

                // Call Range.Sort method
                range.Sort(
                    Key1: key1,
                    Order1: (Excel.XlSortOrder)order1,
                    Key2: (object?)key2 ?? Type.Missing,
                    Order2: (Excel.XlSortOrder)order2,
                    Key3: (object?)key3 ?? Type.Missing,
                    Order3: (Excel.XlSortOrder)order3,
                    Header: (Excel.XlYesNoGuess)header,
                    OrderCustom: 1,
                    MatchCase: false,
                    Orientation: (Excel.XlSortOrientation)1,
                    SortMethod: Excel.XlSortMethod.xlPinYin
                );

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref key3);
                ComUtilities.Release(ref key2);
                ComUtilities.Release(ref key1);
                ComUtilities.Release(ref columns);
                ComUtilities.Release(ref range);
            }
        });
    }

    // === NATIVE EXCEL COM OPERATIONS ===

}
