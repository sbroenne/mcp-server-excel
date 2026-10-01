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
    public RangeFindResult Find(IExcelBatch batch, string sheetName, string rangeAddress, string searchValue, FindOptions findOptions, int maxMatches = 10)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maxMatches);

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
            dynamic? range = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    throw new InvalidOperationException(specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                }

                // Excel COM constants
                int lookIn = replaceOptions.SearchFormulas && replaceOptions.SearchValues ? -4163 : // xlValues
                             replaceOptions.SearchFormulas ? -4123 : -4163; // xlFormulas : xlValues
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
            dynamic? range = null;
            dynamic? key1 = null;
            dynamic? key2 = null;
            dynamic? key3 = null;
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
                key1 = sortColumns.Count >= 1 ? range.Columns[sortColumns[0].ColumnIndex] : Type.Missing;
                key2 = sortColumns.Count >= 2 ? range.Columns[sortColumns[1].ColumnIndex] : Type.Missing;
                key3 = sortColumns.Count >= 3 ? range.Columns[sortColumns[2].ColumnIndex] : Type.Missing;

                // Excel COM constants
                int order1 = sortColumns[0].Ascending ? 1 : 2; // xlAscending : xlDescending
                int order2 = sortColumns.Count >= 2 ? (sortColumns[1].Ascending ? 1 : 2) : 1;
                int order3 = sortColumns.Count >= 3 ? (sortColumns[2].Ascending ? 1 : 2) : 1;
                int header = hasHeaders ? 1 : 2; // xlYes : xlNo

                // Call Range.Sort method
                range.Sort(
                    Key1: key1,
                    Order1: order1,
                    Key2: key2,
                    Order2: order2,
                    Key3: key3,
                    Order3: order3,
                    Header: header,
                    OrderCustom: 1,
                    MatchCase: false,
                    Orientation: 1, // xlTopToBottom
                    SortMethod: 1   // xlPinYin
                );

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref key3);
                ComUtilities.Release(ref key2);
                ComUtilities.Release(ref key1);
                ComUtilities.Release(ref range);
            }
        });
    }

    // === NATIVE EXCEL COM OPERATIONS ===

}

