using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

public partial class WorkbookCommands
{
    private const int MaxOverviewItems = 100;
    private const int MaxOverviewPreviewRows = 10;
    private const int MaxOverviewPreviewColumns = 10;
    private const int MaxOverviewCellCharacters = 200;
    private const int MaxOverviewPreviewCharacters = 8192;

    /// <inheritdoc />
    public WorkbookOverviewResult Inspect(
        IExcelBatch batch,
        string? sheetName = null,
        bool includeSheets = true,
        bool includeTables = true,
        bool includeDefinedNames = true,
        bool includePreview = false,
        string? rangeAddress = null,
        int maxItems = 50,
        int maxPreviewRows = 6,
        int maxPreviewColumns = 6,
        int maxCellCharacters = 160,
        int maxPreviewCharacters = 4096)
    {
        ValidateInspectArguments(sheetName, rangeAddress, includeSheets, includeTables,
            includeDefinedNames, includePreview, maxItems, maxPreviewRows, maxPreviewColumns,
            maxCellCharacters, maxPreviewCharacters);

        return batch.Execute((ctx, ct) =>
        {
            Excel.Sheets? worksheets = null;
            Excel.Names? workbookNames = null;
            Excel.Worksheet? filterWorksheet = null;
            var result = new WorkbookOverviewResult
            {
                Success = true,
                FilePath = batch.WorkbookPath
            };
            if (includeSheets)
                result.Sheets = new WorkbookOverviewSheetSection();
            if (includeTables)
                result.Tables = new WorkbookOverviewTableSection();
            if (includeDefinedNames)
                result.DefinedNames = new WorkbookOverviewNameSection();

            try
            {
                string? selectedSheetName = sheetName;
                if (sheetName is not null)
                {
                    filterWorksheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                    if (filterWorksheet is null)
                        throw new InvalidOperationException($"Worksheet '{sheetName}' was not found.");
                    selectedSheetName = filterWorksheet.Name;
                    ComUtilities.Release(ref filterWorksheet);
                }

                if (result.Sheets is not null || result.Tables is not null)
                {
                    worksheets = ctx.Book.Worksheets;
                    var worksheetTotal = worksheets.Count;
                    for (int index = 1; index <= worksheetTotal; index++)
                    {
                        ct.ThrowIfCancellationRequested();
                        Excel.Worksheet? worksheet = null;
                        try
                        {
                            worksheet = (Excel.Worksheet)worksheets.Item[index];
                            if (selectedSheetName is not null &&
                                !string.Equals(worksheet.Name, selectedSheetName, StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }

                            if (result.Sheets is not null)
                            {
                                result.Sheets.Count++;
                                if (result.Sheets.Items.Count < maxItems)
                                {
                                    Excel.Range? usedRange = null;
                                    Excel.Range? usedRows = null;
                                    Excel.Range? usedColumns = null;
                                    try
                                    {
                                        usedRange = worksheet.UsedRange;
                                        usedRows = usedRange.Rows;
                                        usedColumns = usedRange.Columns;
                                        result.Sheets.Items.Add(new WorkbookOverviewSheet
                                        {
                                            Name = worksheet.Name,
                                            Visibility = GetVisibilityName(worksheet.Visible),
                                            UsedRangeAddress = usedRange.Address,
                                            UsedRowCount = usedRows.Count,
                                            UsedColumnCount = usedColumns.Count
                                        });
                                    }
                                    finally
                                    {
                                        ComUtilities.Release(ref usedColumns);
                                        ComUtilities.Release(ref usedRows);
                                        ComUtilities.Release(ref usedRange);
                                    }
                                }
                            }

                            if (result.Tables is not null)
                            {
                                Excel.ListObjects? tables = null;
                                try
                                {
                                    tables = worksheet.ListObjects;
                                    int tableTotal = tables.Count;
                                    for (int tableIndex = 1; tableIndex <= tableTotal; tableIndex++)
                                    {
                                        ct.ThrowIfCancellationRequested();
                                        result.Tables.Count++;
                                        if (result.Tables.Items.Count >= maxItems)
                                            continue;

                                        Excel.ListObject? table = null;
                                        Excel.Range? tableRange = null;
                                        try
                                        {
                                            table = tables.Item[tableIndex];
                                            tableRange = table.Range;
                                            result.Tables.Items.Add(new WorkbookOverviewTable
                                            {
                                                Name = table.Name,
                                                SheetName = worksheet.Name,
                                                RangeAddress = tableRange.Address
                                            });
                                        }
                                        finally
                                        {
                                            ComUtilities.Release(ref tableRange);
                                            ComUtilities.Release(ref table);
                                        }
                                    }
                                }
                                finally
                                {
                                    ComUtilities.Release(ref tables);
                                }
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref worksheet);
                        }
                    }
                }

                if (result.Sheets is not null)
                    result.Sheets.OmittedCount = result.Sheets.Count - result.Sheets.Items.Count;
                if (result.Tables is not null)
                    result.Tables.OmittedCount = result.Tables.Count - result.Tables.Items.Count;

                if (result.DefinedNames is not null)
                {
                    workbookNames = ctx.Book.Names;
                    int namesTotal = workbookNames.Count;
                    for (int index = 1; index <= namesTotal; index++)
                    {
                        ct.ThrowIfCancellationRequested();
                        Excel.Name? name = null;
                        try
                        {
                            name = workbookNames.Item(index);
                            string fullName = name.Name;
                            if (!name.Visible || IsBuiltInOverviewName(fullName))
                                continue;

                            result.DefinedNames.Count++;
                            if (result.DefinedNames.Items.Count < maxItems)
                            {
                                result.DefinedNames.Items.Add(new WorkbookOverviewName
                                {
                                    Name = fullName,
                                    RefersTo = name.RefersTo
                                });
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref name);
                        }
                    }
                    result.DefinedNames.OmittedCount =
                        result.DefinedNames.Count - result.DefinedNames.Items.Count;
                }

                if (includePreview)
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.Worksheet? worksheet = null;
                    Excel.Range? scope = null;
                    Excel.Range? scopeRows = null;
                    Excel.Range? scopeColumns = null;
                    Excel.Range? previewRange = null;
                    Excel.Worksheet? previewParent = null;
                    Excel.Areas? areas = null;
                    try
                    {
                        worksheet = ComUtilities.FindSheet(ctx.Book, selectedSheetName!);
                        if (worksheet is null)
                            throw new InvalidOperationException($"Worksheet '{selectedSheetName}' was not found.");

                        scope = rangeAddress is null
                            ? worksheet.UsedRange
                            : RangeHelpers.ResolveRange(ctx.Book, selectedSheetName!, rangeAddress);
                        if (scope is null)
                            throw new InvalidOperationException($"Range '{rangeAddress}' on worksheet '{selectedSheetName}' was not found.");

                        areas = scope.Areas;
                        if (areas.Count != 1)
                            throw new ArgumentException("Workbook preview requires one contiguous range.", nameof(rangeAddress));

                        previewParent = (Excel.Worksheet)scope.Parent;
                        if (!string.Equals(previewParent.Name, selectedSheetName, StringComparison.OrdinalIgnoreCase))
                            throw new ArgumentException("The preview range must belong to sheetName.", nameof(rangeAddress));

                        scopeRows = scope.Rows;
                        scopeColumns = scope.Columns;
                        int previewRowCount = Math.Min(scopeRows.Count, maxPreviewRows);
                        int previewColumnCount = Math.Min(scopeColumns.Count, maxPreviewColumns);
                        previewRange = scope.Resize[previewRowCount, previewColumnCount];

                        var previewValues = ExcelValueNormalizer.Normalize(previewRange.Value2);
                        object formulaValues = ctx.Capabilities.SupportsFormula2
                            ? previewRange.Formula2
                            : previewRange.Formula;
                        var previewFormulas = ExcelValueNormalizer.Normalize(formulaValues);
                        NormalizePreviewErrors(previewValues);
                        NormalizePreviewErrors(previewFormulas);

                        int remainingCharacters = maxPreviewCharacters;
                        int returnedCharacters = 0;
                        int truncatedCells = 0;
                        ApplyPreviewTextLimits(previewValues.Values, maxCellCharacters,
                            ref remainingCharacters, ref returnedCharacters, ref truncatedCells);
                        ApplyPreviewTextLimits(previewFormulas.Values, maxCellCharacters,
                            ref remainingCharacters, ref returnedCharacters, ref truncatedCells);

                        result.Preview = new WorkbookOverviewPreview
                        {
                            SheetName = worksheet.Name,
                            RangeAddress = previewRange.Address,
                            RowCount = previewValues.RowCount,
                            ColumnCount = previewValues.ColumnCount,
                            OmittedRowCount = scopeRows.Count - previewRowCount,
                            OmittedColumnCount = scopeColumns.Count - previewColumnCount,
                            TextCharactersReturned = returnedCharacters,
                            TruncatedTextCellCount = truncatedCells,
                            Values = previewValues.Values,
                            Formulas = previewFormulas.Values
                        };
                    }
                    finally
                    {
                        ComUtilities.Release(ref areas);
                        ComUtilities.Release(ref previewParent);
                        ComUtilities.Release(ref previewRange);
                        ComUtilities.Release(ref scopeColumns);
                        ComUtilities.Release(ref scopeRows);
                        ComUtilities.Release(ref scope);
                        ComUtilities.Release(ref worksheet);
                    }
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref filterWorksheet);
                ComUtilities.Release(ref workbookNames);
                ComUtilities.Release(ref worksheets);
            }
        });
    }

    private static void ValidateInspectArguments(
        string? sheetName,
        string? rangeAddress,
        bool includeSheets,
        bool includeTables,
        bool includeDefinedNames,
        bool includePreview,
        int maxItems,
        int maxPreviewRows,
        int maxPreviewColumns,
        int maxCellCharacters,
        int maxPreviewCharacters)
    {
        if (sheetName is not null && string.IsNullOrWhiteSpace(sheetName))
            throw new ArgumentException("sheetName cannot be empty or whitespace.", nameof(sheetName));
        if (includePreview && sheetName is null)
            throw new ArgumentException("sheetName is required when includePreview is true.", nameof(sheetName));
        if (!includePreview && rangeAddress is not null)
            throw new ArgumentException("rangeAddress requires includePreview=true.", nameof(rangeAddress));
        if (rangeAddress is not null && string.IsNullOrWhiteSpace(rangeAddress))
            throw new ArgumentException("rangeAddress cannot be empty or whitespace.", nameof(rangeAddress));
        if (!includeSheets && !includeTables && !includeDefinedNames && !includePreview)
            throw new ArgumentException("At least one overview section must be included.");
        if (maxItems is < 1 or > MaxOverviewItems)
            throw new ArgumentOutOfRangeException(nameof(maxItems), $"maxItems must be between 1 and {MaxOverviewItems}.");
        if (maxPreviewRows is < 1 or > MaxOverviewPreviewRows)
            throw new ArgumentOutOfRangeException(nameof(maxPreviewRows), $"maxPreviewRows must be between 1 and {MaxOverviewPreviewRows}.");
        if (maxPreviewColumns is < 1 or > MaxOverviewPreviewColumns)
            throw new ArgumentOutOfRangeException(nameof(maxPreviewColumns), $"maxPreviewColumns must be between 1 and {MaxOverviewPreviewColumns}.");
        if (maxCellCharacters is < 1 or > MaxOverviewCellCharacters)
            throw new ArgumentOutOfRangeException(nameof(maxCellCharacters), $"maxCellCharacters must be between 1 and {MaxOverviewCellCharacters}.");
        if (maxPreviewCharacters is < 1 or > MaxOverviewPreviewCharacters)
            throw new ArgumentOutOfRangeException(nameof(maxPreviewCharacters), $"maxPreviewCharacters must be between 1 and {MaxOverviewPreviewCharacters}.");
    }

    private static string GetVisibilityName(Excel.XlSheetVisibility visibility) =>
        visibility switch
        {
            Excel.XlSheetVisibility.xlSheetVisible => "Visible",
            Excel.XlSheetVisibility.xlSheetHidden => "Hidden",
            Excel.XlSheetVisibility.xlSheetVeryHidden => "VeryHidden",
            _ => $"Unknown({(int)visibility})"
        };

    private static bool IsBuiltInOverviewName(string fullName)
    {
        string localName = fullName[(fullName.LastIndexOf('!') + 1)..].Trim('\'');
        return localName.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase)
            || localName.Equals("_FilterDatabase", StringComparison.OrdinalIgnoreCase);
    }

    private static void NormalizePreviewErrors(ExcelValueGrid grid)
    {
        for (int row = 0; row < grid.RowCount; row++)
        {
            for (int column = 0; column < grid.ColumnCount; column++)
            {
                if (ExcelErrorMapper.TryGet(grid.Values[row][column], out _, out var error))
                    grid.Values[row][column] = error.Name;
            }
        }
    }

    private static void ApplyPreviewTextLimits(
        List<List<object?>> values,
        int maxCellCharacters,
        ref int remainingCharacters,
        ref int returnedCharacters,
        ref int truncatedCells)
    {
        foreach (var row in values)
        {
            for (int column = 0; column < row.Count; column++)
            {
                if (row[column] is not string text)
                    continue;

                int allowed = Math.Min(maxCellCharacters, remainingCharacters);
                if (text.Length > allowed)
                {
                    row[column] = text[..allowed];
                    truncatedCells++;
                    returnedCharacters += allowed;
                    remainingCharacters -= allowed;
                }
                else
                {
                    returnedCharacters += text.Length;
                    remainingCharacters -= text.Length;
                }
            }
        }
    }
}
