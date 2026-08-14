using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet row and column grouping and outline controls.
/// </summary>
public partial class SheetCommands
{
    /// <inheritdoc />
    public OperationResult Group(IExcelBatch batch, string sheetName, string rangeAddress, OutlineAxis axis)
    {
        return ModifyOutlineRange(batch, sheetName, rangeAddress, axis, "group", target => target.Group());
    }

    /// <inheritdoc />
    public OperationResult Ungroup(IExcelBatch batch, string sheetName, string rangeAddress, OutlineAxis axis)
    {
        return ModifyOutlineRange(batch, sheetName, rangeAddress, axis, "ungroup", target => target.Ungroup());
    }

    /// <inheritdoc />
    public WorksheetOutlineResult GetOutlineInfo(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress,
        OutlineAxis axis)
    {
        var result = new WorksheetOutlineResult
        {
            Action = "get-outline-info",
            FilePath = batch.WorkbookPath,
            SheetName = sheetName,
            RangeAddress = rangeAddress,
            Axis = axis
        };

        return batch.Execute((ctx, ct) =>
        {
            ValidateOutlineAxis(axis);

            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.Range? target = null;
            Excel.Outline? outline = null;
            try
            {
                sheet = FindRequiredSheet(ctx.Book, sheetName);
                range = sheet.Range[rangeAddress];
                target = axis == OutlineAxis.Rows ? range.EntireRow : range.EntireColumn;
                outline = sheet.Outline;

                result.OutlineLevel = Convert.ToInt32(target.OutlineLevel, System.Globalization.CultureInfo.InvariantCulture);
                result.Hidden = Convert.ToBoolean(target.Hidden, System.Globalization.CultureInfo.InvariantCulture);
                result.SummaryRow = outline.SummaryRow == Excel.XlSummaryRow.xlSummaryAbove ? "above" : "below";
                result.SummaryColumn = outline.SummaryColumn == Excel.XlSummaryColumn.xlSummaryOnLeft ? "left" : "right";
                result.AutomaticStyles = outline.AutomaticStyles;
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref outline);
                ComUtilities.Release(ref target);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult SetOutlineSettings(
        IExcelBatch batch,
        string sheetName,
        string? summaryRow = null,
        string? summaryColumn = null,
        bool? automaticStyles = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            if (summaryRow == null && summaryColumn == null && !automaticStyles.HasValue)
            {
                throw new ArgumentException("Provide at least one outline setting to change.");
            }

            Excel.XlSummaryRow? parsedSummaryRow = summaryRow != null ? ParseSummaryRow(summaryRow) : null;
            Excel.XlSummaryColumn? parsedSummaryColumn = summaryColumn != null ? ParseSummaryColumn(summaryColumn) : null;

            Excel.Worksheet? sheet = null;
            Excel.Outline? outline = null;
            try
            {
                sheet = FindRequiredSheet(ctx.Book, sheetName);
                outline = sheet.Outline;

                if (parsedSummaryRow.HasValue)
                {
                    outline.SummaryRow = parsedSummaryRow.Value;
                }

                if (parsedSummaryColumn.HasValue)
                {
                    outline.SummaryColumn = parsedSummaryColumn.Value;
                }

                if (automaticStyles.HasValue)
                {
                    outline.AutomaticStyles = automaticStyles.Value;
                }

                return OutlineOperationResult(batch, "set-outline-settings", sheetName);
            }
            finally
            {
                ComUtilities.Release(ref outline);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult ShowOutlineLevels(
        IExcelBatch batch,
        string sheetName,
        int? rowLevels = null,
        int? columnLevels = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            if (!rowLevels.HasValue && !columnLevels.HasValue)
            {
                throw new ArgumentException("Provide rowLevels, columnLevels, or both.");
            }

            if (rowLevels is <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(rowLevels), rowLevels, "Row outline level must be positive.");
            }

            if (columnLevels is <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(columnLevels), columnLevels, "Column outline level must be positive.");
            }

            Excel.Worksheet? sheet = null;
            Excel.Outline? outline = null;
            object? showResult = null;
            try
            {
                sheet = FindRequiredSheet(ctx.Book, sheetName);
                outline = sheet.Outline;
                showResult = outline.ShowLevels(
                    rowLevels.HasValue ? rowLevels.Value : Type.Missing,
                    columnLevels.HasValue ? columnLevels.Value : Type.Missing);

                return OutlineOperationResult(batch, "show-outline-levels", sheetName);
            }
            finally
            {
                ComUtilities.Release(ref showResult);
                ComUtilities.Release(ref outline);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult ClearOutline(IExcelBatch batch, string sheetName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cells = null;
            object? clearResult = null;
            try
            {
                sheet = FindRequiredSheet(ctx.Book, sheetName);
                cells = sheet.Cells;
                clearResult = cells.ClearOutline();
                return OutlineOperationResult(batch, "clear-outline", sheetName);
            }
            finally
            {
                ComUtilities.Release(ref clearResult);
                ComUtilities.Release(ref cells);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static OperationResult ModifyOutlineRange(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress,
        OutlineAxis axis,
        string action,
        Func<Excel.Range, object> modify)
    {
        return batch.Execute((ctx, ct) =>
        {
            ValidateOutlineAxis(axis);

            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.Range? target = null;
            object? modifyResult = null;
            try
            {
                sheet = FindRequiredSheet(ctx.Book, sheetName);
                range = sheet.Range[rangeAddress];
                target = axis == OutlineAxis.Rows ? range.EntireRow : range.EntireColumn;
                modifyResult = modify(target);
                return OutlineOperationResult(batch, action, sheetName);
            }
            finally
            {
                ComUtilities.Release(ref modifyResult);
                ComUtilities.Release(ref target);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static Excel.Worksheet FindRequiredSheet(Excel.Workbook workbook, string sheetName)
    {
        return ComUtilities.FindSheet(workbook, sheetName)
            ?? throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
    }

    private static Excel.XlSummaryRow ParseSummaryRow(string summaryRow)
    {
        return summaryRow.ToLowerInvariant() switch
        {
            "above" => Excel.XlSummaryRow.xlSummaryAbove,
            "below" => Excel.XlSummaryRow.xlSummaryBelow,
            _ => throw new ArgumentException(
                $"Unknown summary row position: '{summaryRow}'. Valid values: above, below.",
                nameof(summaryRow))
        };
    }

    private static void ValidateOutlineAxis(OutlineAxis axis)
    {
        if (axis is not OutlineAxis.Rows and not OutlineAxis.Columns)
        {
            throw new ArgumentOutOfRangeException(nameof(axis), axis, "Axis must be Rows or Columns.");
        }
    }

    private static Excel.XlSummaryColumn ParseSummaryColumn(string summaryColumn)
    {
        return summaryColumn.ToLowerInvariant() switch
        {
            "left" => Excel.XlSummaryColumn.xlSummaryOnLeft,
            "right" => Excel.XlSummaryColumn.xlSummaryOnRight,
            _ => throw new ArgumentException(
                $"Unknown summary column position: '{summaryColumn}'. Valid values: left, right.",
                nameof(summaryColumn))
        };
    }

    private static OperationResult OutlineOperationResult(IExcelBatch batch, string action, string sheetName)
    {
        return new OperationResult
        {
            Success = true,
            Action = action,
            FilePath = batch.WorkbookPath,
            Message = $"Worksheet outline updated for '{sheetName}'."
        };
    }
}
