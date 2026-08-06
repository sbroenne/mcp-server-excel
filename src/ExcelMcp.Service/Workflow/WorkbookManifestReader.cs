using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Service.Workflow;

/// <summary>
/// Reads a compact, point-in-time workbook manifest in one Excel STA callback.
/// It intentionally keeps no cache: every result reflects the currently opened workbook.
/// </summary>
internal static class WorkbookManifestReader
{
    internal const int DefaultPreviewRows = 3;
    internal const int DefaultPreviewColumns = 3;
    internal const int MaximumPreviewRows = 20;
    internal const int MaximumPreviewColumns = 12;
    private const int MaximumPreviewStringLength = 200;

    internal static WorkbookManifest Read(
        IExcelBatch batch,
        string sessionId,
        int previewRows,
        int previewColumns)
    {
        return batch.Execute((context, cancellationToken) =>
        {
            dynamic? worksheets = null;
            try
            {
                worksheets = context.Book.Worksheets;
                int worksheetCount = Convert.ToInt32(worksheets.Count, CultureInfo.InvariantCulture);
                var sheets = new List<WorkbookSheetManifest>(worksheetCount);

                for (int index = 1; index <= worksheetCount; index++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    dynamic? worksheet = null;
                    dynamic? usedRange = null;
                    dynamic? usedRows = null;
                    dynamic? usedColumns = null;
                    dynamic? previewRange = null;
                    try
                    {
                        worksheet = worksheets.Item(index);
                        usedRange = worksheet.UsedRange;
                        usedRows = usedRange.Rows;
                        usedColumns = usedRange.Columns;

                        int rowCount = Convert.ToInt32(usedRows.Count, CultureInfo.InvariantCulture);
                        int columnCount = Convert.ToInt32(usedColumns.Count, CultureInfo.InvariantCulture);
                        int returnedRows = Math.Min(previewRows, rowCount);
                        int returnedColumns = Math.Min(previewColumns, columnCount);
                        previewRange = usedRange.Resize[returnedRows, returnedColumns];

                        sheets.Add(new WorkbookSheetManifest
                        {
                            Name = Convert.ToString(worksheet.Name, CultureInfo.InvariantCulture) ?? string.Empty,
                            Index = index,
                            Visible = Convert.ToInt32(worksheet.Visible, CultureInfo.InvariantCulture) == -1,
                            UsedRange = Convert.ToString(usedRange.Address, CultureInfo.InvariantCulture) ?? string.Empty,
                            RowCount = rowCount,
                            ColumnCount = columnCount,
                            Preview = ReadPreview(previewRange.Value2, returnedRows, returnedColumns),
                        });
                    }
                    finally
                    {
                        ComUtilities.Release(ref previewRange);
                        ComUtilities.Release(ref usedColumns);
                        ComUtilities.Release(ref usedRows);
                        ComUtilities.Release(ref usedRange);
                        ComUtilities.Release(ref worksheet);
                    }
                }

                return new WorkbookManifest
                {
                    Success = true,
                    SessionId = sessionId,
                    FilePath = batch.WorkbookPath,
                    SheetCount = worksheetCount,
                    PreviewRows = previewRows,
                    PreviewColumns = previewColumns,
                    Sheets = sheets,
                };
            }
            finally
            {
                ComUtilities.Release(ref worksheets);
            }
        });
    }

    private static List<List<object?>> ReadPreview(object? rawValue, int rowCount, int columnCount)
    {
        var preview = new List<List<object?>>(rowCount);
        if (rawValue is Array array && array.Rank == 2)
        {
            int rowLowerBound = array.GetLowerBound(0);
            int columnLowerBound = array.GetLowerBound(1);
            for (int row = 0; row < rowCount; row++)
            {
                var values = new List<object?>(columnCount);
                for (int column = 0; column < columnCount; column++)
                {
                    values.Add(NormalizeCellValue(array.GetValue(row + rowLowerBound, column + columnLowerBound)));
                }

                preview.Add(values);
            }

            return preview;
        }

        preview.Add([NormalizeCellValue(rawValue)]);
        return preview;
    }

    private static object? NormalizeCellValue(object? value)
    {
        if (value is null or DBNull)
        {
            return null;
        }

        if (value is string text && text.Length > MaximumPreviewStringLength)
        {
            return text[..(MaximumPreviewStringLength - 1)] + "…";
        }

        return value is string or bool or byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal
            ? value
            : Convert.ToString(value, CultureInfo.InvariantCulture);
    }
}

internal sealed class WorkbookManifest
{
    public bool Success { get; init; }
    public required string SessionId { get; init; }
    public required string FilePath { get; init; }
    public int SheetCount { get; init; }
    public int PreviewRows { get; init; }
    public int PreviewColumns { get; init; }
    public required IReadOnlyList<WorkbookSheetManifest> Sheets { get; init; }
}

internal sealed class WorkbookSheetManifest
{
    public required string Name { get; init; }
    public int Index { get; init; }
    public bool Visible { get; init; }
    public required string UsedRange { get; init; }
    public int RowCount { get; init; }
    public int ColumnCount { get; init; }
    public required IReadOnlyList<IReadOnlyList<object?>> Preview { get; init; }
}
