using System.Globalization;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using Microsoft.CSharp.RuntimeBinder;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.ServiceClient;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Service.Workflow;

/// <summary>
/// Reads one caller-selected rectangular range after a workflow plan. The verifier never
/// uses selection or UsedRange, never returns an unbounded matrix, and makes partial
/// inspection explicit in the receipt.
/// </summary>
internal static class WorkflowRangeVerifier
{
    internal const int MaximumInspectedCells = 10_000;
    internal const int PreviewRowCount = 2;
    internal const int PreviewColumnCount = 4;
    private const int MaximumPreviewStringLength = 200;

    internal static WorkflowRangeVerificationReceipt Read(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(rangeAddress);

        return batch.Execute((context, cancellationToken) =>
        {
            dynamic? worksheets = null;
            dynamic? worksheet = null;
            dynamic? requestedRange = null;
            dynamic? areas = null;
            dynamic? rows = null;
            dynamic? columns = null;
            dynamic? inspectedRange = null;
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                worksheets = context.Book.Worksheets;
                worksheet = worksheets.Item(sheetName);
                requestedRange = worksheet.Range[rangeAddress];
                areas = requestedRange.Areas;
                int areaCount = Convert.ToInt32(areas.Count, CultureInfo.InvariantCulture);
                if (areaCount != 1)
                {
                    throw new ArgumentException(
                        $"Verification range '{rangeAddress}' contains {areaCount} disjoint areas; one rectangular area is required.",
                        nameof(rangeAddress));
                }

                rows = requestedRange.Rows;
                columns = requestedRange.Columns;
                int rowCount = Convert.ToInt32(rows.Count, CultureInfo.InvariantCulture);
                int columnCount = Convert.ToInt32(columns.Count, CultureInfo.InvariantCulture);
                long cellCount = checked((long)rowCount * columnCount);

                int inspectedRows = Math.Min(rowCount, MaximumInspectedCells);
                int inspectedColumns = Math.Min(
                    columnCount,
                    Math.Max(1, MaximumInspectedCells / inspectedRows));
                int inspectedCellCount = checked(inspectedRows * inspectedColumns);
                bool complete = cellCount == inspectedCellCount;

                inspectedRange = requestedRange.Resize[inspectedRows, inspectedColumns];
                string canonicalAddress = Convert.ToString(requestedRange.Address, CultureInfo.InvariantCulture)
                    ?? rangeAddress;
                string inspectedAddress = Convert.ToString(inspectedRange.Address, CultureInfo.InvariantCulture)
                    ?? canonicalAddress;
                object? rawValues = inspectedRange.Value2;
                object? rawFormulas = ReadFormulas(inspectedRange);

                var preview = new List<IReadOnlyList<object?>>(Math.Min(PreviewRowCount, inspectedRows));
                int nonEmptyCellCount = 0;
                int formulaCellCount = 0;
                using var fingerprint = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
                AppendToken(fingerprint, "schema", "workflow-range-verification-v1");
                AppendToken(fingerprint, "sheet", sheetName);
                AppendToken(fingerprint, "range", canonicalAddress);
                AppendToken(fingerprint, "inspectedRange", inspectedAddress);
                AppendToken(fingerprint, "rows", inspectedRows.ToString(CultureInfo.InvariantCulture));
                AppendToken(fingerprint, "columns", inspectedColumns.ToString(CultureInfo.InvariantCulture));

                for (int row = 0; row < inspectedRows; row++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    List<object?>? previewRow = row < PreviewRowCount
                        ? new List<object?>(Math.Min(PreviewColumnCount, inspectedColumns))
                        : null;
                    for (int column = 0; column < inspectedColumns; column++)
                    {
                        object? value = GetCell(rawValues, row, column);
                        string? formula = GetFormula(GetCell(rawFormulas, row, column));
                        if (formula is not null)
                        {
                            formulaCellCount++;
                        }

                        if (formula is not null || IsNonEmpty(value))
                        {
                            nonEmptyCellCount++;
                        }

                        AppendScalar(fingerprint, "value", value);
                        AppendToken(fingerprint, "formula", formula ?? string.Empty);
                        if (previewRow is not null && column < PreviewColumnCount)
                        {
                            previewRow.Add(NormalizePreviewValue(value));
                        }
                    }

                    if (previewRow is not null)
                    {
                        preview.Add(previewRow);
                    }
                }

                return new WorkflowRangeVerificationReceipt
                {
                    Status = complete ? "verified" : "partiallyVerified",
                    SheetName = sheetName,
                    RangeAddress = canonicalAddress,
                    RowCount = rowCount,
                    ColumnCount = columnCount,
                    CellCount = cellCount,
                    InspectedCellCount = inspectedCellCount,
                    InspectedRangeAddress = inspectedAddress,
                    NonEmptyCellCount = nonEmptyCellCount,
                    FormulaCellCount = formulaCellCount,
                    Fingerprint = Convert.ToHexStringLower(fingerprint.GetHashAndReset()),
                    Preview = preview,
                    Limitation = complete
                        ? null
                        : $"Only the top-left {inspectedCellCount.ToString(CultureInfo.InvariantCulture)} of {cellCount.ToString(CultureInfo.InvariantCulture)} cells were inspected; counts and fingerprint cover that sample.",
                };
            }
            finally
            {
                ComUtilities.Release(ref inspectedRange);
                ComUtilities.Release(ref columns);
                ComUtilities.Release(ref rows);
                ComUtilities.Release(ref areas);
                ComUtilities.Release(ref requestedRange);
                ComUtilities.Release(ref worksheet);
                ComUtilities.Release(ref worksheets);
            }
        });
    }

    internal static WorkflowRangeVerificationReceipt NotVerified(
        string sheetName,
        string rangeAddress,
        string limitation) => new()
        {
            Status = "notVerified",
            SheetName = sheetName,
            RangeAddress = rangeAddress,
            Limitation = limitation,
        };

    private static object? ReadFormulas(dynamic range)
    {
        try
        {
            return range.Formula2;
        }
        catch (Exception ex) when (ex is COMException or RuntimeBinderException)
        {
            return range.Formula;
        }
    }

    private static object? GetCell(object? raw, int row, int column)
    {
        if (raw is Array array && array.Rank == 2)
        {
            if (row >= array.GetLength(0) || column >= array.GetLength(1))
            {
                return null;
            }

            return array.GetValue(
                row + array.GetLowerBound(0),
                column + array.GetLowerBound(1));
        }

        return row == 0 && column == 0 ? raw : null;
    }

    private static string? GetFormula(object? rawFormula) =>
        rawFormula is string text && text.StartsWith('=') ? text : null;

    private static bool IsNonEmpty(object? value) => value switch
    {
        null or DBNull => false,
        string text => text.Length > 0,
        _ => true,
    };

    private static object? NormalizePreviewValue(object? value)
    {
        if (value is null or DBNull)
        {
            return null;
        }

        if (value is string text)
        {
            return text.Length <= MaximumPreviewStringLength
                ? text
                : text[..(MaximumPreviewStringLength - 3)] + "...";
        }

        if (value is ErrorWrapper error)
        {
            return $"#COMERROR:{error.ErrorCode:X8}";
        }

        return value is bool or byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal
            ? value
            : Convert.ToString(value, CultureInfo.InvariantCulture);
    }

    private static void AppendScalar(IncrementalHash hash, string name, object? value)
    {
        switch (value)
        {
            case null:
            case DBNull:
                AppendToken(hash, name + ":null", string.Empty);
                return;
            case string text:
                AppendToken(hash, name + ":string", text);
                return;
            case bool boolean:
                AppendToken(hash, name + ":bool", boolean ? "true" : "false");
                return;
            case DateTime dateTime:
                AppendToken(hash, name + ":datetime", dateTime.ToString("O", CultureInfo.InvariantCulture));
                return;
            case ErrorWrapper error:
                AppendToken(hash, name + ":error", error.ErrorCode.ToString("X8", CultureInfo.InvariantCulture));
                return;
            case IFormattable formattable:
                AppendToken(hash, name + ":" + value.GetType().FullName, formattable.ToString(null, CultureInfo.InvariantCulture));
                return;
            default:
                AppendToken(
                    hash,
                    name + ":" + value.GetType().FullName,
                    Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
                return;
        }
    }

    private static void AppendToken(IncrementalHash hash, string name, string value)
    {
        byte[] header = Encoding.UTF8.GetBytes(
            name + ":" + Encoding.UTF8.GetByteCount(value).ToString(CultureInfo.InvariantCulture) + ":");
        byte[] payload = Encoding.UTF8.GetBytes(value);
        hash.AppendData(header);
        hash.AppendData(payload);
        hash.AppendData([0]);
    }
}
