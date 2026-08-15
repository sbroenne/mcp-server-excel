using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Drawing;

public sealed partial class DrawingCommands
{
    /// <inheritdoc />
    public SparklineListResult ListSparklines(IExcelBatch batch, string sheetName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cells = null;
            Excel.SparklineGroups? groups = null;
            try
            {
                sheet = GetSheet(ctx.Book, sheetName);
                cells = sheet.Cells;
                groups = cells.SparklineGroups;
                var result = new SparklineListResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath
                };

                for (var index = 1; index <= groups.Count; index++)
                {
                    Excel.SparklineGroup? group = null;
                    try
                    {
                        group = groups[index];
                        result.Sparklines.Add(ReadSparkline(group, sheetName));
                    }
                    finally
                    {
                        ComUtilities.Release(ref group);
                    }
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref groups);
                ComUtilities.Release(ref cells);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public SparklineResult GetSparkline(IExcelBatch batch, string sheetName, string locationRange)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cells = null;
            Excel.SparklineGroups? groups = null;
            Excel.SparklineGroup? group = null;
            try
            {
                sheet = GetSheet(ctx.Book, sheetName);
                cells = sheet.Cells;
                groups = cells.SparklineGroups;
                group = FindSparklineGroup(groups, locationRange)
                    ?? throw new InvalidOperationException($"Sparkline at '{locationRange}' not found on sheet '{sheetName}'.");
                return CreateSparklineResult(batch.WorkbookPath, ReadSparkline(group, sheetName));
            }
            finally
            {
                ComUtilities.Release(ref group);
                ComUtilities.Release(ref groups);
                ComUtilities.Release(ref cells);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public SparklineResult AddSparkline(
        IExcelBatch batch,
        string sheetName,
        string sourceRange,
        string locationRange,
        DrawingSparklineType sparklineType = DrawingSparklineType.Line,
        string? lineColor = null,
        bool showMarkers = false)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? location = null;
            Excel.SparklineGroups? groups = null;
            Excel.SparklineGroup? group = null;
            try
            {
                sheet = GetSheet(ctx.Book, sheetName);
                location = sheet.Range[locationRange];
                groups = location.SparklineGroups;
                if (groups.Count > 0)
                {
                    throw new InvalidOperationException($"Range '{locationRange}' already contains a sparkline.");
                }

                group = groups.Add((Excel.XlSparkType)sparklineType, QualifyRange(sheetName, sourceRange));
                ApplySparklineFormatting(group, lineColor, showMarkers);
                return CreateSparklineResult(batch.WorkbookPath, ReadSparkline(group, sheetName));
            }
            finally
            {
                ComUtilities.Release(ref group);
                ComUtilities.Release(ref groups);
                ComUtilities.Release(ref location);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public SparklineResult UpdateSparkline(
        IExcelBatch batch,
        string sheetName,
        string locationRange,
        string? sourceRange = null,
        DrawingSparklineType? sparklineType = null,
        string? lineColor = null,
        bool? showMarkers = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cells = null;
            Excel.SparklineGroups? groups = null;
            Excel.SparklineGroup? group = null;
            try
            {
                sheet = GetSheet(ctx.Book, sheetName);
                cells = sheet.Cells;
                groups = cells.SparklineGroups;
                group = FindSparklineGroup(groups, locationRange)
                    ?? throw new InvalidOperationException($"Sparkline at '{locationRange}' not found on sheet '{sheetName}'.");

                if (sourceRange != null)
                {
                    group.ModifySourceData(QualifyRange(sheetName, sourceRange));
                }

                if (sparklineType.HasValue)
                {
                    group.Type = (Excel.XlSparkType)sparklineType.Value;
                }

                ApplySparklineFormatting(group, lineColor, showMarkers);
                return CreateSparklineResult(batch.WorkbookPath, ReadSparkline(group, sheetName));
            }
            finally
            {
                ComUtilities.Release(ref group);
                ComUtilities.Release(ref groups);
                ComUtilities.Release(ref cells);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult DeleteSparkline(IExcelBatch batch, string sheetName, string locationRange)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cells = null;
            Excel.SparklineGroups? groups = null;
            Excel.SparklineGroup? group = null;
            try
            {
                sheet = GetSheet(ctx.Book, sheetName);
                cells = sheet.Cells;
                groups = cells.SparklineGroups;
                group = FindSparklineGroup(groups, locationRange)
                    ?? throw new InvalidOperationException($"Sparkline at '{locationRange}' not found on sheet '{sheetName}'.");
                group.Delete();
                return new OperationResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    Action = "delete-sparkline"
                };
            }
            finally
            {
                ComUtilities.Release(ref group);
                ComUtilities.Release(ref groups);
                ComUtilities.Release(ref cells);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static Excel.SparklineGroup? FindSparklineGroup(Excel.SparklineGroups groups, string locationRange)
    {
        var normalizedTarget = NormalizeRange(locationRange);
        for (var index = 1; index <= groups.Count; index++)
        {
            Excel.SparklineGroup? group = null;
            Excel.Range? location = null;
            try
            {
                group = groups[index];
                location = group.Location;
                if (string.Equals(GetRangeAddress(location), normalizedTarget, StringComparison.OrdinalIgnoreCase))
                {
                    var found = group;
                    group = null;
                    return found;
                }
            }
            finally
            {
                ComUtilities.Release(ref location);
                ComUtilities.Release(ref group);
            }
        }

        return null;
    }

    private static SparklineInfo ReadSparkline(Excel.SparklineGroup group, string sheetName)
    {
        Excel.Range? location = null;
        Excel.FormatColor? seriesColor = null;
        Excel.SparkPoints? points = null;
        Excel.SparkColor? markers = null;
        try
        {
            location = group.Location;
            seriesColor = group.SeriesColor;
            points = group.Points;
            markers = points.Markers;
            return new SparklineInfo
            {
                SheetName = sheetName,
                LocationRange = GetRangeAddress(location),
                SourceRange = NormalizeRange(group.SourceData),
                SparklineType = (DrawingSparklineType)group.Type,
                LineColor = FormatColor(Convert.ToInt32(seriesColor.Color)),
                ShowMarkers = markers.Visible
            };
        }
        finally
        {
            ComUtilities.Release(ref markers);
            ComUtilities.Release(ref points);
            ComUtilities.Release(ref seriesColor);
            ComUtilities.Release(ref location);
        }
    }

    private static void ApplySparklineFormatting(Excel.SparklineGroup group, string? lineColor, bool? showMarkers)
    {
        Excel.FormatColor? seriesColor = null;
        Excel.SparkPoints? points = null;
        Excel.SparkColor? markers = null;
        try
        {
            if (lineColor != null)
            {
                seriesColor = group.SeriesColor;
                seriesColor.Color = ParseColor(lineColor);
            }

            if (showMarkers.HasValue)
            {
                points = group.Points;
                markers = points.Markers;
                markers.Visible = showMarkers.Value;
            }
        }
        finally
        {
            ComUtilities.Release(ref markers);
            ComUtilities.Release(ref points);
            ComUtilities.Release(ref seriesColor);
        }
    }

    private static SparklineResult CreateSparklineResult(string filePath, SparklineInfo sparkline)
    {
        return new SparklineResult
        {
            Success = true,
            FilePath = filePath,
            Sparkline = sparkline
        };
    }

    private static string QualifyRange(string sheetName, string range)
    {
        var value = range.Trim().TrimStart('=');
        if (value.Contains('!'))
        {
            return value;
        }

        return $"'{sheetName.Replace("'", "''", StringComparison.Ordinal)}'!{value}";
    }

    private static string GetRangeAddress(Excel.Range range)
    {
        return NormalizeRange(range.Address[false, false, Excel.XlReferenceStyle.xlA1]);
    }

    private static string NormalizeRange(string range)
    {
        var value = range.Trim().TrimStart('=');
        var separator = value.LastIndexOf('!');
        if (separator >= 0)
        {
            value = value[(separator + 1)..];
        }

        return value.Replace("$", string.Empty, StringComparison.Ordinal);
    }
}
