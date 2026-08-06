using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Core.Commands.ReportFormat;

/// <summary>Excel COM implementation of deterministic report-format presets and readback.</summary>
public sealed class ReportFormatCommands : IReportFormatCommands
{
    private const int XlCenter = -4108;
    private const int XlLeft = -4131;
    private const int XlGeneral = 1;
    private const int XlContinuous = 1;
    private const int XlThin = 2;
    private const int XlUnderlineNone = -4142;
    private const int MaximumAutoFitColumnWidth = 40;
    private static readonly int[] BorderIndexes = [7, 8, 9, 10, 11, 12];

    /// <inheritdoc />
    public ReportFormatStateResult Apply(
        IExcelBatch batch,
        string sheetName,
        string? titleRange,
        string headerRange,
        string bodyRange,
        string? totalRange,
        ReportFormatPreset preset = ReportFormatPreset.Professional,
        string accentColor = "#1F4E78",
        bool autoFitColumns = true)
    {
        ValidateArguments(sheetName, titleRange, headerRange, bodyRange, totalRange);
        var normalizedAccent = NormalizeColor(accentColor);
        var accent = FormattingHelpers.ParseColor(normalizedAccent);
        var lightAccent = Lighten(accent, 0.82);

        return batch.Execute((ctx, ct) =>
        {
            object? sheet = null;
            List<ResolvedSection>? sections = null;

            try
            {
                sheet = ctx.Book.Worksheets[sheetName];
                sections = ResolveAndValidateSections(sheet, titleRange, headerRange, bodyRange, totalRange);

                foreach (var section in sections)
                {
                    ApplySectionFormat(section, preset, accent, lightAccent);
                }

                if (autoFitColumns)
                {
                    AutoFitAndCapColumns(sections.Single(section => section.Name == "header"));
                }

                var state = ReadState(batch, sheetName, sections, "apply", preset, normalizedAccent, autoFitColumns);
                state.Message = $"Applied {preset.ToString().ToLowerInvariant()} report formatting to {sections.Count} explicit sections";
                return state;
            }
            finally
            {
                ReleaseSections(sections);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public ReportFormatStateResult GetState(
        IExcelBatch batch,
        string sheetName,
        string? titleRange,
        string headerRange,
        string bodyRange,
        string? totalRange)
    {
        ValidateArguments(sheetName, titleRange, headerRange, bodyRange, totalRange);

        return batch.Execute((ctx, ct) =>
        {
            object? sheet = null;
            List<ResolvedSection>? sections = null;

            try
            {
                sheet = ctx.Book.Worksheets[sheetName];
                sections = ResolveAndValidateSections(sheet, titleRange, headerRange, bodyRange, totalRange);
                var state = ReadState(batch, sheetName, sections, "get-state", null, string.Empty, false);
                state.AccentColor = state.Sections.Single(section => section.Name == "header").FillColor ?? string.Empty;
                state.Message = $"Read report formatting from {sections.Count} explicit sections";
                return state;
            }
            finally
            {
                ReleaseSections(sections);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static void ValidateArguments(
        string sheetName,
        string? titleRange,
        string headerRange,
        string bodyRange,
        string? totalRange)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
            throw new ArgumentException("sheetName is required", nameof(sheetName));
        if (string.IsNullOrWhiteSpace(headerRange))
            throw new ArgumentException("headerRange is required", nameof(headerRange));
        if (string.IsNullOrWhiteSpace(bodyRange))
            throw new ArgumentException("bodyRange is required", nameof(bodyRange));
        if (titleRange is not null && string.IsNullOrWhiteSpace(titleRange))
            throw new ArgumentException("titleRange cannot be empty when supplied", nameof(titleRange));
        if (totalRange is not null && string.IsNullOrWhiteSpace(totalRange))
            throw new ArgumentException("totalRange cannot be empty when supplied", nameof(totalRange));
    }

    private static string NormalizeColor(string accentColor)
    {
        if (string.IsNullOrWhiteSpace(accentColor))
            throw new ArgumentException("accentColor is required", nameof(accentColor));

        var parsed = FormattingHelpers.ParseColor(accentColor);
        return FormattingHelpers.ColorToHex(parsed);
    }

    private static int Lighten(int excelColor, double whiteRatio)
    {
        var red = excelColor & 0xFF;
        var green = (excelColor >> 8) & 0xFF;
        var blue = (excelColor >> 16) & 0xFF;
        red = (int)Math.Round(red + ((255 - red) * whiteRatio), MidpointRounding.AwayFromZero);
        green = (int)Math.Round(green + ((255 - green) * whiteRatio), MidpointRounding.AwayFromZero);
        blue = (int)Math.Round(blue + ((255 - blue) * whiteRatio), MidpointRounding.AwayFromZero);
        return red + (green << 8) + (blue << 16);
    }

    private static List<ResolvedSection> ResolveAndValidateSections(
        object sheetObject,
        string? titleRange,
        string headerRange,
        string bodyRange,
        string? totalRange)
    {
        dynamic sheet = sheetObject;
        var sections = new List<ResolvedSection>();

        try
        {
            if (titleRange is not null) sections.Add(ResolveSection(sheet, "title", titleRange));
            sections.Add(ResolveSection(sheet, "header", headerRange));
            sections.Add(ResolveSection(sheet, "body", bodyRange));
            if (totalRange is not null) sections.Add(ResolveSection(sheet, "total", totalRange));

            var header = sections.Single(section => section.Name == "header");
            foreach (var section in sections)
            {
                if (section.FirstColumn != header.FirstColumn || section.ColumnCount != header.ColumnCount)
                {
                    throw new ArgumentException(
                        $"Section '{section.Name}' must span the same columns as header range {header.Address}");
                }
            }

            for (var index = 1; index < sections.Count; index++)
            {
                if (sections[index - 1].LastRow >= sections[index].FirstRow)
                {
                    throw new ArgumentException(
                        $"Report sections must be non-overlapping and ordered title/header/body/total; " +
                        $"'{sections[index - 1].Name}' overlaps or follows '{sections[index].Name}'");
                }
            }

            return sections;
        }
        catch
        {
            ReleaseSections(sections);
            throw;
        }
    }

    private static ResolvedSection ResolveSection(dynamic sheet, string name, string address)
    {
        object? range = null;
        object? columns = null;
        object? rows = null;

        try
        {
            range = sheet.Range[address];
            dynamic dynamicRange = range;
            columns = dynamicRange.Columns;
            rows = dynamicRange.Rows;
            var firstColumn = Convert.ToInt32(dynamicRange.Column, CultureInfo.InvariantCulture);
            var firstRow = Convert.ToInt32(dynamicRange.Row, CultureInfo.InvariantCulture);
            dynamic dynamicColumns = columns;
            dynamic dynamicRows = rows;
            var columnCount = Convert.ToInt32(dynamicColumns.Count, CultureInfo.InvariantCulture);
            var rowCount = Convert.ToInt32(dynamicRows.Count, CultureInfo.InvariantCulture);
            var normalizedAddress = Convert.ToString(dynamicRange.Address, CultureInfo.InvariantCulture)
                ?? throw new InvalidOperationException($"Excel returned no address for report section '{name}'");

            var resolved = new ResolvedSection(
                name,
                normalizedAddress,
                range,
                firstColumn,
                columnCount,
                firstRow,
                rowCount);
            range = null;
            return resolved;
        }
        catch (Exception ex) when (ex is not ArgumentException)
        {
            throw new ArgumentException($"Invalid {name} range '{address}'", name + "Range", ex);
        }
        finally
        {
            ComUtilities.Release(ref rows);
            ComUtilities.Release(ref columns);
            ComUtilities.Release(ref range);
        }
    }

    private static void ApplySectionFormat(
        ResolvedSection section,
        ReportFormatPreset preset,
        int accent,
        int lightAccent)
    {
        dynamic range = section.Range;
        object? font = null;
        object? interior = null;
        object? rows = null;

        try
        {
            font = range.Font;
            interior = range.Interior;
            rows = range.Rows;
            dynamic dynamicFont = font;
            dynamic dynamicInterior = interior;
            dynamic dynamicRows = rows;

            dynamicFont.Name = "Calibri";
            dynamicFont.Italic = false;
            dynamicFont.Underline = XlUnderlineNone;
            range.VerticalAlignment = XlCenter;

            var isMinimal = preset == ReportFormatPreset.Minimal;
            switch (section.Name)
            {
                case "title":
                    dynamicFont.Size = 16d;
                    dynamicFont.Bold = true;
                    dynamicFont.Color = isMinimal ? accent : FormattingHelpers.ParseColor("#FFFFFF");
                    dynamicInterior.Color = isMinimal ? FormattingHelpers.ParseColor("#FFFFFF") : accent;
                    range.HorizontalAlignment = XlLeft;
                    range.WrapText = false;
                    dynamicRows.RowHeight = 24d;
                    ApplyBorders(range, isMinimal ? accent : lightAccent, XlThin);
                    break;

                case "header":
                    dynamicFont.Size = 11d;
                    dynamicFont.Bold = true;
                    dynamicFont.Color = isMinimal ? accent : FormattingHelpers.ParseColor("#FFFFFF");
                    dynamicInterior.Color = isMinimal ? FormattingHelpers.ParseColor("#FFFFFF") : accent;
                    range.HorizontalAlignment = XlCenter;
                    range.WrapText = true;
                    dynamicRows.RowHeight = 20d;
                    ApplyBorders(range, accent, XlThin);
                    break;

                case "body":
                    dynamicFont.Size = 11d;
                    dynamicFont.Bold = false;
                    dynamicFont.Color = FormattingHelpers.ParseColor("#000000");
                    dynamicInterior.Color = FormattingHelpers.ParseColor("#FFFFFF");
                    range.HorizontalAlignment = XlGeneral;
                    range.WrapText = false;
                    ApplyBorders(range, isMinimal ? FormattingHelpers.ParseColor("#D9D9D9") : lightAccent, XlThin);
                    break;

                case "total":
                    dynamicFont.Size = 11d;
                    dynamicFont.Bold = true;
                    dynamicFont.Color = FormattingHelpers.ParseColor("#000000");
                    dynamicInterior.Color = isMinimal ? FormattingHelpers.ParseColor("#FFFFFF") : lightAccent;
                    range.HorizontalAlignment = XlGeneral;
                    range.WrapText = false;
                    ApplyBorders(range, accent, XlThin);
                    break;

                default:
                    throw new InvalidOperationException($"Unknown report section: {section.Name}");
            }
        }
        finally
        {
            ComUtilities.Release(ref rows);
            ComUtilities.Release(ref interior);
            ComUtilities.Release(ref font);
        }
    }

    private static void ApplyBorders(dynamic range, int color, int weight)
    {
        object? borders = null;
        try
        {
            borders = range.Borders;
            dynamic dynamicBorders = borders;
            foreach (var borderIndex in BorderIndexes)
            {
                object? border = null;
                try
                {
                    border = dynamicBorders.Item(borderIndex);
                    dynamic dynamicBorder = border;
                    dynamicBorder.LineStyle = XlContinuous;
                    dynamicBorder.Color = color;
                    dynamicBorder.Weight = weight;
                }
                finally
                {
                    ComUtilities.Release(ref border);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref borders);
        }
    }

    private static void AutoFitAndCapColumns(ResolvedSection header)
    {
        dynamic range = header.Range;
        object? columns = null;
        try
        {
            columns = range.Columns;
            dynamic dynamicColumns = columns;
            dynamicColumns.AutoFit();
            var count = Convert.ToInt32(dynamicColumns.Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                object? column = null;
                try
                {
                    column = dynamicColumns.Item(index);
                    dynamic dynamicColumn = column;
                    var width = Convert.ToDouble(dynamicColumn.ColumnWidth, CultureInfo.InvariantCulture);
                    if (width > MaximumAutoFitColumnWidth)
                    {
                        dynamicColumn.ColumnWidth = MaximumAutoFitColumnWidth;
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
            ComUtilities.Release(ref columns);
        }
    }

    private static ReportFormatStateResult ReadState(
        IExcelBatch batch,
        string sheetName,
        IReadOnlyList<ResolvedSection> sections,
        string action,
        ReportFormatPreset? preset,
        string accentColor,
        bool autoFitColumns)
    {
        var sectionStates = sections.Select(ReadSectionState).ToList();
        var fingerprintSource = JsonSerializer.Serialize(sectionStates);
        var fingerprint = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(fingerprintSource))).ToLowerInvariant();

        return new ReportFormatStateResult
        {
            Success = true,
            FilePath = batch.WorkbookPath,
            Action = action,
            SheetName = sheetName,
            Preset = preset?.ToString().ToLowerInvariant(),
            AccentColor = accentColor,
            AutoFitColumns = autoFitColumns,
            Sections = sectionStates,
            Fingerprint = fingerprint,
        };
    }

    private static ReportFormatSectionState ReadSectionState(ResolvedSection section)
    {
        dynamic range = section.Range;
        object? font = null;
        object? interior = null;
        object? borders = null;
        object? bottomBorder = null;

        try
        {
            font = range.Font;
            interior = range.Interior;
            borders = range.Borders;
            dynamic dynamicBorders = borders;
            bottomBorder = dynamicBorders.Item(9);
            dynamic dynamicFont = font;
            dynamic dynamicInterior = interior;
            dynamic dynamicBorder = bottomBorder;

            return new ReportFormatSectionState
            {
                Name = section.Name,
                RangeAddress = section.Address,
                FontName = ReadString(dynamicFont.Name),
                FontSize = ReadDouble(dynamicFont.Size),
                Bold = ReadBoolean(dynamicFont.Bold),
                Italic = ReadBoolean(dynamicFont.Italic),
                FontColor = ReadColor(dynamicFont.Color),
                FillColor = ReadColor(dynamicInterior.Color),
                HorizontalAlignment = ReadAlignment(range.HorizontalAlignment, vertical: false),
                VerticalAlignment = ReadAlignment(range.VerticalAlignment, vertical: true),
                WrapText = ReadBoolean(range.WrapText),
                NumberFormat = ReadString(range.NumberFormat),
                BorderLineStyle = ReadInteger(dynamicBorder.LineStyle),
                BorderColor = ReadColor(dynamicBorder.Color),
            };
        }
        finally
        {
            ComUtilities.Release(ref bottomBorder);
            ComUtilities.Release(ref borders);
            ComUtilities.Release(ref interior);
            ComUtilities.Release(ref font);
        }
    }

    private static string? ReadString(object? value) =>
        value is null or DBNull ? null : Convert.ToString(value, CultureInfo.InvariantCulture);

    private static double? ReadDouble(object? value) =>
        value is null or DBNull ? null : Convert.ToDouble(value, CultureInfo.InvariantCulture);

    private static int? ReadInteger(object? value) =>
        value is null or DBNull ? null : Convert.ToInt32(value, CultureInfo.InvariantCulture);

    private static bool? ReadBoolean(object? value) =>
        value is null or DBNull ? null : Convert.ToBoolean(value, CultureInfo.InvariantCulture);

    private static string? ReadColor(object? value) =>
        ReadInteger(value) is { } color ? FormattingHelpers.ColorToHex(color) : null;

    private static string? ReadAlignment(object? value, bool vertical)
    {
        var alignment = ReadInteger(value);
        return alignment switch
        {
            null => null,
            XlCenter => "center",
            XlLeft when !vertical => "left",
            XlGeneral when !vertical => "general",
            -4152 when !vertical => "right",
            -4160 when vertical => "top",
            -4107 when vertical => "bottom",
            _ => alignment.Value.ToString(CultureInfo.InvariantCulture),
        };
    }

    private static void ReleaseSections(IEnumerable<ResolvedSection>? sections)
    {
        if (sections is null) return;

        foreach (var section in sections)
        {
            var range = section.Range;
            ComUtilities.Release(ref range);
        }
    }

    private sealed class ResolvedSection(
        string name,
        string address,
        object range,
        int firstColumn,
        int columnCount,
        int firstRow,
        int rowCount)
    {
        internal string Name { get; } = name;
        internal string Address { get; } = address;
        internal object Range { get; } = range;
        internal int FirstColumn { get; } = firstColumn;
        internal int ColumnCount { get; } = columnCount;
        internal int FirstRow { get; } = firstRow;
        internal int LastRow { get; } = firstRow + rowCount - 1;
    }
}
