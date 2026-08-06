using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.ReportFormat;

/// <summary>Deterministic built-in report formatting presets.</summary>
public enum ReportFormatPreset
{
    /// <summary>Dark accent title/header, clean body grid, emphasized total row.</summary>
    Professional,

    /// <summary>White background, accent text, restrained borders.</summary>
    Minimal,
}

/// <summary>Normalized readback for one explicitly addressed report section.</summary>
public sealed class ReportFormatSectionState
{
    /// <summary>Logical section name: title, header, body, or total.</summary>
    public string Name { get; set; } = string.Empty;
    /// <summary>Canonical Excel A1 address.</summary>
    public string RangeAddress { get; set; } = string.Empty;
    /// <summary>Uniform font family, or null when mixed.</summary>
    public string? FontName { get; set; }
    /// <summary>Uniform font size, or null when mixed.</summary>
    public double? FontSize { get; set; }
    /// <summary>Uniform bold state, or null when mixed.</summary>
    public bool? Bold { get; set; }
    /// <summary>Uniform italic state, or null when mixed.</summary>
    public bool? Italic { get; set; }
    /// <summary>Uniform font color as #RRGGBB, or null when mixed.</summary>
    public string? FontColor { get; set; }
    /// <summary>Uniform fill color as #RRGGBB, or null when mixed.</summary>
    public string? FillColor { get; set; }
    /// <summary>Normalized horizontal alignment, or null when mixed.</summary>
    public string? HorizontalAlignment { get; set; }
    /// <summary>Normalized vertical alignment, or null when mixed.</summary>
    public string? VerticalAlignment { get; set; }
    /// <summary>Uniform wrap-text state, or null when mixed.</summary>
    public bool? WrapText { get; set; }
    /// <summary>Uniform number format, or null when mixed. Apply preserves this value.</summary>
    public string? NumberFormat { get; set; }
    /// <summary>Bottom-border Excel line-style value, or null when mixed.</summary>
    public int? BorderLineStyle { get; set; }
    /// <summary>Bottom-border color as #RRGGBB, or null when mixed.</summary>
    public string? BorderColor { get; set; }
}

/// <summary>Applied or inspected report-format state with a deterministic fingerprint.</summary>
public sealed class ReportFormatStateResult : OperationResult
{
    /// <summary>Worksheet containing every section.</summary>
    public string SheetName { get; set; } = string.Empty;
    /// <summary>Applied preset, or null for inspection-only readback.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Preset { get; set; }
    /// <summary>Normalized requested/read accent color.</summary>
    public string AccentColor { get; set; } = string.Empty;
    /// <summary>Whether the apply operation auto-fitted report columns.</summary>
    public bool AutoFitColumns { get; set; }
    /// <summary>Ordered normalized section readback.</summary>
    public List<ReportFormatSectionState> Sections { get; set; } = [];
    /// <summary>SHA-256 fingerprint of the ordered normalized section state.</summary>
    public string Fingerprint { get; set; } = string.Empty;
}

/// <summary>
/// Apply and inspect a deterministic report layout. Every section range is explicit; the
/// implementation never depends on ActiveCell, Selection, UsedRange, or inferred headers.
/// All ranges are validated before the first format mutation.
/// </summary>
[ServiceCategory("reportformat", "ReportFormat")]
[McpTool("report_format", Title = "Deterministic Report Formatting", Destructive = true, Category = "data",
    Description = "Apply or inspect deterministic report formatting using explicit title/header/body/total ranges. No selection or UsedRange inference. apply validates every section before changing Excel, preserves values/formulas/number formats, and returns exact style readback plus a stable fingerprint. Presets: professional or minimal. Colors use #RRGGBB.")]
public interface IReportFormatCommands
{
    /// <summary>
    /// Applies a complete deterministic report preset and returns normalized readback.
    /// Header and body ranges are required. Title and total ranges are optional. Every supplied
    /// section must span the same columns, be non-overlapping, and appear in report order.
    /// </summary>
    [ServiceAction("apply")]
    ReportFormatStateResult Apply(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        string? titleRange,
        [RequiredParameter] string headerRange,
        [RequiredParameter] string bodyRange,
        string? totalRange,
        [FromString] ReportFormatPreset preset = ReportFormatPreset.Professional,
        string accentColor = "#1F4E78",
        bool autoFitColumns = true);

    /// <summary>
    /// Reads the same normalized style fields and fingerprint without changing the workbook.
    /// </summary>
    [ServiceAction("get-state", IsMutation = false)]
    ReportFormatStateResult GetState(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        string? titleRange,
        [RequiredParameter] string headerRange,
        [RequiredParameter] string bodyRange,
        string? totalRange);
}
