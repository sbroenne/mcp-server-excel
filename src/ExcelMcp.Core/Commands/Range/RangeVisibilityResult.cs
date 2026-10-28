using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>Whole worksheet dimensions affected by visibility operations.</summary>
[JsonConverter(typeof(JsonStringEnumConverter<VisibilityAxis>))]
public enum VisibilityAxis
{
    /// <summary>Entire rows intersecting the scope.</summary>
    [JsonStringEnumMemberName("rows")]
    Rows,
    /// <summary>Entire columns intersecting the scope.</summary>
    [JsonStringEnumMemberName("columns")]
    Columns
}

/// <summary>Complete visibility state for unique rows or columns intersecting the requested scope.</summary>
public sealed class RangeVisibilityResult : OperationResult
{
    /// <summary>Resolved worksheet.</summary>
    public string SheetName { get; set; } = string.Empty;
    /// <summary>Exact requested scope.</summary>
    public string RangeAddress { get; set; } = string.Empty;
    /// <summary>Whole row or column axis.</summary>
    public VisibilityAxis Axis { get; set; }
    /// <summary>Whether the worksheet is currently in filter mode.</summary>
    public bool SheetFilterMode { get; set; }
    /// <summary>Row/column state is native; the cause of hiding cannot be established reliably.</summary>
    public string CauseCoverage { get; } = "Native Hidden has no cause flag; filter/outline context is not proof of the cause.";
    /// <summary>Every distinct intersecting row/column, sorted by index.</summary>
    public List<DimensionVisibility> Items { get; set; } = [];
}

/// <summary>Native hidden state and context of one whole dimension.</summary>
public sealed record DimensionVisibility(
    int Index, bool Hidden, string HiddenCause, double? Size, string SizeUnit,
    int OutlineLevel, bool WithinWorksheetAutoFilterDataRows);
