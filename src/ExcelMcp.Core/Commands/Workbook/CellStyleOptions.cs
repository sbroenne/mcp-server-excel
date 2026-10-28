using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

/// <summary>Changes to a custom cell style. Omitted settings retain native values.</summary>
[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed class CellStyleOptions
{
    /// <summary>Shared font, fill, border, alignment and number-format options.</summary>
    public CellFormatOptions? FormatOptions { get; set; }
    /// <summary>Whether applying the style includes its font.</summary>
    public bool? IncludeFont { get; set; }
    /// <summary>Whether applying the style includes number format.</summary>
    public bool? IncludeNumber { get; set; }
    /// <summary>Whether applying the style includes alignment.</summary>
    public bool? IncludeAlignment { get; set; }
    /// <summary>Whether applying the style includes borders.</summary>
    public bool? IncludeBorder { get; set; }
    /// <summary>Whether applying the style includes fill.</summary>
    public bool? IncludePatterns { get; set; }
    /// <summary>Whether applying the style includes protection settings.</summary>
    public bool? IncludeProtection { get; set; }
    /// <summary>Native cell locking, effective only on protected worksheets.</summary>
    public bool? Locked { get; set; }
    /// <summary>Native formula hiding, effective only on protected worksheets.</summary>
    public bool? FormulaHidden { get; set; }
}

/// <summary>Native cell-style identity.</summary>
public sealed record CellStyleInfo(string Name, string NameLocal, bool BuiltIn);

/// <summary>Complete native cell-style definition. Column width/row height are not style properties; the native Style.MergeCells getter is unavailable.</summary>
public sealed record CellStyleDefinition(
    string Name, string NameLocal, bool BuiltIn,
    bool IncludeFont, bool IncludeNumber, bool IncludeAlignment,
    bool IncludeBorder, bool IncludePatterns, bool IncludeProtection,
    CellFormatSnapshot Format);

/// <summary>Every cell style in the selected workbook, without a cap.</summary>
public sealed class CellStyleListResult : ResultBase
{
    /// <summary>Complete native style catalogue.</summary>
    public List<CellStyleInfo> Styles { get; init; } = [];
}

/// <summary>Selected native cell-style definition.</summary>
public sealed class CellStyleResult : ResultBase
{
    /// <summary>Native definition and inclusion flags.</summary>
    public required CellStyleDefinition Style { get; init; }
    /// <summary>Native inspection limits; no unsupported state is invented.</summary>
    public List<string> ReadLimitations { get; init; } =
        ["Cell styles have no column-width/row-height settings or inside borders. Excel does not expose Style.MergeCells for inspection."];
}
