using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

/// <summary>Workbook areas available to integrity validation.</summary>
[JsonConverter(typeof(JsonStringEnumConverter<WorkbookIntegrityCheck>))]
public enum WorkbookIntegrityCheck
{
    /// <summary>Find cells whose formulas currently evaluate to Excel errors.</summary>
    [JsonStringEnumMemberName("formula-errors")]
    FormulaErrors,

    /// <summary>Inspect external Excel workbook links and their current status.</summary>
    [JsonStringEnumMemberName("external-links")]
    ExternalLinks,

    /// <summary>Inspect table structure, headers, and calculated-column consistency.</summary>
    [JsonStringEnumMemberName("tables")]
    Tables,

    /// <summary>Compare caller-supplied expected numeric cells.</summary>
    [JsonStringEnumMemberName("control-totals")]
    ControlTotals
}
