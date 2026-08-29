using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

/// <summary>Workbook concern represented by an integrity finding.</summary>
[JsonConverter(typeof(JsonStringEnumConverter<WorkbookIntegrityCategory>))]
public enum WorkbookIntegrityCategory
{
    /// <summary>A formula currently evaluates to an Excel error.</summary>
    [JsonStringEnumMemberName("formula-error")]
    FormulaError,

    /// <summary>A formula contains a broken cell, range, sheet, or workbook reference.</summary>
    [JsonStringEnumMemberName("broken-reference")]
    BrokenReference,

    /// <summary>An external Excel workbook link and its current status.</summary>
    [JsonStringEnumMemberName("external-link")]
    ExternalLink,

    /// <summary>A likely calculated-column formula inconsistency.</summary>
    [JsonStringEnumMemberName("calculated-column")]
    CalculatedColumn,

    /// <summary>An inconsistency in an Excel table's structural ranges or collections.</summary>
    [JsonStringEnumMemberName("table-structure")]
    TableStructure,

    /// <summary>A missing, hidden, empty, duplicated, or error-valued table header.</summary>
    [JsonStringEnumMemberName("table-header")]
    TableHeader,

    /// <summary>A caller-supplied expected numeric cell does not match.</summary>
    [JsonStringEnumMemberName("control-total")]
    ControlTotal,

    /// <summary>Workbook calculation settings may make cached values stale.</summary>
    [JsonStringEnumMemberName("calculation-state")]
    CalculationState
}
