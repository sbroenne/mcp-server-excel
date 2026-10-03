using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>The cell category to discover within the requested range.</summary>
[JsonConverter(typeof(JsonStringEnumConverter<SpecialCellKind>))]
public enum SpecialCellKind
{
    /// <summary>Cells containing formulas, including formulas displaying empty text.</summary>
    [JsonStringEnumMemberName("formulas")]
    Formulas,

    /// <summary>Cells containing values but no formulas.</summary>
    [JsonStringEnumMemberName("constants")]
    Constants,

    /// <summary>Cells with no stored value or formula.</summary>
    [JsonStringEnumMemberName("blanks")]
    Blanks,

    /// <summary>Cells containing Excel errors, whether constants or formula results.</summary>
    [JsonStringEnumMemberName("errors")]
    Errors,

    /// <summary>Cells outside hidden rows and columns, including filtering and outlining.</summary>
    [JsonStringEnumMemberName("visible")]
    Visible
}
