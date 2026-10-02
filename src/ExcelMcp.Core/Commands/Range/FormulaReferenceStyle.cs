using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>Native formula reference notation.</summary>
[JsonConverter(typeof(JsonStringEnumConverter<FormulaReferenceStyle>))]
public enum FormulaReferenceStyle
{
    /// <summary>A1 cell references.</summary>
    [JsonStringEnumMemberName("a1")]
    A1,
    /// <summary>R1C1 absolute or relative row/column references.</summary>
    [JsonStringEnumMemberName("r1c1")]
    R1C1
}
