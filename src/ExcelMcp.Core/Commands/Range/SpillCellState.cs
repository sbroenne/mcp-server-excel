using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>Native dynamic-array state of an inspected cell.</summary>
[JsonConverter(typeof(JsonStringEnumConverter<SpillCellState>))]
public enum SpillCellState
{
    /// <summary>Not a member of an expanded dynamic-array result.</summary>
    [JsonStringEnumMemberName("ordinary")]
    Ordinary,
    /// <summary>The formula cell generating a dynamic-array result.</summary>
    [JsonStringEnumMemberName("source")]
    Source,
    /// <summary>A result cell belonging to another source cell.</summary>
    [JsonStringEnumMemberName("result")]
    Result,
    /// <summary>A formula returning #SPILL! with no established extent.</summary>
    [JsonStringEnumMemberName("blocked")]
    Blocked
}
