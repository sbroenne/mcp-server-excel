using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Explicit date value types accepted inside a range set-values matrix.
/// </summary>
[JsonConverter(typeof(JsonStringEnumConverter<TypedCellValueType>))]
public enum TypedCellValueType
{
    /// <summary>An ISO calendar date in yyyy-MM-dd form.</summary>
    [JsonStringEnumMemberName("date")]
    Date,

    /// <summary>An ISO local date and time without a timezone.</summary>
    [JsonStringEnumMemberName("datetime")]
    DateTime,

    /// <summary>An ISO date and time with Z or a numeric UTC offset.</summary>
    [JsonStringEnumMemberName("datetime-offset")]
    DateTimeOffset
}
