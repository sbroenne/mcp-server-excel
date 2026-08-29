using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Reason a table conversion changed a header.
/// </summary>
[JsonConverter(typeof(JsonStringEnumConverter<TableHeaderChangeReason>))]
public enum TableHeaderChangeReason
{
    /// <summary>
    /// The header cell was blank.
    /// </summary>
    Blank,

    /// <summary>
    /// The header duplicated an earlier header.
    /// </summary>
    Duplicate
}
