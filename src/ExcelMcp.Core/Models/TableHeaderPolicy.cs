using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Controls how range-to-table conversion handles blank and duplicate headers.
/// </summary>
[JsonConverter(typeof(JsonStringEnumConverter<TableHeaderPolicy>))]
public enum TableHeaderPolicy
{
    /// <summary>
    /// Report invalid headers as blockers without changing the range.
    /// </summary>
    Report,

    /// <summary>
    /// Generate deterministic names for blank and duplicate headers.
    /// </summary>
    Normalize
}
