using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Controls how range-to-table conversion handles merged header cells.
/// </summary>
[JsonConverter(typeof(JsonStringEnumConverter<TableMergedHeaderPolicy>))]
public enum TableMergedHeaderPolicy
{
    /// <summary>
    /// Report merged headers as blockers without changing the range.
    /// </summary>
    Report,

    /// <summary>
    /// Unmerge header-only merged areas and repeat the top-left value.
    /// </summary>
    UnmergeAndFill
}
