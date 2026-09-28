using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Type of problem found while checking a proposed table range.
/// </summary>
[JsonConverter(typeof(JsonStringEnumConverter<TablePreflightFindingKind>))]
public enum TablePreflightFindingKind
{
    /// <summary>
    /// The proposed range intersects one or more merged ranges.
    /// </summary>
    MergedCells,

    /// <summary>
    /// One or more header cells are blank.
    /// </summary>
    BlankHeaders,

    /// <summary>
    /// Two or more header cells resolve to the same name.
    /// </summary>
    DuplicateHeaders,

    /// <summary>
    /// Populated columns in the same current region are excluded.
    /// </summary>
    ExcludedContiguousColumns,

    /// <summary>
    /// A formula contains row references that may not remain aligned after sorting.
    /// </summary>
    SortSensitiveFormula,

    /// <summary>
    /// Formula risk analysis was skipped because the proposed range exceeds the bounded scan size.
    /// </summary>
    FormulaScanSkipped,

    /// <summary>
    /// The requested table name already exists.
    /// </summary>
    TableNameExists
}
