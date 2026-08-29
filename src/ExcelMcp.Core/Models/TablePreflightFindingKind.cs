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
    /// The requested table name already exists.
    /// </summary>
    TableNameExists,

    /// <summary>
    /// A formula outside the table appears aligned with table data rows.
    /// </summary>
    RowAssociatedFormulaOutsideTable,

    /// <summary>
    /// A table column mixes formulas, formula patterns, or literal values.
    /// </summary>
    MixedFormulaColumn,

    /// <summary>
    /// A requested row-key column does not exist or contains an invalid key.
    /// </summary>
    InvalidRowKey,

    /// <summary>
    /// A requested composite row key is not unique.
    /// </summary>
    DuplicateRowKey,

    /// <summary>
    /// A requested control total cannot be evaluated.
    /// </summary>
    InvalidControlTotal,

    /// <summary>
    /// The table range, shape, headers, or totals row changed during sorting.
    /// </summary>
    TableStructureChanged,

    /// <summary>
    /// Complete logical table rows were not preserved.
    /// </summary>
    TableRowsChanged,

    /// <summary>
    /// A previously consistent calculated-column formula pattern changed.
    /// </summary>
    CalculatedColumnChanged,

    /// <summary>
    /// Caller-supplied row identities or their complete row contents changed.
    /// </summary>
    RowKeyMismatch,

    /// <summary>
    /// A caller-supplied numeric control total changed beyond its tolerance.
    /// </summary>
    ControlTotalMismatch
}
