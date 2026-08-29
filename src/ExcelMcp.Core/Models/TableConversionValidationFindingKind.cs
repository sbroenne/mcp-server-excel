using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Type of post-creation table validation failure.
/// </summary>
[JsonConverter(typeof(JsonStringEnumConverter<TableConversionValidationFindingKind>))]
public enum TableConversionValidationFindingKind
{
    /// <summary>
    /// A formula evaluated to an Excel error value.
    /// </summary>
    FormulaError,

    /// <summary>
    /// A formula-bearing data column is not a consistent calculated column.
    /// </summary>
    InconsistentCalculatedColumn,

    /// <summary>
    /// The created table does not match the requested identity or range.
    /// </summary>
    TableMismatch,

    /// <summary>
    /// The requested table style was not applied.
    /// </summary>
    StyleMismatch,

    /// <summary>
    /// Excel unexpectedly enabled a totals row.
    /// </summary>
    UnexpectedTotalsRow,

    /// <summary>
    /// Non-header source content changed during conversion.
    /// </summary>
    SourceContentChanged
}
