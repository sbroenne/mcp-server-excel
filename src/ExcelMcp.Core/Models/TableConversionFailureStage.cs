using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Stage at which range-to-table conversion failed.
/// </summary>
[JsonConverter(typeof(JsonStringEnumConverter<TableConversionFailureStage>))]
public enum TableConversionFailureStage
{
    /// <summary>
    /// The source range failed preflight policy checks.
    /// </summary>
    Preflight,

    /// <summary>
    /// The rollback snapshot could not be created.
    /// </summary>
    Snapshot,

    /// <summary>
    /// Explicit range normalization failed.
    /// </summary>
    Normalization,

    /// <summary>
    /// Excel could not create or name the table.
    /// </summary>
    Creation,

    /// <summary>
    /// Excel could not apply the requested table style.
    /// </summary>
    Styling,

    /// <summary>
    /// The created table failed post-creation validation.
    /// </summary>
    Validation,

    /// <summary>
    /// Rollback or rollback verification failed.
    /// </summary>
    Rollback
}
