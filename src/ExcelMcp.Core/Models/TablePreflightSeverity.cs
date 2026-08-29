using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Severity of a table-integrity finding.
/// </summary>
[JsonConverter(typeof(JsonStringEnumConverter<TablePreflightSeverity>))]
public enum TablePreflightSeverity
{
    /// <summary>
    /// A deterministic problem that prevents a safe table operation.
    /// </summary>
    Blocker,

    /// <summary>
    /// A heuristic concern that should be reviewed but does not prevent creation.
    /// </summary>
    Warning
}
