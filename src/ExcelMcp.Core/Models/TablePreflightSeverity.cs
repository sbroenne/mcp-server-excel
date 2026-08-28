using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Severity of a table-creation preflight finding.
/// </summary>
[JsonConverter(typeof(JsonStringEnumConverter<TablePreflightSeverity>))]
public enum TablePreflightSeverity
{
    /// <summary>
    /// A deterministic problem that prevents safe table creation.
    /// </summary>
    Blocker,

    /// <summary>
    /// A heuristic concern that should be reviewed but does not prevent creation.
    /// </summary>
    Warning
}
