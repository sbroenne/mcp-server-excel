using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Diag;

/// <summary>
/// Diagnostic commands for testing CLI/MCP infrastructure without Excel.
/// These commands validate parameter parsing, routing, JSON serialization,
/// and error handling — no Excel COM session needed.
/// </summary>
[ServiceCategory("diag", "Diag")]
[NoSession]
public interface IDiagCommands
{
    /// <summary>
    /// Returns a simple success response. Used to verify the service is running
    /// and the CLI/MCP pipeline works end-to-end.
    /// </summary>
    /// <returns>Success result with timestamp</returns>
    DiagResult Ping();

    /// <summary>
    /// Echoes back the provided message. Used to verify parameter parsing
    /// and JSON serialization of required string parameters.
    /// </summary>
    /// <param name="message">The message to echo back (required)</param>
    /// <param name="tag">Optional tag to include in the response</param>
    /// <returns>Result containing the echoed message and tag</returns>
    DiagResult Echo(string message, string? tag = null);

    /// <summary>
    /// Validates various parameter types. Used to verify that the CLI/MCP
    /// infrastructure correctly parses and validates different parameter
    /// combinations (required strings, optional strings, booleans, integers).
    /// </summary>
    /// <param name="name">Required name parameter</param>
    /// <param name="count">Required integer parameter</param>
    /// <param name="label">Optional label parameter</param>
    /// <param name="verbose">Optional boolean flag (default: false)</param>
    /// <returns>Result containing all parsed parameter values</returns>
    [ServiceAction("validate-params")]
    DiagResult ValidateParams(string name, int count, string? label = null, bool verbose = false);
}
