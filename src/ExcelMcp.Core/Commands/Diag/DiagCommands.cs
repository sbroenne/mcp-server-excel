using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Diag;

/// <summary>
/// Implementation of diagnostic commands.
/// These commands don't touch Excel COM — they validate CLI/MCP infrastructure.
/// </summary>
public sealed class DiagCommands : IDiagCommands
{
    /// <inheritdoc />
    public DiagResult Ping()
    {
        return new DiagResult
        {
            Success = true,
            Action = "ping",
            Message = "pong",
            Timestamp = DateTime.UtcNow.ToString("o")
        };
    }

    /// <inheritdoc />
    public DiagResult Echo(string message, string? tag = null)
    {
        return new DiagResult
        {
            Success = true,
            Action = "echo",
            Message = message,
            Tag = tag,
            Timestamp = DateTime.UtcNow.ToString("o")
        };
    }

    /// <inheritdoc />
    public DiagResult ValidateParams(string name, int count, string? label = null, bool verbose = false)
    {
        return new DiagResult
        {
            Success = true,
            Action = "validate-params",
            Timestamp = DateTime.UtcNow.ToString("o"),
            Parameters = new Dictionary<string, object?>
            {
                ["name"] = name,
                ["count"] = count,
                ["label"] = label,
                ["verbose"] = verbose
            }
        };
    }
}
