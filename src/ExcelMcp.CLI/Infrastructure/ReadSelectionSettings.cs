using System.ComponentModel;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

internal class ReadSelectionSettings : CommandSettings
{
    [CommandOption("--fields <FIELDS>")]
    [Description("Comma-separated result paths, such as sessions.sessionId. Warnings and inspection limits remain visible.")]
    public string? Fields { get; init; }
}
