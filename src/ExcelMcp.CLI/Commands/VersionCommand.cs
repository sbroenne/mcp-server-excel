using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Displays the current excelcli version (as reported by <see cref="VersionReporter"/>).
/// </summary>
internal sealed class VersionCommand : Command
{
    private readonly ICliConsole _console;

    public VersionCommand(ICliConsole console)
    {
        _console = console ?? throw new ArgumentNullException(nameof(console));
    }

    public override int Execute(CommandContext context, CancellationToken cancellationToken)
    {
        VersionReporter.WriteVersion();
        _console.WriteInfo("excelcli refactor is in progress. Session commands are available. Additional commands will return soon.");
        return 0;
    }
}
