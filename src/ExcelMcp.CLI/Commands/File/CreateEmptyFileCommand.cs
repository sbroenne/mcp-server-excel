using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Core.Commands;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands.File;

internal sealed class CreateEmptyFileCommand : Command<CreateEmptyFileCommand.Settings>
{
    private readonly IFileCommands _fileCommands;
    private readonly ICliConsole _console;

    public CreateEmptyFileCommand(IFileCommands fileCommands, ICliConsole console)
    {
        _fileCommands = fileCommands ?? throw new ArgumentNullException(nameof(fileCommands));
        _console = console ?? throw new ArgumentNullException(nameof(console));
    }

    public override int Execute(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        try
        {
            _fileCommands.CreateEmpty(settings.FilePath, settings.Overwrite);

            _console.WriteJson(new
            {
                success = true,
                filePath = settings.FilePath,
                settings.Overwrite
            });

            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to create Excel file: {ex.Message}");
            return 1;
        }
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<file>")]
        public string FilePath { get; init; } = string.Empty;

        [CommandOption("-o|--overwrite")]
        public bool Overwrite { get; init; }
    }
}
