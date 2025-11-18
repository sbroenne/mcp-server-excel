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
        var result = _fileCommands.CreateEmpty(settings.FilePath, settings.Overwrite);
        if (!result.Success)
        {
            _console.WriteError(result.ErrorMessage ?? "Failed to create Excel file.");
            return -1;
        }

        _console.WriteJson(new
        {
            success = true,
            filePath = settings.FilePath,
            settings.Overwrite
        });
        return 0;
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<file>")]
        public string FilePath { get; init; } = string.Empty;

        [CommandOption("-o|--overwrite")]
        public bool Overwrite { get; init; }
    }
}
