using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Infrastructure.Session;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands.Session;

internal sealed class SessionOpenCommand : Command<SessionOpenCommand.Settings>
{
    private readonly ISessionService _sessionService;
    private readonly ICliConsole _console;

    public SessionOpenCommand(ISessionService sessionService, ICliConsole console)
    {
        _sessionService = sessionService ?? throw new ArgumentNullException(nameof(sessionService));
        _console = console ?? throw new ArgumentNullException(nameof(console));
    }

    public override int Execute(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        var sessionId = _sessionService.Create(settings.FilePath);
        _console.WriteJson(new
        {
            success = true,
            sessionId,
            filePath = settings.FilePath
        });
        return 0;
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<file>")]
        public string FilePath { get; init; } = string.Empty;
    }
}
