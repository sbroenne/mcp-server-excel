using Spectre.Console.Cli;

using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Infrastructure.Session;

namespace Sbroenne.ExcelMcp.CLI.Commands.Session;

internal sealed class SessionSaveCommand : Command<SessionSaveCommand.Settings>
{
    private readonly ISessionService _sessionService;
    private readonly ICliConsole _console;

    public SessionSaveCommand(ISessionService sessionService, ICliConsole console)
    {
        _sessionService = sessionService ?? throw new ArgumentNullException(nameof(sessionService));
        _console = console ?? throw new ArgumentNullException(nameof(console));
    }

    public override int Execute(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        var saved = _sessionService.Save(settings.SessionId);
        if (!saved)
        {
            _console.WriteError($"Session '{settings.SessionId}' not found.");
            return -1;
        }

        _console.WriteJson(new
        {
            success = true,
            sessionId = settings.SessionId,
            action = "save"
        });
        return 0;
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<session>")]
        public string SessionId { get; init; } = string.Empty;
    }
}
