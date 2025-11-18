using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Infrastructure.Session;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands.Session;

internal sealed class SessionCloseCommand : Command<SessionCloseCommand.Settings>
{
    private readonly ISessionService _sessionService;
    private readonly ICliConsole _console;

    public SessionCloseCommand(ISessionService sessionService, ICliConsole console)
    {
        _sessionService = sessionService ?? throw new ArgumentNullException(nameof(sessionService));
        _console = console ?? throw new ArgumentNullException(nameof(console));
    }

    public override int Execute(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        var closed = _sessionService.Close(settings.SessionId);
        if (!closed)
        {
            _console.WriteError($"Session '{settings.SessionId}' not found.");
            return -1;
        }

        _console.WriteJson(new
        {
            success = true,
            sessionId = settings.SessionId,
            action = "close"
        });
        return 0;
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<session>")]
        public string SessionId { get; init; } = string.Empty;
    }
}
