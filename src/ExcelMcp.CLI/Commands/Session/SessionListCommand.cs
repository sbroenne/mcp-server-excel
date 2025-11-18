using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Infrastructure.Session;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands.Session;

internal sealed class SessionListCommand : Command
{
    private readonly ISessionService _sessionService;
    private readonly ICliConsole _console;

    public SessionListCommand(ISessionService sessionService, ICliConsole console)
    {
        _sessionService = sessionService ?? throw new ArgumentNullException(nameof(sessionService));
        _console = console ?? throw new ArgumentNullException(nameof(console));
    }

    public override int Execute(CommandContext context, CancellationToken cancellationToken)
    {
        var sessions = _sessionService.List();
        _console.WriteJson(new
        {
            success = true,
            action = "list",
            sessions = sessions.Select(s => new
            {
                s.SessionId,
                s.FilePath
            })
        });
        return 0;
    }
}
