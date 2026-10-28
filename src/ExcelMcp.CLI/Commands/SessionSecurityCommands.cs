using System.ComponentModel;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Telemetry;
using Sbroenne.ExcelMcp.Service;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

internal sealed class SessionNoticesCommand : AsyncCommand<SessionNoticesCommand.Settings>
{
    public override Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken) =>
        SessionSecurityCommand.SendAsync("notices", settings.SessionId, null, cancellationToken);

    internal sealed class Settings : CommandSettings
    {
        [CommandOption("-s|--session <SESSION>")]
        [Description("Exact workbook session to inspect without enabling content")]
        public string SessionId { get; init; } = string.Empty;
    }
}

internal sealed class SessionSecurityCommand : AsyncCommand<SessionSecurityCommand.Settings>
{
    public override Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        var action = context.Name;
        if (action == "enable-macros" && settings.AllowRememberedTrust
            || action != "enable-macros" && settings.AllowReopen)
            return Task.FromResult(CliErrorOutput.WriteError("Use --allow-reopen only for enable-macros and --allow-remembered-trust only for enable-editing or enable-connections."));
        object args = action == "enable-macros"
            ? new { approved = settings.Approved, allowReopen = settings.AllowReopen }
            : new { approved = settings.Approved, allowRememberedTrust = settings.AllowRememberedTrust };
        return SendAsync(action, settings.SessionId, args, cancellationToken);
    }

    internal static async Task<int> SendAsync(string action, string sessionId, object? args, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
            return CliErrorOutput.WriteError("Session ID is required.");
        using var client = await CliCommandRuntime.Current.ClientFactory.ConnectAsync(cancellationToken);
        var request = new ServiceRequest
        {
            Command = "session." + action,
            SessionId = sessionId,
            Args = args is null ? null : JsonSerializer.Serialize(args, ServiceProtocol.JsonOptions)
        };
        var response = await CliTelemetry.TrackCommandAsync(request, () => client.SendAsync(request, cancellationToken));
        if (!response.Success)
            return CliErrorOutput.WriteServiceError(response);
        return CliResultOutput.Write(response, request.Command);
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandOption("-s|--session <SESSION>")]
        [Description("Exact workbook session to change")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--approved")]
        [Description("Explicit approval of this action on this workbook. Approval flags are not confirmation prompts and do not bypass company policy.")]
        public bool Approved { get; init; }

        [CommandOption("--allow-remembered-trust")]
        [Description("Approve Excel remembering trust for this file; required for editing and connection enabling")]
        public bool AllowRememberedTrust { get; init; }

        [CommandOption("--allow-reopen")]
        [Description("Separately approve closing and reopening a saved workbook; opening macros can run")]
        public bool AllowReopen { get; init; }
    }
}
