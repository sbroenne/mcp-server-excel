using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Sbroenne.ExcelMcp.Service;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

// ============================================================================
// SESSION COMMANDS
// ============================================================================

internal sealed class SessionCreateCommand : AsyncCommand<SessionCreateCommand.Settings>
{
    protected override async Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.FilePath))
        {
            return CliErrorOutput.WriteError("File path is required.");
        }

        try
        {
            ParameterTransforms.ValidateTimeoutSeconds(
                settings.TimeoutSeconds,
                "timeout",
                minimumSeconds: 10,
                maximumSeconds: 3600);
        }
        catch (ArgumentOutOfRangeException ex)
        {
            return CliErrorOutput.WriteError(ex.Message);
        }

        using var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = "session.create",
            Args = JsonSerializer.Serialize(new
            {
                filePath = settings.FilePath,
                show = settings.Show,
                timeoutSeconds = settings.TimeoutSeconds
            }, ServiceProtocol.JsonOptions)
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(response.Result);
            return 0;
        }
        else
        {
            return CliErrorOutput.WriteServiceError(response);
        }
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<FILE>")]
        [Description("Path to the new Excel file to create")]
        public string FilePath { get; init; } = string.Empty;

        [CommandOption("--timeout <SECONDS>")]
        [Description("Session open/create and operation timeout in whole seconds (default: 120; range: 10-3600)")]
        public int? TimeoutSeconds { get; init; }

        [CommandOption("--show")]
        [Description("Show the Excel window for IRM/auth prompts instead of running hidden")]
        public bool Show { get; init; }
    }
}

internal sealed class SessionOpenCommand : AsyncCommand<SessionOpenCommand.Settings>
{
    protected override async Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.FilePath))
        {
            return CliErrorOutput.WriteError("File path is required.");
        }

        try
        {
            ParameterTransforms.ValidateTimeoutSeconds(
                settings.TimeoutSeconds,
                "timeout",
                minimumSeconds: 10,
                maximumSeconds: 3600);
        }
        catch (ArgumentOutOfRangeException ex)
        {
            return CliErrorOutput.WriteError(ex.Message);
        }

        using var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(new
            {
                filePath = settings.FilePath,
                show = settings.Show,
                timeoutSeconds = settings.TimeoutSeconds
            }, ServiceProtocol.JsonOptions)
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(response.Result);
            return 0;
        }
        else
        {
            return CliErrorOutput.WriteServiceError(response);
        }
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<FILE>")]
        [Description("Path to the Excel file to open")]
        public string FilePath { get; init; } = string.Empty;

        [CommandOption("--timeout <SECONDS>")]
        [Description("Session open and operation timeout in whole seconds (default: 120; range: 10-3600)")]
        public int? TimeoutSeconds { get; init; }

        [CommandOption("--show")]
        [Description("Show the Excel window for IRM/auth prompts instead of running hidden")]
        public bool Show { get; init; }
    }
}

internal sealed class SessionCloseCommand : AsyncCommand<SessionCloseCommand.Settings>
{
    protected override async Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.SessionId))
        {
            return CliErrorOutput.WriteError("Session ID is required.");
        }

        using var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = settings.SessionId,
            Args = JsonSerializer.Serialize(new { save = settings.Save }, ServiceProtocol.JsonOptions)
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = true, message = settings.Save ? "Session closed and saved." : "Session closed." }, ServiceProtocol.JsonOptions));
            return 0;
        }
        else
        {
            return CliErrorOutput.WriteServiceError(response);
        }
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandOption("-s|--session <SESSION>")]
        [Description("Session ID to close")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--save")]
        [Description("Save changes before closing")]
        public bool Save { get; init; }
    }
}

internal sealed class SessionListCommand : AsyncCommand
{
    protected override async Task<int> ExecuteAsync(CommandContext context, CancellationToken cancellationToken)
    {
        var pipeName = DaemonAutoStart.GetPipeName();
        var observation = DaemonConnectionPolicy.Observe(pipeName);
        var response = await DaemonConnectionPolicy.SendControlRequestAsync(
            pipeName,
            new ServiceRequest { Command = "session.list" },
            cancellationToken,
            observation.IsStopped
                ? DaemonConnectionPolicy.InitialProbeTimeout
                : DaemonConnectionPolicy.ControlTimeout);
        if (response.Success && response.Result != null)
        {
            var result = JsonNode.Parse(response.Result) as JsonObject
                ?? throw new JsonException("Service returned an invalid session list response.");
            result["daemonState"] = DaemonConnectionPolicy.RunningState;
            Console.WriteLine(result.ToJsonString(ServiceProtocol.JsonOptions));
            return 0;
        }

        if (response.Success)
        {
            response = new ServiceResponse
            {
                Success = false,
                Command = "session.list",
                ErrorCategory = "InvalidResponse",
                ErrorMessage = "Service returned an invalid session list response."
            };
        }

        var failureState = DaemonConnectionPolicy.ResolveFailureState(pipeName, response);
        if (failureState.Name == DaemonConnectionPolicy.StoppedState)
        {
            return WriteStoppedSessionList();
        }

        return CliErrorOutput.WriteDaemonError(response, failureState.Name, failureState.Running);
    }

    private static int WriteStoppedSessionList()
    {
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            success = true,
            daemonState = DaemonConnectionPolicy.StoppedState,
            sessions = Array.Empty<object>(),
            count = 0
        }, ServiceProtocol.JsonOptions));
        return 0;
    }
}

internal sealed class SessionTestCommand : AsyncCommand<SessionTestCommand.Settings>
{
    protected override async Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.FilePath))
        {
            return CliErrorOutput.WriteError("File path is required.");
        }

        using var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = "session.test",
            Args = JsonSerializer.Serialize(
                new { filePath = settings.FilePath },
                ServiceProtocol.JsonOptions)
        }, cancellationToken);

        if (!response.Success)
        {
            return CliErrorOutput.WriteServiceError(response);
        }

        if (string.IsNullOrWhiteSpace(response.Result))
        {
            return CliErrorOutput.WriteError("Service returned an invalid file test response.");
        }

        var result = ServiceProtocol.Deserialize<FileValidationInfo>(response.Result);
        if (result == null)
        {
            return CliErrorOutput.WriteError("Service returned an invalid file test response.");
        }

        Console.WriteLine(response.Result);
        return result.CanOpen ? 0 : 1;
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<FILE>")]
        [Description("Full path to test for existence, validity, openability, and IRM/AIP requirements")]
        public string FilePath { get; init; } = string.Empty;
    }
}
