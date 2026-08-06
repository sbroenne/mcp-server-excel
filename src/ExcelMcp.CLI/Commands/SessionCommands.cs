using System.ComponentModel;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Core.Models;
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
        [Description("Session open/create and operation timeout in seconds (default: 120)")]
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
        [Description("Session open and operation timeout in seconds (default: 120)")]
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
    private static readonly TimeSpan CommandTimeout = TimeSpan.FromSeconds(2);

    protected override async Task<int> ExecuteAsync(CommandContext context, CancellationToken cancellationToken)
    {
        var pipeName = DaemonAutoStart.GetPipeName();
        using var client = new ServiceClient(pipeName, connectTimeout: CommandTimeout, requestTimeout: CommandTimeout);

        try
        {
            var response = await client.SendAsync(new ServiceRequest { Command = "session.list" }, cancellationToken);
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
        catch (Exception)
        {
            // Daemon not running — no sessions
            Console.WriteLine(JsonSerializer.Serialize(new { sessions = Array.Empty<object>() }, ServiceProtocol.JsonOptions));
            return 0;
        }
    }
}

internal sealed class SessionPreflightCommand : AsyncCommand<SessionPreflightCommand.Settings>
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
            Command = "session.preflight",
            SessionId = settings.SessionId
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(response.Result);
            return 0;
        }

        return CliErrorOutput.WriteServiceError(response);
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandOption("--session-id <SESSION>")]
        [Description("Session ID to inspect")]
        public string SessionId { get; init; } = string.Empty;
    }
}

internal sealed class SessionConfigureSafetyCommand : AsyncCommand<SessionConfigureSafetyCommand.Settings>
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
            Command = "session.configure-safety",
            SessionId = settings.SessionId,
            Args = JsonSerializer.Serialize(new SafetyConfigurationOptions
            {
                ReviewMode = settings.ReviewMode,
                CheckpointMode = settings.CheckpointMode,
                JournalMode = settings.JournalMode,
                VerificationMode = settings.VerificationMode,
                AbnormalShutdownPolicy = settings.AbnormalShutdownPolicy
            }, ServiceProtocol.JsonOptions)
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(response.Result);
            return 0;
        }

        return CliErrorOutput.WriteServiceError(response);
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandOption("--session-id <SESSION>")]
        [Description("Session ID to configure")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--review-mode <MODE>")]
        [Description("Safety review mode: off, optional, required")]
        public SafetyReviewMode? ReviewMode { get; init; }

        [CommandOption("--checkpoint-mode <MODE>")]
        [Description("Safety checkpoint mode: off, onRequest, required")]
        public SafetyCheckpointMode? CheckpointMode { get; init; }

        [CommandOption("--journal-mode <MODE>")]
        [Description("Safety journal mode: off, on")]
        public SafetyJournalMode? JournalMode { get; init; }

        [CommandOption("--verification-mode <MODE>")]
        [Description("Safety verification mode: off, on")]
        public SafetyVerificationMode? VerificationMode { get; init; }

        [CommandOption("--abnormal-shutdown-policy <POLICY>")]
        [Description("Safety shutdown policy: legacyAutoSave, discardWithRecoveryEvidence")]
        public SafetyAbnormalShutdownPolicy? AbnormalShutdownPolicy { get; init; }
    }
}

internal sealed class SessionJournalCommand : AsyncCommand<SessionJournalCommand.Settings>
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
            Command = "session.journal",
            SessionId = settings.SessionId
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(response.Result);
            return 0;
        }

        return CliErrorOutput.WriteServiceError(response);
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandOption("--session-id <SESSION>")]
        [Description("Session ID whose durable operation journal to render")]
        public string SessionId { get; init; } = string.Empty;
    }
}

internal sealed class SessionRecoveriesCommand : AsyncCommand
{
    protected override async Task<int> ExecuteAsync(CommandContext context, CancellationToken cancellationToken)
    {
        using var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);
        var response = await client.SendAsync(new ServiceRequest { Command = "recovery.list" }, cancellationToken);
        if (response.Success)
        {
            Console.WriteLine(response.Result);
            return 0;
        }

        return CliErrorOutput.WriteServiceError(response);
    }
}

internal sealed class SessionRecoverCommand : AsyncCommand<SessionRecoverCommand.Settings>
{
    protected override async Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.RecoveryId))
        {
            return CliErrorOutput.WriteError("Recovery ID is required.");
        }

        using var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = "recovery.recover",
            Args = JsonSerializer.Serialize(new
            {
                recoveryId = settings.RecoveryId,
                show = settings.Show,
                timeoutSeconds = settings.TimeoutSeconds
            }, ServiceProtocol.JsonOptions)
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(response.Result);
            return 0;
        }

        return CliErrorOutput.WriteServiceError(response);
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandOption("--recovery-id <RECOVERY>")]
        [Description("Recovery ID returned by session recoveries")]
        public string RecoveryId { get; init; } = string.Empty;

        [CommandOption("--show")]
        [Description("Show the recovered Excel session")]
        public bool Show { get; init; }

        [CommandOption("--timeout <SECONDS>")]
        [Description("Recovery session timeout in seconds (default: 120)")]
        public int? TimeoutSeconds { get; init; }
    }
}

internal sealed class SessionSaveCommand : AsyncCommand<SessionSaveCommand.Settings>
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
            Command = "session.save",
            SessionId = settings.SessionId
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = true, message = "Session saved." }, ServiceProtocol.JsonOptions));
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
        [Description("Session ID to save")]
        public string SessionId { get; init; } = string.Empty;
    }
}


