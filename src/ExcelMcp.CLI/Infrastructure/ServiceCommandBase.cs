using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

/// <summary>
/// Base class for CLI commands that send requests to the service.
/// Handles common validation and execution patterns.
/// </summary>
internal abstract class ServiceCommandBase<TSettings> : AsyncCommand<TSettings>
    where TSettings : CommandSettings
{
    /// <summary>
    /// Gets the session ID from settings.
    /// </summary>
    protected abstract string? GetSessionId(TSettings settings);

    /// <summary>
    /// Gets the action from settings.
    /// </summary>
    protected abstract string? GetAction(TSettings settings);

    /// <summary>
    /// Gets the valid actions for this command.
    /// </summary>
    protected abstract IReadOnlyList<string> ValidActions { get; }

    /// <summary>
    /// Routes the action to a service command and args.
    /// </summary>
    protected abstract (string command, object? args) Route(TSettings settings, string action);

    /// <summary>
    /// Whether this command requires a session ID. Default is true.
    /// Override to return false for commands that don't need a session.
    /// </summary>
    protected virtual bool RequiresSession => true;

    /// <summary>
    /// Validates settings and executes the command.
    /// Returns early with error code if validation fails.
    /// </summary>
    public sealed override async Task<int> ExecuteAsync(CommandContext context, TSettings settings, CancellationToken cancellationToken)
    {
        // Session validation
        var sessionId = GetSessionId(settings);
        if (RequiresSession && string.IsNullOrWhiteSpace(sessionId))
        {
            AnsiConsole.MarkupLine("[red]Session ID is required. Use --session <id>[/]");
            return 1;
        }

        // Action validation
        var rawAction = GetAction(settings);
        if (string.IsNullOrWhiteSpace(rawAction))
        {
            AnsiConsole.MarkupLine("[red]Action is required.[/]");
            return 1;
        }

        var action = rawAction.Trim().ToLowerInvariant();
        if (!ValidActions.Contains(action, StringComparer.OrdinalIgnoreCase))
        {
            var validList = string.Join(", ", ValidActions);
            AnsiConsole.MarkupLine($"[red]Invalid action '{action}'. Valid actions: {validList}[/]");
            return 1;
        }

        // Route and execute
        string command;
        object? args;
        try
        {
            (command, args) = Route(settings, action);
        }
        catch (ArgumentException ex)
        {
            // Parameter validation failed (e.g., required param missing)
            // Return clean JSON error with exit code 1 instead of unhandled crash
            Console.WriteLine(JsonSerializer.Serialize(
                new { success = false, error = ex.Message },
                ServiceProtocol.JsonOptions));
            return 1;
        }

        // Connect to CLI daemon service (env var override for testing)
        var pipeName = Environment.GetEnvironmentVariable("EXCELMCP_CLI_PIPE") ?? ServiceSecurity.GetCliPipeName();
        using var client = new ServiceClient(pipeName);
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = command,
            SessionId = sessionId,
            Args = args != null ? JsonSerializer.Serialize(args, ServiceProtocol.JsonOptions) : null
        }, cancellationToken);

        // Output result
        if (response.Success)
        {
            Console.WriteLine(!string.IsNullOrEmpty(response.Result)
                ? response.Result
                : JsonSerializer.Serialize(new { success = true }, ServiceProtocol.JsonOptions));
            return 0;
        }
        else
        {
            Console.WriteLine(JsonSerializer.Serialize(
                new { success = false, error = response.ErrorMessage },
                ServiceProtocol.JsonOptions));
            return 1;
        }
    }
}
