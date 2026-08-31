using System.ComponentModel;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Core.Utilities;
using Sbroenne.ExcelMcp.Service;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

internal sealed class FileCreateSavepointCommand
    : AsyncCommand<FileCreateSavepointCommand.Settings>
{
    protected override Task<int> ExecuteAsync(
        CommandContext context,
        Settings settings,
        CancellationToken cancellationToken) =>
        FileSavepointCommandRunner.ExecuteAsync(
            "create-savepoint",
            settings.SessionId,
            settings.Name,
            settings.TimeoutSeconds,
            cancellationToken);

    internal sealed class Settings : FileSavepointTimeoutSettings
    {
    }
}

internal sealed class FileRollbackSavepointCommand
    : AsyncCommand<FileRollbackSavepointCommand.Settings>
{
    protected override Task<int> ExecuteAsync(
        CommandContext context,
        Settings settings,
        CancellationToken cancellationToken) =>
        FileSavepointCommandRunner.ExecuteAsync(
            "rollback-savepoint",
            settings.SessionId,
            settings.Name,
            settings.TimeoutSeconds,
            cancellationToken);

    internal sealed class Settings : FileSavepointTimeoutSettings
    {
    }
}

internal sealed class FileReleaseSavepointCommand
    : AsyncCommand<FileReleaseSavepointCommand.Settings>
{
    protected override Task<int> ExecuteAsync(
        CommandContext context,
        Settings settings,
        CancellationToken cancellationToken) =>
        FileSavepointCommandRunner.ExecuteAsync(
            "release-savepoint",
            settings.SessionId,
            settings.Name,
            timeoutSeconds: null,
            cancellationToken);

    internal sealed class Settings : FileSavepointSettings
    {
    }
}

internal sealed class FileListSavepointsCommand
    : AsyncCommand<FileListSavepointsCommand.Settings>
{
    protected override Task<int> ExecuteAsync(
        CommandContext context,
        Settings settings,
        CancellationToken cancellationToken) =>
        FileSavepointCommandRunner.ExecuteAsync(
            "list-savepoints",
            settings.SessionId,
            name: null,
            timeoutSeconds: null,
            cancellationToken);

    internal sealed class Settings : CommandSettings
    {
        [CommandOption("-s|--session <SESSION>")]
        [Description("Session ID that owns the savepoints")]
        public string SessionId { get; init; } = string.Empty;
    }
}

internal class FileSavepointSettings : CommandSettings
{
    [CommandOption("-s|--session <SESSION>")]
    [Description("Session ID that owns the savepoint")]
    public string SessionId { get; init; } = string.Empty;

    [CommandOption("--name <NAME>")]
    [Description("Savepoint name (1-128 ASCII letters, digits, '.', '_', or '-')")]
    public string Name { get; init; } = string.Empty;
}

internal class FileSavepointTimeoutSettings : FileSavepointSettings
{
    [CommandOption("--timeout <SECONDS>")]
    [Description("Savepoint operation timeout in whole seconds (default: 120; range: 10-3600)")]
    public int TimeoutSeconds { get; init; } = 120;
}

internal static class FileSavepointCommandRunner
{
    internal static async Task<int> ExecuteAsync(
        string action,
        string sessionId,
        string? name,
        int? timeoutSeconds,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return CliErrorOutput.WriteError("Session ID is required.");
        }

        if (action is not "list-savepoints" && string.IsNullOrWhiteSpace(name))
        {
            return CliErrorOutput.WriteError("Savepoint name is required.");
        }

        try
        {
            ParameterTransforms.ValidateTimeoutSeconds(
                timeoutSeconds,
                "timeout",
                minimumSeconds: 10,
                maximumSeconds: 3600);
        }
        catch (ArgumentOutOfRangeException ex)
        {
            return CliErrorOutput.WriteError(ex.Message);
        }

        object? args = action switch
        {
            "create-savepoint" or "rollback-savepoint" => new
            {
                name,
                timeoutSeconds
            },
            "release-savepoint" => new { name },
            "list-savepoints" => null,
            _ => throw new ArgumentException(
                $"Unknown savepoint action: {action}",
                nameof(action))
        };

        using var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = $"session.{action}",
            SessionId = sessionId,
            Args = args == null
                ? null
                : JsonSerializer.Serialize(args, ServiceProtocol.JsonOptions)
        }, cancellationToken);

        if (!response.Success)
        {
            return CliErrorOutput.WriteServiceErrorWithResult(response);
        }

        if (string.IsNullOrWhiteSpace(response.Result))
        {
            return CliErrorOutput.WriteError(
                $"Service returned an invalid {action} response.");
        }

        Console.WriteLine(response.Result);
        return 0;
    }
}
