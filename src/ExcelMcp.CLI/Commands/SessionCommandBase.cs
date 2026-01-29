using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Infrastructure.Session;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Base class for CLI commands that operate on Excel sessions.
/// Provides common validation, error handling, and result writing.
/// </summary>
internal abstract class SessionCommandBase<TSettings> : Command<TSettings>
    where TSettings : SessionCommandBase<TSettings>.SessionSettings
{
    protected readonly ISessionService SessionService;
    protected readonly ICliConsole Console;

    protected SessionCommandBase(ISessionService sessionService, ICliConsole console)
    {
        SessionService = sessionService ?? throw new ArgumentNullException(nameof(sessionService));
        Console = console ?? throw new ArgumentNullException(nameof(console));
    }

    /// <summary>
    /// Command name for error messages (e.g., "chart", "table", "slicer").
    /// </summary>
    protected abstract string CommandName { get; }

    public sealed override int Execute(CommandContext context, TSettings settings, CancellationToken cancellationToken)
    {
        // Validate session
        if (string.IsNullOrWhiteSpace(settings.SessionId))
        {
            Console.WriteError("Session ID is required. Use 'session open' first.");
            return ExitCodes.MissingSession;
        }

        // Validate action
        var action = settings.Action?.Trim().ToLowerInvariant();
        if (string.IsNullOrEmpty(action))
        {
            Console.WriteError("Action is required.");
            return ExitCodes.MissingAction;
        }

        // Get batch and execute
        var batch = SessionService.GetBatch(settings.SessionId);
        return ExecuteAction(context, settings, batch, action, cancellationToken);
    }

    /// <summary>
    /// Execute the specific action. Override in derived classes.
    /// </summary>
    protected abstract int ExecuteAction(
        CommandContext context,
        TSettings settings,
        IExcelBatch batch,
        string action,
        CancellationToken cancellationToken);

    #region Helper Methods

    /// <summary>
    /// Writes a result object as JSON and returns appropriate exit code.
    /// </summary>
    protected int WriteResult(ResultBase result)
    {
        Console.WriteJson(result);
        return result.Success ? ExitCodes.Success : ExitCodes.OperationFailed;
    }

    /// <summary>
    /// Reports an unknown action and returns error code.
    /// </summary>
    protected int ReportUnknown(string action)
    {
        Console.WriteError($"Unknown {CommandName} action '{action}'.");
        return ExitCodes.UnknownAction;
    }

    /// <summary>
    /// Validates that a required parameter is provided.
    /// </summary>
    protected bool RequireParameter(string? value, string parameterName, string actionName)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            Console.WriteError($"--{parameterName} is required for {actionName}.");
            return false;
        }
        return true;
    }

    /// <summary>
    /// Validates that multiple required parameters are provided.
    /// </summary>
    protected bool RequireParameters(string actionName, params (string? value, string name)[] parameters)
    {
        var missing = parameters.Where(p => string.IsNullOrWhiteSpace(p.value)).Select(p => $"--{p.name}").ToList();
        if (missing.Count > 0)
        {
            Console.WriteError($"{string.Join(", ", missing)} {(missing.Count == 1 ? "is" : "are")} required for {actionName}.");
            return false;
        }
        return true;
    }

    /// <summary>
    /// Validates that a nullable value parameter is provided.
    /// </summary>
    protected bool RequireValueParameter<T>(T? value, string parameterName, string actionName) where T : struct
    {
        if (!value.HasValue)
        {
            Console.WriteError($"--{parameterName} is required for {actionName}.");
            return false;
        }
        return true;
    }

    /// <summary>
    /// Executes an action with try-catch and standard error handling.
    /// </summary>
    protected int ExecuteWithErrorHandling(string actionDescription, Func<int> action)
    {
        try
        {
            return action();
        }
        catch (Exception ex)
        {
            Console.WriteError($"Failed to {actionDescription}: {ex.Message}");
            return ExitCodes.OperationFailed;
        }
    }

    /// <summary>
    /// Executes a void action and returns success JSON.
    /// </summary>
    protected int ExecuteVoidAction(string actionDescription, Action action, object? successResult = null)
    {
        try
        {
            action();
            Console.WriteJson(successResult ?? new { success = true, message = $"{actionDescription} completed successfully." });
            return ExitCodes.Success;
        }
        catch (Exception ex)
        {
            Console.WriteError($"Failed to {actionDescription}: {ex.Message}");
            return ExitCodes.OperationFailed;
        }
    }

    #endregion

    /// <summary>
    /// Base settings class with common session and action parameters.
    /// </summary>
    internal abstract class SessionSettings : CommandSettings
    {
        [CommandArgument(0, "<action>")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        public string SessionId { get; init; } = string.Empty;
    }
}

/// <summary>
/// Standardized exit codes for CLI commands.
/// </summary>
internal static class ExitCodes
{
    /// <summary>Operation completed successfully.</summary>
    public const int Success = 0;

    /// <summary>Operation failed (business logic error).</summary>
    public const int OperationFailed = 1;

    /// <summary>Missing required session ID.</summary>
    public const int MissingSession = 2;

    /// <summary>Missing required action parameter.</summary>
    public const int MissingAction = 3;

    /// <summary>Unknown action specified.</summary>
    public const int UnknownAction = 4;

    /// <summary>Missing required parameter.</summary>
    public const int MissingParameter = 5;

    /// <summary>Invalid parameter value.</summary>
    public const int InvalidParameter = 6;

    /// <summary>File not found.</summary>
    public const int FileNotFound = 7;

    /// <summary>Permission denied.</summary>
    public const int PermissionDenied = 8;
}
