using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Text.Json;
using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.McpServer.Telemetry;

#pragma warning disable IL2070 // 'this' argument does not satisfy 'DynamicallyAccessedMembersAttribute' requirements

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Base class for Excel MCP tools providing common patterns and utilities.
/// All Excel tools inherit from this to ensure consistency for LLM usage.
///
/// The MCP Server forwards ALL requests to the in-process ExcelMCP Service.
/// </summary>
[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicMethods)]
public static class ExcelToolsBase
{
    /// <summary>
    /// Ensures the ExcelMCP Service is running.
    /// The service is required for all MCP Server operations.
    /// </summary>
    public static async Task<bool> EnsureServiceAsync(CancellationToken cancellationToken = default)
    {
        return await ServiceBridge.ServiceBridge.EnsureServiceAsync(cancellationToken);
    }

    /// <summary>
    /// JSON serializer options optimized for LLM token efficiency.
    /// Uses compact formatting to reduce token consumption.
    /// </summary>
    /// <remarks>
    /// Token optimization settings:
    /// - WriteIndented = false: Removes whitespace (saves ~20% tokens)
    /// - DefaultIgnoreCondition = WhenWritingNull: Omits null properties
    /// - PropertyNamingPolicy = CamelCase: Consistent naming (e.g., success, errorMessage, filePath)
    /// - JsonStringEnumConverter: Human-readable enum values
    /// </remarks>
    public static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = false,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        Converters = { new JsonStringEnumConverter() }
    };

    /// <summary>
    /// Delegate wrapper for ForwardToService matching the generated code signature.
    /// Used by generated RouteAction methods.
    /// </summary>
    public static readonly Func<string, string, object?, string> ForwardToServiceFunc =
        (command, sessionId, args) => ForwardToService(command, sessionId, args);

    /// <summary>
    /// Forwards a command to the ExcelMCP Service and returns the JSON response.
    /// This is the primary method for MCP tools to execute commands.
    ///
    /// The command format is "category.action", e.g., "sheet.list", "range.get-values".
    /// The service handles session management and Core command execution.
    /// </summary>
    /// <param name="command">Service command in format "category.action"</param>
    /// <param name="sessionId">Session ID for the operation</param>
    /// <param name="args">Optional arguments object to serialize</param>
    /// <param name="timeoutSeconds">Optional timeout override</param>
    /// <returns>JSON response from service</returns>
    public static string ForwardToService(
        string command,
        string sessionId,
        object? args = null,
        int? timeoutSeconds = null)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = "sessionId is required. Use file 'open' action to start a session.",
                isError = true
            }, JsonOptions);
        }

        var response = ServiceBridge.ServiceBridge.SendAsync(command, sessionId, args, timeoutSeconds).GetAwaiter().GetResult();

        if (!response.Success)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = response.ErrorMessage ?? $"Command '{command}' failed",
                isError = true
            }, JsonOptions);
        }

        return response.Result ?? JsonSerializer.Serialize(new
        {
            success = true
        }, JsonOptions);
    }

    /// <summary>
    /// Forwards a command to the ExcelMCP Service without a session.
    /// Used for commands that don't require an active session (e.g., service.status).
    /// </summary>
    public static string ForwardToServiceNoSession(
        string command,
        object? args = null,
        int? timeoutSeconds = null)
    {
        var response = ServiceBridge.ServiceBridge.SendAsync(command, null, args, timeoutSeconds).GetAwaiter().GetResult();

        if (!response.Success)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = response.ErrorMessage ?? $"Command '{command}' failed",
                isError = true
            }, JsonOptions);
        }

        return response.Result ?? JsonSerializer.Serialize(new
        {
            success = true
        }, JsonOptions);
    }

    /// <summary>
    /// Executes a tool operation and serializes any exception using shared error formatting.
    /// Tracks tool usage telemetry (if enabled).
    /// </summary>
    /// <param name="toolName">Tool name for telemetry (e.g., "range").</param>
    /// <param name="actionName">Action string (kebab-case) included in error context.</param>
    /// <param name="operation">Synchronous operation to execute.</param>
    /// <param name="customHandler">Optional handler that can override default error serialization. Return null/empty to fall back to default.</param>
    /// <returns>Serialized JSON response.</returns>
    public static string ExecuteToolAction(
        string toolName,
        string actionName,
        Func<string> operation,
        Func<Exception, string?>? customHandler = null) =>
        ExecuteToolAction(toolName, actionName, null, operation, customHandler);

    /// <summary>
    /// Executes a tool operation and serializes any exception using shared error formatting.
    /// Tracks tool usage telemetry (if enabled).
    /// </summary>
    /// <param name="toolName">Tool name for telemetry (e.g., "range").</param>
    /// <param name="actionName">Action string (kebab-case) included in error context.</param>
    /// <param name="path">Optional Excel path for context in error messages.</param>
    /// <param name="operation">Synchronous operation to execute.</param>
    /// <param name="customHandler">Optional handler that can override default error serialization. Return null/empty to fall back to default.</param>
    /// <returns>Serialized JSON response.</returns>
    public static string ExecuteToolAction(
        string toolName,
        string actionName,
        string? path,
        Func<string> operation,
        Func<Exception, string?>? customHandler = null)
    {
        var stopwatch = Stopwatch.StartNew();
        var success = false;

        try
        {
            var result = operation();
            success = true;
            return result;
        }
        catch (Exception ex)
        {
            // Log COM exceptions to stderr for diagnostic capture
            if (ex is System.Runtime.InteropServices.COMException comEx)
            {
                Console.Error.WriteLine($"[ExcelMcp] COM Exception in {toolName}/{actionName}: HResult=0x{comEx.HResult:X8}, Message={comEx.Message}");
                if (ex.StackTrace != null)
                {
                    Console.Error.WriteLine($"[ExcelMcp] StackTrace: {ex.StackTrace[..Math.Min(500, ex.StackTrace.Length)]}");
                }
            }
            else if (ex.InnerException is System.Runtime.InteropServices.COMException innerComEx)
            {
                Console.Error.WriteLine($"[ExcelMcp] Inner COM Exception in {toolName}/{actionName}: HResult=0x{innerComEx.HResult:X8}, Message={innerComEx.Message}");
            }

            if (customHandler != null)
            {
                var custom = customHandler(ex);
                if (!string.IsNullOrWhiteSpace(custom))
                {
                    return custom!;
                }
            }

            return SerializeToolError(actionName, path, ex);
        }
        finally
        {
            stopwatch.Stop();
            ExcelMcpTelemetry.TrackToolInvocation(toolName, actionName, stopwatch.ElapsedMilliseconds, success, path);
        }
    }

    /// <summary>
    /// Validates that a path is a valid Windows absolute path.
    /// Returns null if valid, or a JSON error response if invalid.
    /// </summary>
    /// <param name="path">The path to validate</param>
    /// <returns>JSON error response if invalid, null if valid</returns>
    public static string? ValidateWindowsPath(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return null; // Let existing null checks handle this
        }

        // Use .NET's built-in check for fully qualified Windows paths
        // Returns false for Unix paths like /home/user/file.xlsx, relative paths like ./file.xlsx
        if (!Path.IsPathFullyQualified(path))
        {
            // Extract filename from the invalid path (works for both Unix and Windows separators)
            var fileName = Path.GetFileName(path.Replace('/', Path.DirectorySeparatorChar));
            if (string.IsNullOrEmpty(fileName))
            {
                fileName = "workbook.xlsx";
            }

            // Get user's actual Documents folder to provide a valid suggestion
            var documentsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var suggestedPath = Path.Combine(documentsFolder, fileName);

            var errorMessage = path.StartsWith('/')
                ? $"Invalid path format: '{path}' appears to be a Unix/Linux path. This server runs on Windows. Use: '{suggestedPath}'"
                : $"Invalid path format: '{path}' is not an absolute Windows path. Use: '{suggestedPath}'";

            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage,
                filePath = path,
                suggestedPath,
                documentsFolder,
                isError = true
            }, JsonOptions);
        }

        return null;
    }

    /// <summary>
    /// Serializes a tool error response with consistent structure.
    /// Uses camelCase property names matching JsonNamingPolicy: success, errorMessage, isError.
    /// Includes detailed COM exception info for diagnostics.
    /// </summary>
    /// <param name="actionName">Action string (kebab-case) included in message.</param>
    /// <param name="path">Optional Excel path context.</param>
    /// <param name="ex">Exception to serialize.</param>
    /// <returns>Serialized JSON error payload.</returns>
    public static string SerializeToolError(string actionName, string? path, Exception ex)
    {
        var errorMessage = path != null
            ? $"{actionName} failed for '{path}': {ex.Message}"
            : $"{actionName} failed: {ex.Message}";

        // Add detailed COM exception info for diagnostics
        string? exceptionType = ex.GetType().Name;
        string? hresult = null;
        string? innerError = null;

        if (ex is System.Runtime.InteropServices.COMException comEx)
        {
            hresult = $"0x{comEx.HResult:X8}";
            errorMessage += $" [COM Error: {hresult}]";
        }

        if (ex.InnerException != null)
        {
            innerError = ex.InnerException.Message;
            if (ex.InnerException is System.Runtime.InteropServices.COMException innerComEx)
            {
                innerError += $" [COM: 0x{innerComEx.HResult:X8}]";
            }
        }

        var payload = new
        {
            success = false,
            errorMessage,
            isError = true,
            exceptionType,
            hresult,
            innerError
        };

        return JsonSerializer.Serialize(payload, JsonOptions);
    }
}




