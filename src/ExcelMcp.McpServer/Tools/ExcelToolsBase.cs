using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Serialization;
using ModelContextProtocol.Protocol;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Sbroenne.ExcelMcp.McpServer.Telemetry;
using Sbroenne.ExcelMcp.Service;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>Converts shared Excel results into SDK results without replacing SDK invocation.</summary>
public static class ExcelToolsBase
{
    public static readonly JsonSerializerOptions JsonOptions = new()
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        Converters = { new JsonStringEnumConverter() }
    };

    public static async Task<string> ForwardToServiceAsync(
        ServiceBridge.ServiceBridge bridge,
        string command,
        string? sessionId,
        object? args,
        CancellationToken cancellationToken)
    {
        var response = await bridge.SendAsync(command, sessionId,
            args, cancellationToken: cancellationToken);
        var result = ReadServiceResponse(response, command);
        return result.ToToolJson();
    }

    internal static ResultJson ReadServiceResponse(ServiceResponse response, string command)
    {
        var contract = ServiceRegistry.GetResultContract(command);
        if (response.Success && response.Result is null && contract is { ReadOnly: true })
            throw new InvalidOperationException($"Read operation '{command}' returned no result.");
        return ResultJson.Read(response.Result, command, response.Success, response.ErrorMessage,
            response.ErrorCategory, response.SessionId, response.ExceptionType, response.HResult,
            response.InnerError, response.PartialState, contract);
    }

    internal static string SerializeServiceResponse(ServiceResponse response, string? filePath = null)
    {
        var result = ReadServiceResponse(response, response.Command ?? "file");
        if (filePath is not null && result.Error is not null)
            result.Error["filePath"] = filePath;
        return result.ToToolJson();
    }

    public static Task<CallToolResult> ExecuteToolActionAsync(
        string toolName,
        string actionName,
        Func<Task<string>> operation,
        CancellationToken cancellationToken,
        Func<string, CallToolResult>? resultFactory = null) =>
        ExecuteToolActionAsync(toolName, actionName, operation, cancellationToken,
            ExcelMcpTelemetry.TrackToolInvocation, resultFactory);

    internal static async Task<CallToolResult> ExecuteToolActionAsync(
        string toolName,
        string actionName,
        Func<Task<string>> operation,
        CancellationToken cancellationToken,
        Action<string, string, long, ToolInvocationResult> trackInvocation,
        Func<string, CallToolResult>? resultFactory = null)
    {
        var stopwatch = Stopwatch.StartNew();
        var invocation = new ToolInvocationResult(ToolInvocationOutcome.Failed, ToolFailureClass.Unclassified);
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            var json = await operation();
            cancellationToken.ThrowIfCancellationRequested();
            var command = ServiceRegistry.GetMcpCommand(toolName, actionName);
            var outcome = ResultJson.Read(json, command, contract: ServiceRegistry.GetResultContract(command));
            json = outcome.ToToolJson();
            invocation = ClassifyToolResponse(toolName, actionName, json);
            var result = resultFactory?.Invoke(json)
                ?? CreateToolResult(json, !outcome.Success);
            if (result.IsError is true && invocation.Outcome == ToolInvocationOutcome.Succeeded)
                invocation = new(ToolInvocationOutcome.Failed, ToolFailureClass.Unclassified);
            return result;
        }
        catch (ArgumentException ex)
        {
            var json = SerializeToolError(actionName, null, ex);
            invocation = ClassifyToolResponse(toolName, actionName, json);
            return CreateToolResult(json, isError: true);
        }
        catch (Exception ex)
        {
            var failureCategory = OperationFailureClassifier.Classify(ex);
            invocation = new(
                ToolInvocationOutcome.Failed,
                ex is JsonException
                    ? ToolFailureClass.InternalProductFault
                    : ClassifyFailure(failureCategory),
                ClassifyFailureCause(failureCategory));
            throw; // Let the SDK preserve cancellation and redact unexpected invocation errors.
        }
        finally
        {
            trackInvocation(toolName, actionName, stopwatch.ElapsedMilliseconds, invocation);
        }
    }

    internal static CallToolResult CreateToolResult(string json, bool? isError = null)
    {
        using var document = JsonDocument.Parse(json);
        var root = document.RootElement;
        var failed = !ResultJson.Read(json, "tool").Success;
        return new CallToolResult
        {
            IsError = isError ?? failed,
            Content = [new TextContentBlock { Text = json }],
            // Older MCP versions require an object here; retain existing text for scalar results.
            StructuredContent = root.ValueKind == JsonValueKind.Object
                ? root.Clone()
                : JsonSerializer.SerializeToElement(new { result = root }, JsonOptions)
        };
    }

    private static ToolInvocationResult ClassifyToolResponse(string toolName, string actionName, string response)
    {
        using var document = JsonDocument.Parse(response);
        var root = document.RootElement;
        if (root.ValueKind != JsonValueKind.Object)
            return new(ToolInvocationOutcome.Succeeded, null);

        var negative = root.TryGetProperty("success", out var success) && success.ValueKind == JsonValueKind.False;
        var error = root.TryGetProperty("isError", out var isError) && isError.ValueKind == JsonValueKind.True;
        if (negative && !error && toolName == "file_read" && actionName == "test")
            return new(ToolInvocationOutcome.ExpectedNegative, null);

        if (negative || error)
        {
            var category = root.TryGetProperty("errorCategory", out var property)
                && property.ValueKind == JsonValueKind.String ? property.GetString() : null;
            return new(
                ToolInvocationOutcome.Failed,
                ClassifyFailure(category),
                ClassifyFailureCause(category));
        }
        return new(ToolInvocationOutcome.Succeeded, null);
    }

    private static ToolFailureCause? ClassifyFailureCause(string? category) => category switch
    {
        "Timeout" => ToolFailureCause.Timeout,
        "Cancelled" => ToolFailureCause.Cancellation,
        _ => null
    };

    private static ToolFailureClass ClassifyFailure(string? category) => category switch
    {
        "InvalidInput" or "SessionNotFound" or "SessionUnavailable" or "NotFound" or "Conflict"
            or "Syntax" or "Expression" or "Prerequisite" => ToolFailureClass.InputState,
        "Privacy" or "PrivacyFirewall" or "Authentication" or "Connectivity" or "Permissions" or "DependencyUnavailable"
            => ToolFailureClass.ExternalDependency,
        "Timeout" or "Cancelled" or "SessionInvalidated" => ToolFailureClass.TimeoutCancellation,
        "ComInterop" or "ExcelProcessDied" or "Cleanup" => ToolFailureClass.ExcelRuntime,
        "ServiceUnavailable" or "ServiceStartup" or "InvalidResponse" => ToolFailureClass.InternalProductFault,
        _ => ToolFailureClass.Unclassified
    };

    public static string? ValidateWindowsPath(string? path)
    {
        if (string.IsNullOrWhiteSpace(path) || Path.IsPathFullyQualified(path))
            return null;

        var fileName = Path.GetFileName(path.Replace('/', Path.DirectorySeparatorChar));
        if (string.IsNullOrEmpty(fileName))
            fileName = "workbook.xlsx";
        var documentsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var suggestedPath = Path.Combine(documentsFolder, fileName);
        var errorMessage = path.StartsWith('/')
            ? $"Invalid path format: '{path}' appears to be a Unix/Linux path. This server runs on Windows. Use: '{suggestedPath}'"
            : $"Invalid path format: '{path}' is not an absolute Windows path. Use: '{suggestedPath}'";
        return JsonSerializer.Serialize(new
        {
            success = false,
            errorMessage,
            errorCategory = "InvalidInput",
            filePath = path,
            suggestedPath,
            documentsFolder,
            isError = true
        }, JsonOptions);
    }

    public static string SerializeToolError(string actionName, string? path, Exception ex)
    {
        var errorMessage = path is null
            ? $"{actionName} failed: {ex.Message}"
            : $"{actionName} failed for '{path}': {ex.Message}";
        return ResultJson.Read(null, actionName, delivered: false, message: errorMessage,
            category: OperationFailureClassifier.Classify(ex), exceptionType: ex.GetType().Name,
            hresult: OperationFailureClassifier.GetComHResult(ex), innerError: ex.InnerException?.Message).ToToolJson();
    }
}
