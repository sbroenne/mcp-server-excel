using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.Service;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

internal static class CliErrorOutput
{
    public static int WriteException(Exception ex, string? errorCategory = null)
    {
        Console.WriteLine(Serialize(
            ex.Message,
            errorCategory,
            null,
            null,
            ex.GetType().Name,
            null,
            ex.InnerException?.Message,
            null,
            null));
        return 1;
    }

    public static int WriteServiceError(ServiceResponse response)
    {
        Console.WriteLine(Serialize(
            response.ErrorMessage,
            response.ErrorCategory,
            response.Command,
            response.SessionId,
            response.ExceptionType,
            response.HResult,
            response.InnerError,
            null,
            null));
        return 1;
    }

    public static int WriteServiceErrorWithResult(ServiceResponse response)
    {
        if (string.IsNullOrWhiteSpace(response.Result))
        {
            return WriteServiceError(response);
        }

        try
        {
            var result = JsonNode.Parse(response.Result) as JsonObject;
            if (result == null)
            {
                return WriteServiceError(response);
            }

            result["success"] = false;
            result["error"] = response.ErrorMessage ?? "Unknown error.";
            result["errorMessage"] = response.ErrorMessage ?? "Unknown error.";
            result["errorCategory"] = response.ErrorCategory;
            result["command"] = response.Command;
            result["sessionId"] ??= response.SessionId;
            result["isError"] = true;
            result["exceptionType"] = response.ExceptionType;
            result["hresult"] = response.HResult;
            result["innerError"] = response.InnerError;
            Console.WriteLine(result.ToJsonString(ServiceProtocol.JsonOptions));
            return 1;
        }
        catch (JsonException)
        {
            return WriteServiceError(response);
        }
    }

    public static int WriteDaemonError(
        ServiceResponse response,
        string daemonState,
        bool running)
    {
        Console.WriteLine(Serialize(
            response.ErrorMessage,
            response.ErrorCategory,
            response.Command,
            response.SessionId,
            response.ExceptionType,
            response.HResult,
            response.InnerError,
            daemonState,
            running));
        return 1;
    }

    public static int WriteError(string errorMessage, string? errorCategory = null)
    {
        Console.WriteLine(Serialize(errorMessage, errorCategory, null, null, null, null, null, null, null));
        return 1;
    }

    private static string Serialize(
        string? errorMessage,
        string? errorCategory,
        string? command,
        string? sessionId,
        string? exceptionType,
        string? hresult,
        string? innerError,
        string? daemonState,
        bool? running)
    {
        return JsonSerializer.Serialize(new ErrorEnvelope
        {
            Success = false,
            Error = errorMessage ?? "Unknown error.",
            ErrorMessage = errorMessage ?? "Unknown error.",
            ErrorCategory = errorCategory,
            Command = command,
            SessionId = sessionId,
            IsError = true,
            ExceptionType = exceptionType,
            HResult = hresult,
            InnerError = innerError,
            DaemonState = daemonState,
            Running = running
        }, ServiceProtocol.JsonOptions);
    }

    private sealed class ErrorEnvelope
    {
        public bool Success { get; init; }

        public string Error { get; init; } = string.Empty;

        public string ErrorMessage { get; init; } = string.Empty;

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? ErrorCategory { get; init; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? Command { get; init; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? SessionId { get; init; }

        public bool IsError { get; init; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? ExceptionType { get; init; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        [JsonPropertyName("hresult")]
        public string? HResult { get; init; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? InnerError { get; init; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? DaemonState { get; init; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public bool? Running { get; init; }
    }
}
