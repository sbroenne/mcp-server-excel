using System.Text.Json;
using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Sbroenne.ExcelMcp.Service;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

internal static class CliErrorOutput
{
    public static int WriteException(Exception ex, string? errorCategory = null)
    {
        CliCommandRuntime.Current.Output.WriteLine(Serialize(
            ex.Message,
            errorCategory ?? OperationFailureClassifier.Classify(ex),
            null,
            null,
            ex.GetType().Name,
            OperationFailureClassifier.GetComHResult(ex),
            ex.InnerException?.Message,
            null,
            null));
        return 1;
    }

    public static int WriteServiceError(ServiceResponse response) =>
        CliResultOutput.Write(response);

    public static int WriteDaemonError(
        ServiceResponse response,
        string daemonState,
        bool running,
        IReadOnlyList<OwnedProcessCleanup.RetainedExcelProcess>? retainedExcelProcesses = null,
        string? retainedExcelError = null)
    {
        CliCommandRuntime.Current.Output.WriteLine(Serialize(
            response.ErrorMessage,
            response.ErrorCategory,
            response.Command,
            response.SessionId,
            response.ExceptionType,
            response.HResult,
            response.InnerError,
            daemonState,
            running,
            response.PartialState,
            retainedExcelProcesses,
            retainedExcelError));
        return 1;
    }

    public static int WriteError(string errorMessage, string? errorCategory = "InvalidInput")
    {
        CliCommandRuntime.Current.Output.WriteLine(
            Serialize(errorMessage, errorCategory, null, null, null, null, null, null, null));
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
        bool? running,
        string? partialState = null,
        IReadOnlyList<OwnedProcessCleanup.RetainedExcelProcess>? retainedExcelProcesses = null,
        string? retainedExcelError = null)
    {
        command ??= CliCommandRuntime.Current.CommandName;
        var json = JsonSerializer.Serialize(new ErrorEnvelope
        {
            Success = false,
            ErrorMessage = errorMessage ?? "Unknown error.",
            ErrorCategory = errorCategory,
            Command = command,
            SessionId = sessionId,
            IsError = true,
            ExceptionType = exceptionType,
            HResult = hresult,
            InnerError = innerError,
            PartialState = partialState is null ? null : JsonSerializer.Deserialize<JsonElement>(partialState),
            DaemonState = daemonState,
            Running = running,
            RetainedExcelProcessCount = retainedExcelProcesses?.Count,
            RetainedExcelProcesses = retainedExcelProcesses,
            RetainedExcelError = retainedExcelError,
            RecoveryMessage = retainedExcelProcesses?.Count > 0 ? OwnedProcessCleanup.RecoveryMessage : null
        }, ServiceProtocol.JsonOptions);
        return ResultJson.Read(json, command).ToCliJson(
            command, CliCommandRuntime.Current.OperationTimer.ElapsedMilliseconds);
    }

    private sealed class ErrorEnvelope
    {
        public bool Success { get; init; }

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
        public JsonElement? PartialState { get; init; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? DaemonState { get; init; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public bool? Running { get; init; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public int? RetainedExcelProcessCount { get; init; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public IReadOnlyList<OwnedProcessCleanup.RetainedExcelProcess>? RetainedExcelProcesses { get; init; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? RetainedExcelError { get; init; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? RecoveryMessage { get; init; }
    }
}
