using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Generated;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

/// <summary>
/// Uses shared operation-outcome rules for telemetry, including authoritative
/// nested failures, without allowing malformed payloads to interrupt reporting.
/// </summary>
internal static class ServiceResultOutcome
{
    internal static bool TryReadNegative(
        string? resultJson,
        out string? errorMessage,
        out string? errorCategory,
        string command = "invocation")
    {
        errorMessage = null;
        errorCategory = null;
        if (string.IsNullOrWhiteSpace(resultJson))
        {
            return false;
        }

        try
        {
            var outcome = ResultJson.Read(resultJson, command, contract: ServiceRegistry.GetResultContract(command));
            errorMessage = outcome.Error?["message"]?.GetValue<string>();
            errorCategory = outcome.Error?["category"]?.GetValue<string>();
            return !outcome.Success;
        }
        catch (JsonException exception)
        {
            errorMessage = exception.Message;
            errorCategory = "InvalidResponse";
            return true;
        }
    }
}
