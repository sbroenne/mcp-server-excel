using System.Text.Json;
using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.Service;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

internal static class CliErrorOutput
{
    public static int WriteServiceError(ServiceResponse response)
    {
        Console.WriteLine(Serialize(response.ErrorMessage, response.ErrorCategory));
        return 1;
    }

    public static int WriteError(string errorMessage, string? errorCategory = null)
    {
        Console.WriteLine(Serialize(errorMessage, errorCategory));
        return 1;
    }

    private static string Serialize(string? errorMessage, string? errorCategory)
    {
        return JsonSerializer.Serialize(new ErrorEnvelope
        {
            Success = false,
            Error = errorMessage ?? "Unknown error.",
            ErrorCategory = errorCategory
        }, ServiceProtocol.JsonOptions);
    }

    private sealed class ErrorEnvelope
    {
        public bool Success { get; init; }

        public string Error { get; init; } = string.Empty;

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? ErrorCategory { get; init; }
    }
}
