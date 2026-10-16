using System.Net.Http.Json;
using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.ComInterop.Formatting;

/// <summary>
/// Formats Power Query M code using the remote powerqueryformatter.com API.
/// Call only after explicit user consent because the M code is sent to the service.
/// </summary>
/// <remarks>
/// <para><b>Design Principles:</b></para>
/// <list type="bullet">
/// <item>Returns original M code for expected network, timeout, or response parsing failures</item>
/// <item>Uses powerqueryformatter.com API (by mogularGmbH, MIT License)</item>
/// <item>Gracefully handles network failures, API errors, and rate limiting</item>
/// <item>Expected service failures preserve the original M code; cancellation and unexpected failures propagate</item>
/// </list>
/// <para><b>Usage:</b></para>
/// <code>
/// string formatted = await MCodeFormatter.FormatAsync("let Source=Excel.CurrentWorkbook() in Source");
/// // Returns formatted M code with indentation, or original if formatting fails
/// </code>
/// <para><b>Performance:</b></para>
/// <list type="bullet">
/// <item>Network latency: Typically 100-500ms per API call</item>
/// <item>Singleton HttpClient instance shared across all calls for efficiency</item>
/// <item>10-second timeout prevents indefinite blocking</item>
/// <item>Expected service unavailability does not prevent workbook operations</item>
/// </list>
/// <para><b>API Reference:</b></para>
/// <list type="bullet">
/// <item>Endpoint: https://m-formatter.azurewebsites.net/api/v2</item>
/// <item>Method: POST with JSON body</item>
/// <item>Source: https://github.com/mogulargmbh/m-formatter (MIT License)</item>
/// </list>
/// </remarks>
public static class MCodeFormatter
{
    private const string ApiEndpoint = "https://m-formatter.azurewebsites.net/api/v2";

    // Singleton HttpClient - reused across all calls for better performance
    // HttpClient is designed to be reused and is thread-safe
    private static readonly HttpClient _httpClient = new()
    {
        Timeout = TimeSpan.FromSeconds(15) // Overall timeout as safety net
    };

    /// <summary>
    /// Formats Power Query M code using the remote powerqueryformatter.com API.
    /// </summary>
    /// <param name="mCode">The M code to format</param>
    /// <param name="cancellationToken">Cancellation token for the HTTP request</param>
    /// <returns>Formatted M code, or original code on an expected service failure</returns>
    /// <remarks>
    /// Expected network, response parsing, and internal timeout failures return the original
    /// code unchanged. Caller cancellation and unexpected programming failures propagate.
    /// </remarks>
    public static async Task<string> FormatAsync(string mCode, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (string.IsNullOrWhiteSpace(mCode))
            return mCode;

        try
        {
            // Use timeout wrapper for 10 second limit
            using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeoutCts.CancelAfter(TimeSpan.FromSeconds(10));

            // Prepare request
            var request = new MFormatterRequest { Code = mCode, ResultType = "text" };

            // Call the API
            using var response = await _httpClient.PostAsJsonAsync(ApiEndpoint, request, timeoutCts.Token)
                .ConfigureAwait(false);

            response.EnsureSuccessStatusCode();

            // Parse response
            var result = await response.Content.ReadFromJsonAsync<MFormatterResponse>(timeoutCts.Token)
                .ConfigureAwait(false);

            // Return formatted M code if successful, otherwise original
            return result is { Success: true } && !string.IsNullOrWhiteSpace(result.Result)
                ? result.Result
                : mCode;
        }
        catch (HttpRequestException)
        {
            return mCode;
        }
        catch (System.Text.Json.JsonException)
        {
            return mCode;
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            return mCode;
        }
    }

    /// <summary>
    /// Request payload for the M-Formatter API.
    /// </summary>
    private sealed class MFormatterRequest
    {
        [JsonPropertyName("code")]
        public required string Code { get; init; }

        [JsonPropertyName("resultType")]
        public required string ResultType { get; init; }
    }

    /// <summary>
    /// Response payload from the M-Formatter API.
    /// </summary>
    private sealed class MFormatterResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; init; }

        [JsonPropertyName("result")]
        public string? Result { get; init; }

        [JsonPropertyName("errors")]
        public object? Errors { get; init; }
    }
}
