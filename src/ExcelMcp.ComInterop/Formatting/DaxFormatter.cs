using System.Net.Http.Json;
using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.ComInterop.Formatting;

/// <summary>
/// Formats DAX (Data Analysis Expressions) code using the daxformatter.com API.
/// Provides automatic pretty-printing with proper indentation and line breaks.
/// </summary>
/// <remarks>
/// <para><b>Design Principles:</b></para>
/// <list type="bullet">
/// <item>Never throws exceptions - returns original DAX on any failure</item>
/// <item>Uses HttpClient for API calls to daxformatter.com</item>
/// <item>Gracefully handles network failures, API errors, and rate limiting</item>
/// <item>Formatting is best-effort - original DAX is always preserved if formatting fails</item>
/// </list>
/// <para><b>Usage:</b></para>
/// <code>
/// string formatted = await DaxFormatter.FormatAsync("CALCULATE(SUM(Sales[Amount]),FILTER(ALL(Calendar),Calendar[Year]=2024))");
/// // Returns formatted DAX with indentation, or original if formatting fails
/// </code>
/// </remarks>
public static class DaxFormatter
{
    private static readonly HttpClient _httpClient = new()
    {
        Timeout = TimeSpan.FromSeconds(10) // 10 second timeout for API calls
    };

    private const string DaxFormatterApiUrl = "https://www.daxformatter.com/api/daxformatter/DaxFormat";

    /// <summary>
    /// Formats DAX code using the daxformatter.com API.
    /// </summary>
    /// <param name="daxCode">The DAX code to format</param>
    /// <param name="cancellationToken">Cancellation token for the HTTP request</param>
    /// <returns>Formatted DAX code, or original code if formatting fails</returns>
    /// <remarks>
    /// This method NEVER throws exceptions. If formatting fails for any reason
    /// (network error, API error, timeout, invalid DAX), it returns the original code unchanged.
    /// This ensures that DAX operations never break due to formatting issues.
    /// </remarks>
    public static async Task<string> FormatAsync(string daxCode, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(daxCode))
            return daxCode;

        try
        {
            // Prepare request payload
            var request = new DaxFormatterRequest { Dax = daxCode };

            // Call daxformatter.com API
            var response = await _httpClient.PostAsJsonAsync(
                DaxFormatterApiUrl,
                request,
                cancellationToken).ConfigureAwait(false);

            // Check if request was successful
            if (!response.IsSuccessStatusCode)
            {
                // API returned error - return original DAX
                return daxCode;
            }

            // Parse response
            var result = await response.Content.ReadFromJsonAsync<DaxFormatterResponse>(
                cancellationToken: cancellationToken).ConfigureAwait(false);

            // Return formatted DAX if available, otherwise original
            return !string.IsNullOrWhiteSpace(result?.Formatted)
                ? result.Formatted
                : daxCode;
        }
        catch
        {
            // Any error (network, timeout, parsing, etc.) - return original DAX
            // This includes: HttpRequestException, TaskCanceledException, JsonException, etc.
            return daxCode;
        }
    }

    /// <summary>
    /// Request payload for daxformatter.com API
    /// </summary>
    private sealed class DaxFormatterRequest
    {
        [JsonPropertyName("Dax")]
        public string Dax { get; set; } = string.Empty;
    }

    /// <summary>
    /// Response from daxformatter.com API
    /// </summary>
    private sealed class DaxFormatterResponse
    {
        [JsonPropertyName("formatted")]
        public string? Formatted { get; set; }
    }
}
