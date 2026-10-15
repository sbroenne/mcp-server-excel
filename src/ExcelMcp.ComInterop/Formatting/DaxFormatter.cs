using Dax.Formatter;

namespace Sbroenne.ExcelMcp.ComInterop.Formatting;

/// <summary>
/// Formats DAX (Data Analysis Expressions) code using the remote Dax.Formatter service.
/// Call only after explicit user consent because the DAX formula is sent to daxformatter.com.
/// </summary>
/// <remarks>
/// <para><b>Design Principles:</b></para>
/// <list type="bullet">
/// <item>Returns original DAX for expected network, timeout, or response parsing failures</item>
/// <item>Uses official Dax.Formatter NuGet package (by SQLBI), which invokes a remote service</item>
/// <item>Gracefully handles network failures, API errors, and rate limiting</item>
/// <item>Expected service failures preserve the original DAX; cancellation and unexpected failures propagate</item>
/// </list>
/// <para><b>Usage:</b></para>
/// <code>
/// string formatted = await DaxFormatter.FormatAsync("CALCULATE(SUM(Sales[Amount]),FILTER(ALL(Calendar),Calendar[Year]=2024))");
/// // Returns formatted DAX with indentation, or original if formatting fails
/// </code>
/// <para><b>Performance:</b></para>
/// <list type="bullet">
/// <item>Network latency: Typically 100-500ms per API call</item>
/// <item>Singleton client instance shared across all calls for efficiency</item>
/// <item>10-second timeout prevents indefinite blocking</item>
/// <item>Expected service unavailability does not prevent workbook operations</item>
/// </list>
/// </remarks>
public static class DaxFormatter
{
    // Singleton instance - reused across all calls for better performance
    private static readonly DaxFormatterClient _formatterClient = new();

    /// <summary>
    /// Formats DAX code using the remote Dax.Formatter service.
    /// </summary>
    /// <param name="daxCode">The DAX code to format</param>
    /// <param name="cancellationToken">Cancellation token for the HTTP request</param>
    /// <returns>Formatted DAX code, or original code on an expected service failure</returns>
    /// <remarks>
    /// Expected network, response parsing, and internal timeout failures return the original
    /// code unchanged. Caller cancellation and unexpected programming failures propagate.
    /// </remarks>
    public static async Task<string> FormatAsync(string daxCode, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (string.IsNullOrWhiteSpace(daxCode))
            return daxCode;

        try
        {
            // Use timeout wrapper for 10 second limit
            using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeoutCts.CancelAfter(TimeSpan.FromSeconds(10));

            // Call official Dax.Formatter library
            var response = await _formatterClient.FormatAsync(daxCode, timeoutCts.Token)
                .ConfigureAwait(false);

            // Return formatted DAX if available, otherwise original
            return !string.IsNullOrWhiteSpace(response?.Formatted)
                ? response.Formatted
                : daxCode;
        }
        catch (HttpRequestException)
        {
            return daxCode;
        }
        catch (System.Text.Json.JsonException)
        {
            return daxCode;
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            return daxCode;
        }
    }
}
