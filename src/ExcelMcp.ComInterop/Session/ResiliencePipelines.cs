using System.Runtime.InteropServices;
using Polly;
using Polly.Retry;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Provides pre-configured resilience pipelines for Excel COM interop operations.
/// </summary>
internal static class ResiliencePipelines
{
    // Known COM HResults for transient busy/retry conditions
    private const int RPC_E_SERVERCALL_RETRYLATER = unchecked((int)0x8001010A); // -2147417846
    private const int RPC_E_CALL_REJECTED = unchecked((int)0x80010001);          // -2147418111

    /// <summary>
    /// Creates a retry pipeline for Excel.Quit() operations.
    /// Handles transient COM busy conditions with exponential backoff + jitter.
    /// </summary>
    /// <returns>Configured resilience pipeline</returns>
    public static ResiliencePipeline CreateExcelQuitPipeline()
    {
        return new ResiliencePipelineBuilder()
            .AddRetry(new RetryStrategyOptions
            {
                MaxRetryAttempts = 6,
                BackoffType = DelayBackoffType.Exponential,
                UseJitter = true,
                Delay = TimeSpan.FromMilliseconds(200),

                // Only retry on known transient COM busy errors
                ShouldHandle = new PredicateBuilder().Handle<COMException>(ex =>
                    ex.HResult == RPC_E_SERVERCALL_RETRYLATER ||
                    ex.HResult == RPC_E_CALL_REJECTED),

                // Log retry attempts for diagnostics
                OnRetry = args =>
                {
                    // Optional: logging will be done by caller
                    return ValueTask.CompletedTask;
                }
            })
            .Build();
    }
}
