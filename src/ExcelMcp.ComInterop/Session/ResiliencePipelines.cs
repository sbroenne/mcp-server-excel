using System.Runtime.InteropServices;
using Polly;
using Polly.Retry;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Provides pre-configured resilience pipelines for Excel COM interop operations.
/// </summary>
public static class ResiliencePipelines
{
    // Known COM HResults for transient busy/retry conditions
    private const int RPC_E_SERVERCALL_RETRYLATER = unchecked((int)0x8001010A); // -2147417846
    private const int RPC_E_CALL_REJECTED = unchecked((int)0x80010001);          // -2147418111

    // Data Model specific error - intermittent failure during measure/table operations
    // GitHub Issue #315: https://github.com/sbroenne/mcp-server-excel/issues/315
    private const int DATA_MODEL_BUSY = unchecked((int)0x800AC472);              // -2146827150

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
                Delay = TimeSpan.FromMilliseconds(500),

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

    // Power Query specific error - failure when updating queries loaded to Data Model
    // GitHub Issue #316: https://github.com/sbroenne/mcp-server-excel/issues/316
    private const int POWER_QUERY_DATA_MODEL_ERROR = unchecked((int)0x800A03EC);  // -2146827284

    /// <summary>
    /// Creates a retry pipeline for Data Model operations (measures, relationships, tables).
    /// Handles intermittent 0x800AC472 errors with exponential backoff + jitter.
    /// </summary>
    /// <remarks>
    /// The 0x800AC472 error occurs intermittently when performing Data Model operations
    /// on workbooks with active Power Pivot models. The operation typically succeeds on retry.
    /// See GitHub Issue #315 for details.
    /// </remarks>
    /// <returns>Configured resilience pipeline</returns>
    public static ResiliencePipeline CreateDataModelPipeline()
    {
        return new ResiliencePipelineBuilder()
            .AddRetry(new RetryStrategyOptions
            {
                MaxRetryAttempts = 5,
                BackoffType = DelayBackoffType.Exponential,
                UseJitter = true,
                Delay = TimeSpan.FromMilliseconds(1000),

                // Retry on Data Model busy error and standard COM busy errors
                ShouldHandle = new PredicateBuilder().Handle<COMException>(ex =>
                    ex.HResult == DATA_MODEL_BUSY ||
                    ex.HResult == RPC_E_SERVERCALL_RETRYLATER ||
                    ex.HResult == RPC_E_CALL_REJECTED),

                OnRetry = args =>
                {
                    // Logging done by caller
                    return ValueTask.CompletedTask;
                }
            })
            .Build();
    }

    /// <summary>
    /// Creates a retry pipeline for Power Query update operations.
    /// Handles 0x800A03EC errors that can occur when updating queries loaded to Data Model.
    /// </summary>
    /// <remarks>
    /// The 0x800A03EC error can occur when updating Power Query M code for queries
    /// that are loaded to the Data Model. The error may be transient in some scenarios.
    /// See GitHub Issue #316 for details.
    /// </remarks>
    /// <returns>Configured resilience pipeline</returns>
    public static ResiliencePipeline CreatePowerQueryPipeline()
    {
        return new ResiliencePipelineBuilder()
            .AddRetry(new RetryStrategyOptions
            {
                MaxRetryAttempts = 5,
                BackoffType = DelayBackoffType.Exponential,
                UseJitter = true,
                Delay = TimeSpan.FromMilliseconds(1000),

                // Retry on Power Query Data Model error and standard COM busy errors
                ShouldHandle = new PredicateBuilder().Handle<COMException>(ex =>
                    ex.HResult == POWER_QUERY_DATA_MODEL_ERROR ||
                    ex.HResult == RPC_E_SERVERCALL_RETRYLATER ||
                    ex.HResult == RPC_E_CALL_REJECTED),

                OnRetry = args =>
                {
                    // Logging done by caller
                    return ValueTask.CompletedTask;
                }
            })
            .Build();
    }
}
