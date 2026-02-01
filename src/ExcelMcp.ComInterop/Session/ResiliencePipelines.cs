using System.Runtime.InteropServices;
using Polly;
using Polly.Retry;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Provides pre-configured resilience pipelines for Excel COM interop operations.
/// </summary>
public static class ResiliencePipelines
{
    #region COM HResult Constants

    /// <summary>
    /// RPC_E_SERVERCALL_RETRYLATER - COM server is busy, retry later.
    /// </summary>
    public const int RPC_E_SERVERCALL_RETRYLATER = unchecked((int)0x8001010A); // -2147417846

    /// <summary>
    /// RPC_E_CALL_REJECTED - COM call was rejected.
    /// </summary>
    public const int RPC_E_CALL_REJECTED = unchecked((int)0x80010001);          // -2147418111

    /// <summary>
    /// RPC_E_CALL_FAILED - RPC connection failed. Excel is unreachable.
    /// FATAL ERROR - Do not retry, Excel process must be force-killed.
    /// </summary>
    public const int RPC_E_CALL_FAILED = unchecked((int)0x800706BE);            // -2147023170

    /// <summary>
    /// Data Model specific error - intermittent failure during measure/table operations.
    /// See GitHub Issue #315.
    /// </summary>
    public const int DATA_MODEL_BUSY = unchecked((int)0x800AC472);              // -2146827150

    #endregion

    #region Pipeline Configuration

    /// <summary>
    /// Default retry configuration for standard COM busy operations.
    /// </summary>
    private static readonly PipelineConfig DefaultComConfig = new(
        MaxRetryAttempts: 6,
        DelayMs: 500,
        AdditionalHResults: []);

    /// <summary>
    /// Retry configuration for Data Model operations.
    /// </summary>
    private static readonly PipelineConfig DataModelConfig = new(
        MaxRetryAttempts: 5,
        DelayMs: 1000,
        AdditionalHResults: [DATA_MODEL_BUSY]);

    #endregion

    #region Factory Methods

    /// <summary>
    /// Creates a retry pipeline for Excel.Quit() operations.
    /// Handles transient COM busy conditions with exponential backoff + jitter.
    /// </summary>
    /// <returns>Configured resilience pipeline</returns>
    public static ResiliencePipeline CreateExcelQuitPipeline() => CreatePipeline(DefaultComConfig);

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
    public static ResiliencePipeline CreateDataModelPipeline() => CreatePipeline(DataModelConfig);

    #endregion

    #region Private Implementation

    /// <summary>
    /// Creates a resilience pipeline with the specified configuration.
    /// </summary>
    private static ResiliencePipeline CreatePipeline(PipelineConfig config)
    {
        return new ResiliencePipelineBuilder()
            .AddRetry(new RetryStrategyOptions
            {
                MaxRetryAttempts = config.MaxRetryAttempts,
                BackoffType = DelayBackoffType.Exponential,
                UseJitter = true,
                Delay = TimeSpan.FromMilliseconds(config.DelayMs),

                ShouldHandle = new PredicateBuilder().Handle<COMException>(ex =>
                    // Only retry transient errors, NOT fatal RPC connection failures
                    ex.HResult != RPC_E_CALL_FAILED &&
                    (ex.HResult == RPC_E_SERVERCALL_RETRYLATER ||
                     ex.HResult == RPC_E_CALL_REJECTED ||
                     config.AdditionalHResults.Contains(ex.HResult))),

                OnRetry = static _ => ValueTask.CompletedTask
            })
            .Build();
    }

    /// <summary>
    /// Configuration record for pipeline creation.
    /// </summary>
    private sealed record PipelineConfig(
        int MaxRetryAttempts,
        int DelayMs,
        int[] AdditionalHResults);

    #endregion
}
