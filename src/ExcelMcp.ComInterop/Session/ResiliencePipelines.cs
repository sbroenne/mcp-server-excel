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
    /// RPC_S_SERVER_UNAVAILABLE - The RPC server is unavailable. Excel process has died.
    /// FATAL ERROR - Do not retry, session must be cleaned up.
    /// </summary>
    public const int RPC_S_SERVER_UNAVAILABLE = unchecked((int)0x800706BA);     // -2147023174

    /// <summary>
    /// CO_E_SERVER_EXEC_FAILURE - COM class factory failed to start the server (Excel).
    /// Transient during session creation when system resources are constrained.
    /// </summary>
    public const int CO_E_SERVER_EXEC_FAILURE = unchecked((int)0x80080005);     // -2146959355

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

    /// <summary>
    /// Retry configuration for session creation (starting new Excel instances).
    /// Handles CO_E_SERVER_EXEC_FAILURE when system can't launch Excel temporarily.
    /// Also handles RPC_E_CALL_FAILED since a new process hasn't been contacted yet.
    /// </summary>
    private static readonly PipelineConfig SessionCreationConfig = new(
        MaxRetryAttempts: 3,
        DelayMs: 2000,
        AdditionalHResults: [CO_E_SERVER_EXEC_FAILURE, RPC_E_CALL_FAILED]);

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

    /// <summary>
    /// Creates a retry pipeline for session creation operations.
    /// Handles CO_E_SERVER_EXEC_FAILURE (0x80080005) and RPC_E_CALL_FAILED (0x800706BE)
    /// which occur when the system cannot start a new Excel process temporarily.
    /// </summary>
    /// <remarks>
    /// Unlike mid-session pipelines where RPC_E_CALL_FAILED is fatal (Excel died),
    /// during session creation these errors are transient — the system may be temporarily
    /// unable to launch a new Excel process due to resource constraints.
    /// </remarks>
    /// <returns>Configured resilience pipeline</returns>
    public static ResiliencePipeline CreateSessionCreationPipeline() => CreateSessionPipeline(SessionCreationConfig);

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

                // Only retry transient errors, NOT fatal RPC connection failures
                ShouldHandle = new PredicateBuilder().Handle<COMException>(ex =>
                    ex.HResult != RPC_E_CALL_FAILED &&
                    ex.HResult != RPC_S_SERVER_UNAVAILABLE &&
                    (ex.HResult == RPC_E_SERVERCALL_RETRYLATER ||
                     ex.HResult == RPC_E_CALL_REJECTED ||
                     config.AdditionalHResults.Contains(ex.HResult))),

                OnRetry = static _ => ValueTask.CompletedTask
            })
            .Build();
    }

    /// <summary>
    /// Creates a resilience pipeline for session creation.
    /// Unlike <see cref="CreatePipeline"/>, this does NOT exclude RPC_E_CALL_FAILED as fatal
    /// because during session creation there is no existing Excel process to have died.
    /// </summary>
    private static ResiliencePipeline CreateSessionPipeline(PipelineConfig config)
    {
        return new ResiliencePipelineBuilder()
            .AddRetry(new RetryStrategyOptions
            {
                MaxRetryAttempts = config.MaxRetryAttempts,
                BackoffType = DelayBackoffType.Exponential,
                UseJitter = true,
                Delay = TimeSpan.FromMilliseconds(config.DelayMs),

                ShouldHandle = new PredicateBuilder().Handle<COMException>(ex =>
                    ex.HResult == RPC_E_SERVERCALL_RETRYLATER ||
                    ex.HResult == RPC_E_CALL_REJECTED ||
                    config.AdditionalHResults.Contains(ex.HResult)),

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


