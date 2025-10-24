namespace Sbroenne.ExcelMcp.Core.Session;

/// <summary>
/// Exception thrown when the Excel instance pool is at maximum capacity.
/// Provides actionable guidance for LLMs to handle the situation.
/// </summary>
public class ExcelPoolCapacityException : InvalidOperationException
{
    /// <summary>
    /// Current number of active instances in the pool.
    /// </summary>
    public int ActiveInstances { get; }

    /// <summary>
    /// Maximum allowed instances in the pool.
    /// </summary>
    public int MaxInstances { get; }

    /// <summary>
    /// Idle timeout after which instances are automatically disposed.
    /// </summary>
    public TimeSpan IdleTimeout { get; }

    /// <summary>
    /// Suggested actions for the LLM to resolve the capacity issue.
    /// </summary>
    public List<string> SuggestedActions { get; }

    /// <summary>
    /// Initializes a new instance of the ExcelPoolCapacityException.
    /// </summary>
    /// <param name="activeInstances">Current number of active instances</param>
    /// <param name="maxInstances">Maximum allowed instances</param>
    /// <param name="idleTimeout">Idle timeout for automatic cleanup</param>
    public ExcelPoolCapacityException(int activeInstances, int maxInstances, TimeSpan idleTimeout)
        : base(BuildMessage(activeInstances, maxInstances, idleTimeout))
    {
        ActiveInstances = activeInstances;
        MaxInstances = maxInstances;
        IdleTimeout = idleTimeout;
        SuggestedActions = BuildSuggestedActions(idleTimeout);
    }

    private static string BuildMessage(int activeInstances, int maxInstances, TimeSpan idleTimeout)
    {
        return $"Excel instance pool is at maximum capacity ({activeInstances}/{maxInstances} instances active). " +
               $"Instances are automatically cleaned up after {idleTimeout.TotalSeconds:F0} seconds of inactivity. " +
               "Wait for idle instances to be cleaned up or close workbooks explicitly.";
    }

    private static List<string> BuildSuggestedActions(TimeSpan idleTimeout)
    {
        return new List<string>
        {
            $"Wait {idleTimeout.TotalSeconds:F0} seconds for idle instances to be automatically cleaned up",
            "Close workbooks you're no longer using with excel_file action='close-workbook'",
            "Check which files are currently open and close any you don't need",
            "Consider working on fewer files simultaneously (current pool limit is for system stability)"
        };
    }
}
