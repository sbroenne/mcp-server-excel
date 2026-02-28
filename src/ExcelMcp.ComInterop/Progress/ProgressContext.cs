namespace Sbroenne.ExcelMcp.ComInterop;

/// <summary>
/// Ambient progress context using AsyncLocal.
/// Set by the MCP layer before invoking Core commands;
/// read by the generated DispatchToCore to inject into Core methods that accept IProgress.
/// </summary>
public static class ProgressContext
{
    private static readonly AsyncLocal<IProgress<ProgressInfo>?> CurrentValue = new();

    /// <summary>
    /// Gets or sets the current progress reporter for the async flow.
    /// </summary>
    public static IProgress<ProgressInfo>? Current
    {
        get => CurrentValue.Value;
        set => CurrentValue.Value = value;
    }
}
