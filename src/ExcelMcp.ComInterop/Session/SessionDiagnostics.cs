namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Diagnostic tracing for Excel COM session operations.
/// </summary>
public static class SessionDiagnostics
{
    private const string DiagnosticsEnvVar = "EXCELMCP_DIAGNOSTICS";

    /// <summary>
    /// Gets whether diagnostic tracing is enabled via EXCELMCP_DIAGNOSTICS environment variable.
    /// </summary>
    public static bool IsEnabled =>
        string.Equals(Environment.GetEnvironmentVariable(DiagnosticsEnvVar), "1", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(Environment.GetEnvironmentVariable(DiagnosticsEnvVar), "true", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Writes a diagnostic message to stderr if diagnostics are enabled.
    /// </summary>
    /// <param name="message">Message to write</param>
    public static void WriteStdErr(string message)
    {
        if (IsEnabled)
        {
            Console.Error.WriteLine(message);
        }
    }
}
