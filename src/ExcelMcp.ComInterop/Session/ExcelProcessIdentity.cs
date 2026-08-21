namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Identifies an Excel process by PID and creation time so PID reuse cannot
/// transfer ownership to an unrelated process.
/// </summary>
public readonly record struct ExcelProcessIdentity(
    int ProcessId,
    long StartedAtUtcFileTime);
