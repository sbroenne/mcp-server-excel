namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Describes one immutable workbook savepoint owned by an active session.
/// </summary>
/// <param name="Name">Caller-provided savepoint name.</param>
/// <param name="CreatedAtUtc">UTC creation time.</param>
/// <param name="SizeBytes">Snapshot size in bytes.</param>
/// <param name="WorkbookPath">Workbook path captured by the savepoint.</param>
public sealed record WorkbookSavepointDescriptor(
    string Name,
    DateTime CreatedAtUtc,
    long SizeBytes,
    string WorkbookPath);

/// <summary>
/// Describes a completed savepoint rollback.
/// </summary>
/// <param name="SessionId">Public session identifier retained after rollback.</param>
/// <param name="Name">Rolled-back savepoint name.</param>
/// <param name="WorkbookPath">Restored workbook path.</param>
/// <param name="RestoredAtUtc">UTC completion time.</param>
public sealed record WorkbookSavepointRollback(
    string SessionId,
    string Name,
    string WorkbookPath,
    DateTime RestoredAtUtc);

/// <summary>
/// Savepoint storage limits enforced by one service process.
/// </summary>
/// <param name="MaxSavepointsPerSession">Maximum retained savepoints per session.</param>
/// <param name="MaxBytesPerSession">Maximum retained bytes per session.</param>
/// <param name="MaxBytesPerProcess">Maximum retained bytes across the process.</param>
public sealed record WorkbookSavepointLimits(
    int MaxSavepointsPerSession,
    long MaxBytesPerSession,
    long MaxBytesPerProcess);
