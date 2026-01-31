using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Manages active Excel sessions for the MCP server and CLI.
/// Maps user-facing sessionId to internal IExcelBatch instances.
/// </summary>
/// <remarks>
/// <para><b>Concurrency Model:</b></para>
/// <list type="bullet">
/// <item><b>Within-session operations are SERIAL:</b> Each session queues operations on one STA thread</item>
/// <item><b>Between-session operations CAN be parallel:</b> Different sessions = different Excel processes</item>
/// <item><b>Same-file prevention:</b> Cannot open the same file in multiple sessions (matches Excel UI behavior)</item>
/// </list>
/// <para><b>Resource Limits:</b></para>
/// <list type="bullet">
/// <item>Each session = one Excel.Application process (~50-100MB+ memory)</item>
/// <item>Recommended maximum: 3-5 concurrent sessions on typical desktop machines</item>
/// <item>Always close sessions promptly to free resources</item>
/// </list>
/// </remarks>
public sealed class SessionManager : IDisposable
{
    private readonly ConcurrentDictionary<string, IExcelBatch> _activeSessions = new();
    private readonly ConcurrentDictionary<string, string> _activeFilePaths = new();
    private readonly ConcurrentDictionary<string, string> _sessionFilePaths = new(StringComparer.OrdinalIgnoreCase);
    private readonly ConcurrentDictionary<string, int> _activeOperationCounts = new();
    private readonly ConcurrentDictionary<string, bool> _showExcelFlags = new();
    private bool _disposed;

    /// <summary>
    /// Creates a new session for the specified Excel file.
    /// </summary>
    /// <param name="filePath">Path to the Excel file to open</param>
    /// <param name="showExcel">Whether to show the Excel window (default: false for background automation)</param>
    /// <param name="operationTimeout">Maximum time for any operation in this session (default: 5 minutes)</param>
    /// <returns>Unique session ID for this session</returns>
    /// <exception cref="FileNotFoundException">File does not exist</exception>
    /// <exception cref="InvalidOperationException">Failed to create session or file already open in another session</exception>
    /// <remarks>
    /// <para><b>Resource Impact:</b> Creates a new Excel.Application process (~50-100MB+ memory).</para>
    /// <para><b>Same-file prevention:</b> Throws if file is already open in another session.</para>
    /// <para><b>Concurrency:</b> You can create multiple sessions for DIFFERENT files. Operations within each session execute serially.</para>
    /// </remarks>
    public string CreateSession(string filePath, bool showExcel = false, TimeSpan? operationTimeout = null)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException($"Excel file not found: {filePath}", filePath);
        }

        // Normalize file path for comparison
        string normalizedPath = Path.GetFullPath(filePath);

        // Check if file is already open in another session
        if (_activeFilePaths.ContainsKey(normalizedPath))
        {
            throw new InvalidOperationException($"File '{filePath}' is already open in another session. Excel cannot open the same file multiple times.");
        }

        // Generate unique session ID
        string sessionId = Guid.NewGuid().ToString("N");

        IExcelBatch? batch = null;
        try
        {
            // Create batch session using Core API
            batch = ExcelSession.BeginBatch(showExcel, operationTimeout, filePath);

            // Store in active sessions
            if (!_activeSessions.TryAdd(sessionId, batch))
            {
                throw new InvalidOperationException($"Session ID collision: {sessionId}");
            }

            // Track the file path
            if (!_activeFilePaths.TryAdd(normalizedPath, sessionId))
            {
                // Cleanup if file path tracking fails
                _activeSessions.TryRemove(sessionId, out _);
                throw new InvalidOperationException($"Failed to track file path for session: {sessionId}");
            }

            if (!_sessionFilePaths.TryAdd(sessionId, normalizedPath))
            {
                _activeSessions.TryRemove(sessionId, out _);
                _activeFilePaths.TryRemove(normalizedPath, out _);
                throw new InvalidOperationException($"Failed to record session metadata for: {sessionId}");
            }

            // Initialize operation counter and showExcel flag
            _activeOperationCounts[sessionId] = 0;
            _showExcelFlags[sessionId] = showExcel;

            // Success - transfer ownership to dictionary
            var result = sessionId;
            batch = null;  // Prevent disposal in finally
            return result;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to create session for '{filePath}': {ex.Message}", ex);
        }
        finally
        {
            // Dispose batch only if we didn't successfully add it to dictionary
            batch?.Dispose();
        }
    }

    /// <summary>
    /// Creates a new Excel file and opens a session for it in one operation.
    /// This is the preferred method for creating new workbooks with sessions.
    /// </summary>
    /// <param name="filePath">Path for the new Excel file (.xlsx or .xlsm)</param>
    /// <param name="showExcel">Whether to show the Excel window (default: false)</param>
    /// <param name="operationTimeout">Maximum time for any operation in this session (default: 5 minutes)</param>
    /// <returns>Unique session ID for this session</returns>
    /// <exception cref="InvalidOperationException">File already exists, or failed to create session</exception>
    /// <exception cref="DirectoryNotFoundException">Target directory does not exist</exception>
    /// <remarks>
    /// <para><b>Single Excel Start:</b> This method starts Excel only once, creating the file and session together.</para>
    /// <para><b>File Format:</b> Determined by extension - .xlsm creates macro-enabled workbook.</para>
    /// <para><b>Directory:</b> Target directory must exist - will not be created automatically.</para>
    /// </remarks>
    public string CreateSessionForNewFile(string filePath, bool showExcel = false, TimeSpan? operationTimeout = null)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        string normalizedPath = Path.GetFullPath(filePath);

        // Validate extension
        string extension = Path.GetExtension(normalizedPath).ToLowerInvariant();
        if (extension is not (".xlsx" or ".xlsm"))
        {
            throw new ArgumentException($"Invalid file extension '{extension}'. Only .xlsx and .xlsm are supported.");
        }

        // Check if file already exists
        if (File.Exists(normalizedPath))
        {
            throw new InvalidOperationException($"File already exists: {normalizedPath}. Use CreateSession to open existing files.");
        }

        // Check if file is already open in another session
        if (_activeFilePaths.ContainsKey(normalizedPath))
        {
            throw new InvalidOperationException($"File '{filePath}' is already open in another session.");
        }

        // Generate unique session ID
        string sessionId = Guid.NewGuid().ToString("N");
        bool isMacroEnabled = extension == ".xlsm";

        ExcelBatch? batch = null;
        try
        {
            // Create new workbook and keep session open
            batch = ExcelBatch.CreateNewWorkbook(normalizedPath, isMacroEnabled, logger: null, showExcel: showExcel, operationTimeout: operationTimeout);

            // Store in active sessions
            if (!_activeSessions.TryAdd(sessionId, batch))
            {
                throw new InvalidOperationException($"Session ID collision: {sessionId}");
            }

            // Track the file path
            if (!_activeFilePaths.TryAdd(normalizedPath, sessionId))
            {
                _activeSessions.TryRemove(sessionId, out _);
                throw new InvalidOperationException($"Failed to track file path for session: {sessionId}");
            }

            if (!_sessionFilePaths.TryAdd(sessionId, normalizedPath))
            {
                _activeSessions.TryRemove(sessionId, out _);
                _activeFilePaths.TryRemove(normalizedPath, out _);
                throw new InvalidOperationException($"Failed to record session metadata for: {sessionId}");
            }

            // Initialize operation counter and showExcel flag
            _activeOperationCounts[sessionId] = 0;
            _showExcelFlags[sessionId] = showExcel;

            // Success - transfer ownership to dictionary
            var result = sessionId;
            batch = null;  // Prevent disposal in finally
            return result;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to create session for new file '{filePath}': {ex.Message}", ex);
        }
        finally
        {
            // Dispose batch only if we didn't successfully add it to dictionary
            batch?.Dispose();
        }
    }



    /// <summary>
    /// Gets an active session by ID.
    /// </summary>
    /// <param name="sessionId">Session ID returned from CreateSession</param>
    /// <returns>IExcelBatch instance, or null if session not found</returns>
    public IExcelBatch? GetSession(string sessionId)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return null;
        }

        _activeSessions.TryGetValue(sessionId, out var batch);
        return batch;
    }

    /// <summary>
    /// Increments the active operation count for a session.
    /// Call this when starting an operation on the session.
    /// </summary>
    /// <param name="sessionId">Session ID</param>
    public void BeginOperation(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId)) return;
        _activeOperationCounts.AddOrUpdate(sessionId, 1, (_, count) => count + 1);
    }

    /// <summary>
    /// Decrements the active operation count for a session.
    /// Call this when an operation completes (success or failure).
    /// </summary>
    /// <param name="sessionId">Session ID</param>
    public void EndOperation(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId)) return;
        _activeOperationCounts.AddOrUpdate(sessionId, 0, (_, count) => Math.Max(0, count - 1));
    }

    /// <summary>
    /// Gets the number of active operations for a session.
    /// </summary>
    /// <param name="sessionId">Session ID</param>
    /// <returns>Number of active operations, or 0 if session not found</returns>
    public int GetActiveOperationCount(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId)) return 0;
        return _activeOperationCounts.TryGetValue(sessionId, out var count) ? count : 0;
    }

    /// <summary>
    /// Gets whether Excel is visible for a session.
    /// </summary>
    /// <param name="sessionId">Session ID</param>
    /// <returns>True if showExcel was true when session was created</returns>
    public bool IsExcelVisible(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId)) return false;
        return _showExcelFlags.TryGetValue(sessionId, out var visible) && visible;
    }

    /// <summary>
    /// Validates whether a session can be closed safely.
    /// Returns information about blocking conditions.
    /// </summary>
    /// <param name="sessionId">Session ID</param>
    /// <returns>Validation result with details about any blocking conditions</returns>
    public CloseValidationResult ValidateClose(string sessionId)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return new CloseValidationResult(false, false, 0, "Session ID is required");
        }

        if (!_activeSessions.ContainsKey(sessionId))
        {
            return new CloseValidationResult(false, false, 0, $"Session '{sessionId}' not found");
        }

        var activeOps = GetActiveOperationCount(sessionId);
        var isVisible = IsExcelVisible(sessionId);

        if (activeOps > 0)
        {
            return new CloseValidationResult(true, isVisible, activeOps,
                $"Cannot close: {activeOps} operation(s) still running. Wait for operations to complete before closing.");
        }

        return new CloseValidationResult(true, isVisible, 0, null);
    }

    /// <summary>
    /// Closes the specified session with optional save.
    /// If save is true, saves changes before closing to ensure atomic operation.
    /// </summary>
    /// <param name="sessionId">Session ID</param>
    /// <param name="save">Whether to save changes before closing (default: false)</param>
    /// <param name="force">Force close even if operations are running (default: false)</param>
    /// <returns>True if session was found and closed, false if session not found</returns>
    /// <exception cref="InvalidOperationException">Save operation failed or operations still running</exception>
    public bool CloseSession(string sessionId, bool save = false, bool force = false)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return false;
        }

        // Check for running operations (unless force is true)
        if (!force)
        {
            var activeOps = GetActiveOperationCount(sessionId);
            if (activeOps > 0)
            {
                throw new InvalidOperationException(
                    $"Cannot close session '{sessionId}': {activeOps} operation(s) still running. " +
                    "Wait for all operations to complete before closing, or use force=true to close anyway.");
            }
        }

        // Save first if requested (blocks until complete)
        if (save)
        {
            var batch = GetSession(sessionId);
            if (batch != null)
            {
                try
                {
                    batch.Save();
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException($"Failed to save session '{sessionId}' before closing: {ex.Message}", ex);
                }
            }
        }

        // Then close
        return CloseSessionSync(sessionId);
    }

    private bool CloseSessionSync(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return false;
        }

        if (!_activeSessions.TryRemove(sessionId, out var batch))
        {
            return false;
        }

        // Remove file path metadata so it can be opened again
        if (_sessionFilePaths.TryRemove(sessionId, out var normalizedPath))
        {
            _activeFilePaths.TryRemove(normalizedPath, out _);
        }
        else
        {
            var filePathEntry = _activeFilePaths.FirstOrDefault(kvp => kvp.Value == sessionId);
            if (!filePathEntry.Equals(default(KeyValuePair<string, string>)))
            {
                _activeFilePaths.TryRemove(filePathEntry.Key, out _);
            }
        }

        // Clean up operation tracking data
        _activeOperationCounts.TryRemove(sessionId, out _);
        _showExcelFlags.TryRemove(sessionId, out _);

        try
        {
            batch.Dispose();
            return true;
        }
        catch
        {
            // Best effort - session is already removed from dictionary
            return true;
        }
    }

    /// <summary>
    /// Gets the number of active sessions.
    /// </summary>
    public int ActiveSessionCount => _activeSessions.Count;

    /// <summary>
    /// Checks if the Excel process for a session is still alive.
    /// </summary>
    /// <param name="sessionId">Session ID</param>
    /// <returns>True if session exists and Excel process is alive, false otherwise</returns>
    public bool IsSessionAlive(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId)) return false;
        if (!_activeSessions.TryGetValue(sessionId, out var batch)) return false;
        return batch.IsExcelProcessAlive();
    }

    /// <summary>
    /// Gets all active session IDs.
    /// </summary>
    public IEnumerable<string> ActiveSessionIds => _activeSessions.Keys.ToList();

    /// <summary>
    /// Returns a snapshot of active sessions with associated workbook paths.
    /// </summary>
    public IReadOnlyList<SessionDescriptor> GetActiveSessions()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        var snapshot = new List<SessionDescriptor>(_sessionFilePaths.Count);
        foreach (var kvp in _sessionFilePaths)
        {
            snapshot.Add(new SessionDescriptor(kvp.Key, kvp.Value));
        }

        return snapshot;
    }

    /// <summary>
    /// Attempts to get the workbook path associated with a session ID.
    /// </summary>
    public bool TryGetFilePath(string sessionId, [NotNullWhen(true)] out string? filePath)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        if (string.IsNullOrWhiteSpace(sessionId))
        {
            filePath = null;
            return false;
        }

        return _sessionFilePaths.TryGetValue(sessionId, out filePath);
    }

    /// <summary>
    /// Disposes all active sessions.
    /// </summary>
    /// <remarks>
    /// <para><b>CRITICAL:</b> Sessions are disposed SEQUENTIALLY to avoid COM threading issues.</para>
    /// <para>Excel COM objects must be disposed on their STA threads. Parallel disposal causes deadlocks.</para>
    /// </remarks>
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;

        // Close all active sessions SEQUENTIALLY to avoid COM threading issues
        // Excel COM objects must be disposed on their STA threads, parallel disposal causes deadlocks
        var sessions = _activeSessions.Values.ToList();
        _activeSessions.Clear();
        _activeFilePaths.Clear();
        _sessionFilePaths.Clear();

        foreach (var session in sessions)
        {
            try
            {
                // Dispose sequentially - ExcelBatch.Dispose() handles its own Excel cleanup
                // via ExcelShutdownService with proper timeouts and retry logic
                session.Dispose();
            }
            catch
            {
                // Best effort cleanup - continue with remaining sessions
            }
        }
    }
}

/// <summary>
/// Represents a snapshot of an active Excel session managed by <see cref="SessionManager"/>.
/// </summary>
/// <param name="SessionId">Public session identifier shared with clients.</param>
/// <param name="FilePath">Normalized workbook path associated with the session.</param>
public sealed record SessionDescriptor(string SessionId, string FilePath);

/// <summary>
/// Result of validating whether a session can be closed.
/// </summary>
/// <param name="SessionExists">Whether the session was found.</param>
/// <param name="IsExcelVisible">Whether Excel is visible (showExcel=true).</param>
/// <param name="ActiveOperationCount">Number of operations currently running.</param>
/// <param name="BlockingReason">Reason why close is blocked, or null if close is allowed.</param>
public sealed record CloseValidationResult(
    bool SessionExists,
    bool IsExcelVisible,
    int ActiveOperationCount,
    string? BlockingReason)
{
    /// <summary>
    /// Whether the session can be closed (no blocking conditions).
    /// </summary>
    public bool CanClose => SessionExists && ActiveOperationCount == 0;
}

