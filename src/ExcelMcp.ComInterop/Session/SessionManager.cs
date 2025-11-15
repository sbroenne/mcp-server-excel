using System.Collections.Concurrent;

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
    private bool _disposed;

    /// <summary>
    /// Creates a new session for the specified Excel file.
    /// </summary>
    /// <param name="filePath">Path to the Excel file to open</param>
    /// <returns>Unique session ID for this session</returns>
    /// <exception cref="FileNotFoundException">File does not exist</exception>
    /// <exception cref="InvalidOperationException">Failed to create session or file already open in another session</exception>
    /// <remarks>
    /// <para><b>Resource Impact:</b> Creates a new Excel.Application process (~50-100MB+ memory).</para>
    /// <para><b>Same-file prevention:</b> Throws if file is already open in another session.</para>
    /// <para><b>Concurrency:</b> You can create multiple sessions for DIFFERENT files. Operations within each session execute serially.</para>
    /// </remarks>
    public string CreateSession(string filePath)
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

        try
        {
            // Create batch session using Core API
            var batch = ExcelSession.BeginBatch(filePath);

            // Store in active sessions
            if (!_activeSessions.TryAdd(sessionId, batch))
            {
                batch.Dispose();
                throw new InvalidOperationException($"Session ID collision: {sessionId}");
            }

            // Track the file path
            if (!_activeFilePaths.TryAdd(normalizedPath, sessionId))
            {
                // Cleanup if file path tracking fails
                _activeSessions.TryRemove(sessionId, out _);
                batch.Dispose();
                throw new InvalidOperationException($"Failed to track file path for session: {sessionId}");
            }

            return sessionId;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to create session for '{filePath}': {ex.Message}", ex);
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
    /// Saves changes for the specified session.
    /// </summary>
    /// <param name="sessionId">Session ID</param>
    /// <returns>True if session was found and saved, false if session not found</returns>
    /// <exception cref="InvalidOperationException">Save operation failed</exception>
    public bool SaveSession(string sessionId)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        var batch = GetSession(sessionId);
        if (batch == null)
        {
            return false;
        }

        try
        {
            batch.Save();
            return true;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to save session '{sessionId}': {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Closes the specified session without saving changes.
    /// </summary>
    /// <param name="sessionId">Session ID</param>
    /// <returns>True if session was found and closed, false if session not found</returns>
    public bool CloseSession(string sessionId)
    {
        return CloseSessionSync(sessionId);
    }

    private bool CloseSessionSync(string sessionId)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return false;
        }

        if (!_activeSessions.TryRemove(sessionId, out var batch))
        {
            return false;
        }

        // Remove file path from tracking so it can be opened again
        var filePathEntry = _activeFilePaths.FirstOrDefault(kvp => kvp.Value == sessionId);
        if (!filePathEntry.Equals(default(KeyValuePair<string, string>)))
        {
            _activeFilePaths.TryRemove(filePathEntry.Key, out _);
        }

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
    /// Gets all active session IDs.
    /// </summary>
    public IEnumerable<string> ActiveSessionIds => _activeSessions.Keys.ToList();

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

        for (int i = 0; i < sessions.Count; i++)
        {
            try
            {
                // Dispose synchronously - Excel COM deadlock handled in Dispose() itself
                sessions[i].Dispose();

                // CRITICAL: Wait for Excel process to actually terminate before disposing next session
                // Excel COM has known synchronization issues when multiple instances are disposed rapidly
                // Without this delay, the second disposal can deadlock waiting for the first to complete
                if (i < sessions.Count - 1) // Don't delay after the last one
                {
                    // Wait up to 5 seconds for any EXCEL processes to terminate
                    var startWait = DateTime.UtcNow;
                    var maxWait = TimeSpan.FromSeconds(5);

                    while (DateTime.UtcNow - startWait < maxWait)
                    {
                        var excelProcesses = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                        if (excelProcesses.Length == 0)
                        {
                            // All Excel processes terminated, safe to proceed
                            break;
                        }

                        // Still have Excel processes, wait a bit
                        Thread.Sleep(200);

                        foreach (var p in excelProcesses)
                        {
                            p.Dispose();
                        }
                    }
                }
            }
            catch
            {
                // Best effort cleanup
            }
        }
    }
}

