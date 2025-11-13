using System.Collections.Concurrent;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Core;

/// <summary>
/// Manages active Excel sessions for the MCP server and CLI.
/// Maps user-facing sessionId to internal IExcelBatch instances.
/// </summary>
public sealed class SessionManager : IAsyncDisposable
{
    private readonly ConcurrentDictionary<string, IExcelBatch> _activeSessions = new();
    private bool _disposed;

    /// <summary>
    /// Creates a new session for the specified Excel file.
    /// </summary>
    /// <param name="filePath">Path to the Excel file to open</param>
    /// <returns>Unique session ID for this session</returns>
    /// <exception cref="FileNotFoundException">File does not exist</exception>
    /// <exception cref="InvalidOperationException">Failed to create session</exception>
    public async Task<string> CreateSessionAsync(string filePath)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException($"Excel file not found: {filePath}", filePath);
        }

        // Generate unique session ID
        string sessionId = Guid.NewGuid().ToString("N");

        try
        {
            // Create batch session using Core API
            var batch = await ExcelSession.BeginBatchAsync(filePath);

            // Store in active sessions
            if (!_activeSessions.TryAdd(sessionId, batch))
            {
                await batch.DisposeAsync();
                throw new InvalidOperationException($"Session ID collision: {sessionId}");
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
    /// <param name="sessionId">Session ID returned from CreateSessionAsync</param>
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
    public async Task<bool> SaveSessionAsync(string sessionId)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        var batch = GetSession(sessionId);
        if (batch == null)
        {
            return false;
        }

        try
        {
            await batch.SaveAsync();
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
    public async Task<bool> CloseSessionAsync(string sessionId)
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

        try
        {
            await batch.DisposeAsync();
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
    public async ValueTask DisposeAsync()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;

        // Close all active sessions
        var sessions = _activeSessions.Values.ToList();
        _activeSessions.Clear();

        foreach (var batch in sessions)
        {
            try
            {
                await batch.DisposeAsync();
            }
            catch
            {
                // Best effort cleanup
            }
        }
    }
}
