using System.Diagnostics;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

internal sealed class WorkbookSavepointStore : IDisposable
{
    internal const int MaxSavepointsPerSession = 8;
    internal const long MaxBytesPerSession = 1024L * 1024L * 1024L;
    internal const long MaxBytesPerProcess = 4L * 1024L * 1024L * 1024L;

    private readonly object _sync = new();
    private readonly Dictionary<string, Dictionary<string, SavepointEntry>> _entries =
        new(StringComparer.Ordinal);
    private readonly Dictionary<ReservationKey, long> _reservations = [];
    private readonly Dictionary<string, TransientEntry> _transientFiles =
        new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, long> _orphanedDirectories =
        new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, long> _orphanedFiles =
        new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _preservedRecoveryFiles =
        new(StringComparer.OrdinalIgnoreCase);
    private readonly string _instanceDirectory;
    private bool _disposed;

    internal WorkbookSavepointStore()
    {
        var root = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ExcelMcp",
            "savepoints");
        TryDeleteStaleInstanceDirectories(root);

        using var process = Process.GetCurrentProcess();
        var identity = $"{process.Id}-{process.StartTime.ToUniversalTime().ToFileTimeUtc()}";
        _instanceDirectory = Path.Combine(root, $"{identity}-{Guid.NewGuid():N}");
    }

    internal static WorkbookSavepointLimits Limits => new(
        MaxSavepointsPerSession,
        MaxBytesPerSession,
        MaxBytesPerProcess);

    internal static void ValidateName(string name)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        if (name.Length > 128 ||
            !IsAsciiLetterOrDigit(name[0]) ||
            name.Any(character =>
                !IsAsciiLetterOrDigit(character) &&
                character is not ('.' or '_' or '-')))
        {
            throw new ArgumentException(
                "Savepoint name must be 1-128 characters, start with an ASCII letter or digit, " +
                "and contain only ASCII letters, digits, '.', '_', or '-'.",
                nameof(name));
        }
    }

    internal string CreateSnapshotPath(string sessionId, string extension)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        var sessionDirectory = GetSessionDirectory(sessionId);
        Directory.CreateDirectory(sessionDirectory);
        return Path.Combine(sessionDirectory, $"{Guid.NewGuid():N}{extension}");
    }

    internal string CreateTransientPath(
        string sessionId,
        string extension,
        long estimatedSizeBytes)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        if (estimatedSizeBytes < 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(estimatedSizeBytes),
                "Estimated transient size cannot be negative.");
        }

        var sessionDirectory = GetSessionDirectory(sessionId);
        Directory.CreateDirectory(sessionDirectory);
        var path = Path.Combine(
            sessionDirectory,
            $".transient-{Guid.NewGuid():N}{extension}");
        lock (_sync)
        {
            var processBytes = _entries.Values
                                   .SelectMany(entries => entries.Values)
                                   .Sum(entry => entry.SizeBytes) +
                               _reservations.Values.Sum() +
                               _transientFiles.Values.Sum(entry => entry.SizeBytes) +
                               _orphanedDirectories.Values.Sum() +
                               _orphanedFiles.Values.Sum();
            EnsureWithinLimit(
                processBytes,
                estimatedSizeBytes,
                MaxBytesPerProcess,
                "Savepoint and rollback storage for this service process");
            _transientFiles.Add(
                path,
                new TransientEntry(sessionId, estimatedSizeBytes));
        }
        return path;
    }

    internal bool Contains(string sessionId, string name)
    {
        lock (_sync)
        {
            return _entries.TryGetValue(sessionId, out var sessionEntries) &&
                   sessionEntries.ContainsKey(name);
        }
    }

    internal void Reserve(string sessionId, string name, long estimatedSizeBytes)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        RetryOrphanedStorage();
        if (estimatedSizeBytes < 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(estimatedSizeBytes),
                "Estimated snapshot size cannot be negative.");
        }

        var key = new ReservationKey(sessionId, name);
        lock (_sync)
        {
            var sessionEntries = GetOrCreateSessionEntries(sessionId);
            if (sessionEntries.ContainsKey(name) || _reservations.ContainsKey(key))
            {
                throw new InvalidOperationException(
                    $"Savepoint '{name}' already exists for session '{sessionId}'.");
            }

            var sessionReservations = _reservations
                .Where(reservation => string.Equals(
                    reservation.Key.SessionId,
                    sessionId,
                    StringComparison.Ordinal))
                .ToArray();
            if (sessionEntries.Count + sessionReservations.Length >= MaxSavepointsPerSession)
            {
                throw new WorkbookSavepointStorageLimitException(
                    $"Session '{sessionId}' already has the maximum of {MaxSavepointsPerSession} savepoints.");
            }

            var sessionBytes = sessionEntries.Values.Sum(entry => entry.SizeBytes) +
                               sessionReservations.Sum(reservation => reservation.Value);
            EnsureWithinLimit(
                sessionBytes,
                estimatedSizeBytes,
                MaxBytesPerSession,
                $"Savepoint storage for session '{sessionId}'");

            var processBytes = _entries.Values
                                   .SelectMany(entries => entries.Values)
                                   .Sum(entry => entry.SizeBytes) +
                               _reservations.Values.Sum() +
                               _transientFiles.Values.Sum(entry => entry.SizeBytes) +
                               _orphanedDirectories.Values.Sum() +
                               _orphanedFiles.Values.Sum();
            EnsureWithinLimit(
                processBytes,
                estimatedSizeBytes,
                MaxBytesPerProcess,
                "Savepoint storage for this service process");

            _reservations.Add(key, estimatedSizeBytes);
        }
    }

    internal void CancelReservation(string sessionId, string name)
    {
        lock (_sync)
        {
            _reservations.Remove(new ReservationKey(sessionId, name));
        }
    }

    internal SavepointEntry Commit(
        string sessionId,
        string name,
        string workbookPath,
        string snapshotPath)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        EnsureOwnedPath(snapshotPath);
        var fileInfo = new FileInfo(snapshotPath);
        if (!fileInfo.Exists)
        {
            throw new IOException("Excel did not create the savepoint snapshot.");
        }

        lock (_sync)
        {
            var sessionEntries = GetOrCreateSessionEntries(sessionId);
            var reservationKey = new ReservationKey(sessionId, name);
            if (!_reservations.ContainsKey(reservationKey))
            {
                throw new InvalidOperationException(
                    $"Savepoint reservation '{name}' is no longer active for session '{sessionId}'.");
            }

            var sessionBytes = sessionEntries.Values.Sum(entry => entry.SizeBytes);
            EnsureWithinLimit(
                sessionBytes,
                fileInfo.Length,
                MaxBytesPerSession,
                $"Savepoint storage for session '{sessionId}'");

            var processBytes = _entries.Values
                .SelectMany(entries => entries.Values)
                .Sum(entry => entry.SizeBytes) +
                _reservations
                    .Where(reservation => reservation.Key != reservationKey)
                    .Sum(reservation => reservation.Value) +
                _transientFiles.Values.Sum(entry => entry.SizeBytes) +
                _orphanedDirectories.Values.Sum() +
                _orphanedFiles.Values.Sum();
            EnsureWithinLimit(
                processBytes,
                fileInfo.Length,
                MaxBytesPerProcess,
                "Savepoint storage for this service process");

            var entry = new SavepointEntry(
                name,
                DateTime.UtcNow,
                fileInfo.Length,
                Path.GetFullPath(workbookPath),
                snapshotPath);
            _reservations.Remove(reservationKey);
            sessionEntries.Add(name, entry);
            return entry;
        }
    }

    internal SavepointEntry GetRequired(string sessionId, string name)
    {
        lock (_sync)
        {
            if (_entries.TryGetValue(sessionId, out var sessionEntries) &&
                sessionEntries.TryGetValue(name, out var entry))
            {
                return entry;
            }
        }

        throw new KeyNotFoundException(
            $"Savepoint '{name}' was not found for session '{sessionId}'.");
    }

    internal IReadOnlyList<SavepointEntry> List(string sessionId)
    {
        lock (_sync)
        {
            return _entries.TryGetValue(sessionId, out var sessionEntries)
                ? sessionEntries.Values
                    .OrderBy(entry => entry.CreatedAtUtc)
                    .ThenBy(entry => entry.Name, StringComparer.Ordinal)
                    .ToArray()
                : [];
        }
    }

    internal bool Release(string sessionId, string name)
    {
        lock (_sync)
        {
            if (!_entries.TryGetValue(sessionId, out var sessionEntries) ||
                !sessionEntries.TryGetValue(name, out var entry))
            {
                return false;
            }

            DeleteOwnedFile(entry.SnapshotPath);
            sessionEntries.Remove(name);
            if (sessionEntries.Count == 0)
            {
                _entries.Remove(sessionId);
                TryDeleteDirectory(GetSessionDirectory(sessionId));
            }

            return true;
        }
    }

    internal bool ReleaseAll(string sessionId)
    {
        var sessionDirectory = GetSessionDirectory(sessionId);
        string[] preservedFiles;
        lock (_sync)
        {
            long orphanedBytes = 0;
            if (_entries.Remove(sessionId, out var sessionEntries))
            {
                orphanedBytes += sessionEntries.Values.Sum(entry => entry.SizeBytes);
            }

            foreach (var reservation in _reservations.Keys
                         .Where(reservation => string.Equals(
                             reservation.SessionId,
                             sessionId,
                             StringComparison.Ordinal))
                         .ToArray())
            {
                orphanedBytes += _reservations[reservation];
                _reservations.Remove(reservation);
            }

            foreach (var transientPath in _transientFiles
                         .Where(transient => string.Equals(
                             transient.Value.SessionId,
                             sessionId,
                             StringComparison.Ordinal))
                         .Select(transient => transient.Key)
                         .ToArray())
            {
                orphanedBytes += _transientFiles[transientPath].SizeBytes;
                _transientFiles.Remove(transientPath);
            }

            var sessionPrefix = sessionDirectory + Path.DirectorySeparatorChar;
            foreach (var orphanedFile in _orphanedFiles.Keys
                         .Where(path => path.StartsWith(
                             sessionPrefix,
                             StringComparison.OrdinalIgnoreCase))
                         .ToArray())
            {
                orphanedBytes += _orphanedFiles[orphanedFile];
                _orphanedFiles.Remove(orphanedFile);
            }

            _orphanedDirectories[sessionDirectory] = orphanedBytes;
            preservedFiles = _preservedRecoveryFiles
                .Where(path => path.StartsWith(
                    sessionPrefix,
                    StringComparison.OrdinalIgnoreCase))
                .ToArray();
        }

        if (preservedFiles.Length > 0)
        {
            var preserved = preservedFiles.ToHashSet(StringComparer.OrdinalIgnoreCase);
            var cleanupComplete = true;
            try
            {
                foreach (var file in Directory.EnumerateFiles(sessionDirectory))
                {
                    if (preserved.Contains(file))
                    {
                        continue;
                    }

                    try
                    {
                        File.Delete(file);
                    }
                    catch (IOException)
                    {
                        cleanupComplete = false;
                    }
                    catch (UnauthorizedAccessException)
                    {
                        cleanupComplete = false;
                    }
                }
            }
            catch (IOException)
            {
                cleanupComplete = false;
            }
            catch (UnauthorizedAccessException)
            {
                cleanupComplete = false;
            }

            lock (_sync)
            {
                _orphanedDirectories.Remove(sessionDirectory);
            }
            return cleanupComplete;
        }

        if (TryDeleteDirectory(sessionDirectory))
        {
            lock (_sync)
            {
                _orphanedDirectories.Remove(sessionDirectory);
            }
            return true;
        }

        lock (_sync)
        {
            _orphanedDirectories[sessionDirectory] =
                Math.Max(
                    _orphanedDirectories[sessionDirectory],
                    GetDirectorySize(sessionDirectory));
        }
        return false;
    }

    internal void DeleteTransient(string path, long minimumSizeBytes = 0)
    {
        try
        {
            DeleteOwnedFile(path);
            lock (_sync)
            {
                _transientFiles.Remove(path);
                _orphanedFiles.Remove(path);
            }
        }
        catch (IOException)
        {
            TrackOrphanedFile(path, minimumSizeBytes);
        }
        catch (UnauthorizedAccessException)
        {
            TrackOrphanedFile(path, minimumSizeBytes);
        }
    }

    internal void UpdateTransientSize(string path, long sizeBytes)
    {
        if (sizeBytes < 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(sizeBytes),
                "Transient size cannot be negative.");
        }

        lock (_sync)
        {
            if (!_transientFiles.TryGetValue(path, out var transient))
            {
                throw new InvalidOperationException(
                    "The rollback recovery file is no longer owned by this savepoint store.");
            }

            var processBytes = _entries.Values
                                   .SelectMany(entries => entries.Values)
                                   .Sum(entry => entry.SizeBytes) +
                               _reservations.Values.Sum() +
                               _transientFiles
                                   .Where(entry => !string.Equals(
                                       entry.Key,
                                       path,
                                       StringComparison.OrdinalIgnoreCase))
                                   .Sum(entry => entry.Value.SizeBytes) +
                               _orphanedDirectories.Values.Sum() +
                               _orphanedFiles.Values.Sum();
            EnsureWithinLimit(
                processBytes,
                sizeBytes,
                MaxBytesPerProcess,
                "Savepoint and rollback storage for this service process");
            _transientFiles[path] = transient with { SizeBytes = sizeBytes };
        }
    }

    internal string PromoteRecoveryFile(string recoveryPath, string workbookPath)
    {
        EnsureOwnedPath(recoveryPath);
        var workbookDirectory = Path.GetDirectoryName(workbookPath)
            ?? throw new InvalidOperationException("Workbook directory is unavailable.");
        var fileName = Path.GetFileNameWithoutExtension(workbookPath);
        var extension = Path.GetExtension(workbookPath);
        var recoveryFileName =
            $".{fileName}.excelmcp-recovery-{DateTime.UtcNow:yyyyMMddHHmmss}-{Guid.NewGuid():N}{extension}";

        Exception workbookDirectoryFailure;
        try
        {
            var workbookRecoveryPath = Path.Combine(
                workbookDirectory,
                recoveryFileName);
            File.Move(recoveryPath, workbookRecoveryPath);
            UntrackTransient(recoveryPath);
            return workbookRecoveryPath;
        }
        catch (IOException ex)
        {
            workbookDirectoryFailure = ex;
        }
        catch (UnauthorizedAccessException ex)
        {
            workbookDirectoryFailure = ex;
        }

        var fallbackDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ExcelMcp",
            "recovery");
        try
        {
            Directory.CreateDirectory(fallbackDirectory);
            var fallbackPath = Path.Combine(fallbackDirectory, recoveryFileName);
            File.Move(recoveryPath, fallbackPath);
            UntrackTransient(recoveryPath);
            return fallbackPath;
        }
        catch (IOException ex)
        {
            throw new IOException(
                "The emergency recovery copy could not be moved to a durable location.",
                new AggregateException(workbookDirectoryFailure, ex));
        }
        catch (UnauthorizedAccessException ex)
        {
            throw new UnauthorizedAccessException(
                "The emergency recovery copy could not be moved to a durable location.",
                new AggregateException(workbookDirectoryFailure, ex));
        }
    }

    internal string PreserveRecoveryFileInPlace(string recoveryPath)
    {
        EnsureOwnedPath(recoveryPath);
        if (!File.Exists(recoveryPath))
        {
            throw new FileNotFoundException(
                "The emergency recovery copy no longer exists.",
                recoveryPath);
        }

        lock (_sync)
        {
            _transientFiles.Remove(recoveryPath);
            _orphanedFiles.Remove(recoveryPath);
            _preservedRecoveryFiles.Add(recoveryPath);
        }

        File.WriteAllText(
            Path.Combine(_instanceDirectory, ".preserve-recovery"),
            "This directory contains a caller-owned recovery workbook.");
        return recoveryPath;
    }

    private void UntrackTransient(string path)
    {
        lock (_sync)
        {
            _transientFiles.Remove(path);
            _orphanedFiles.Remove(path);
        }
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        lock (_sync)
        {
            _entries.Clear();
            _reservations.Clear();
            _transientFiles.Clear();
            _orphanedDirectories.Clear();
            _orphanedFiles.Clear();
            if (_preservedRecoveryFiles.Count > 0)
            {
                return;
            }
        }

        TryDeleteDirectory(_instanceDirectory);
    }

    private Dictionary<string, SavepointEntry> GetOrCreateSessionEntries(string sessionId)
    {
        if (!_entries.TryGetValue(sessionId, out var sessionEntries))
        {
            sessionEntries = new Dictionary<string, SavepointEntry>(StringComparer.Ordinal);
            _entries.Add(sessionId, sessionEntries);
        }

        return sessionEntries;
    }

    private static void EnsureWithinLimit(
        long existingBytes,
        long candidateBytes,
        long limitBytes,
        string description)
    {
        if (candidateBytes > limitBytes || existingBytes > limitBytes - candidateBytes)
        {
            throw new WorkbookSavepointStorageLimitException(
                $"{description} would exceed the {limitBytes} byte limit.");
        }
    }

    private void RetryOrphanedStorage()
    {
        string[] directories;
        string[] files;
        lock (_sync)
        {
            directories = _orphanedDirectories.Keys.ToArray();
            files = _orphanedFiles.Keys.ToArray();
        }

        foreach (var directory in directories)
        {
            if (!TryDeleteDirectory(directory))
            {
                lock (_sync)
                {
                    _orphanedDirectories[directory] =
                        Math.Max(
                            _orphanedDirectories[directory],
                            GetDirectorySize(directory));
                }
                continue;
            }

            lock (_sync)
            {
                _orphanedDirectories.Remove(directory);
            }
        }

        foreach (var file in files)
        {
            try
            {
                DeleteOwnedFile(file);
                lock (_sync)
                {
                    _orphanedFiles.Remove(file);
                }
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
        }
    }

    private void TrackOrphanedFile(string path, long minimumSizeBytes)
    {
        EnsureOwnedPath(path);
        long sizeBytes = minimumSizeBytes;
        lock (_sync)
        {
            if (_transientFiles.TryGetValue(path, out var transient))
            {
                sizeBytes = Math.Max(sizeBytes, transient.SizeBytes);
            }
        }

        try
        {
            if (File.Exists(path))
            {
                sizeBytes = Math.Max(sizeBytes, new FileInfo(path).Length);
            }
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }

        lock (_sync)
        {
            _transientFiles.Remove(path);
            _orphanedFiles[path] = sizeBytes;
        }
    }

    private static long GetDirectorySize(string directory)
    {
        try
        {
            return Directory.Exists(directory)
                ? Directory.EnumerateFiles(
                        directory,
                        "*",
                        SearchOption.AllDirectories)
                    .Sum(path => new FileInfo(path).Length)
                : 0;
        }
        catch (IOException)
        {
            return 0;
        }
        catch (UnauthorizedAccessException)
        {
            return 0;
        }
    }

    private string GetSessionDirectory(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId) ||
            sessionId.Any(character => !IsAsciiLetterOrDigit(character)))
        {
            throw new ArgumentException("Session ID contains unsupported characters.", nameof(sessionId));
        }

        return Path.Combine(_instanceDirectory, sessionId);
    }

    private void DeleteOwnedFile(string path)
    {
        EnsureOwnedPath(path);
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }

    private void EnsureOwnedPath(string path)
    {
        var fullPath = Path.GetFullPath(path);
        var root = Path.GetFullPath(_instanceDirectory) + Path.DirectorySeparatorChar;
        if (!fullPath.StartsWith(root, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("Refusing to access a file outside the savepoint store.");
        }
    }

    private static bool IsAsciiLetterOrDigit(char character) =>
        character is >= 'A' and <= 'Z' or >= 'a' and <= 'z' or >= '0' and <= '9';

    private static void TryDeleteStaleInstanceDirectories(string root)
    {
        if (!Directory.Exists(root))
        {
            return;
        }

        try
        {
            foreach (var directory in Directory.EnumerateDirectories(root))
            {
                if (File.Exists(Path.Combine(directory, ".preserve-recovery")))
                {
                    continue;
                }

                var name = Path.GetFileName(directory);
                var parts = name.Split('-', 3);
                if (parts.Length != 3 ||
                    !int.TryParse(parts[0], out var processId) ||
                    !long.TryParse(parts[1], out var startTimeFileTime))
                {
                    continue;
                }

                if (!IsExactProcessAlive(processId, startTimeFileTime))
                {
                    TryDeleteDirectory(directory);
                }
            }
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    private static bool IsExactProcessAlive(int processId, long startTimeFileTime)
    {
        try
        {
            using var process = Process.GetProcessById(processId);
            return !process.HasExited &&
                   process.StartTime.ToUniversalTime().ToFileTimeUtc() == startTimeFileTime;
        }
        catch (ArgumentException)
        {
            return false;
        }
        catch (InvalidOperationException)
        {
            return false;
        }
        catch (System.ComponentModel.Win32Exception)
        {
            return true;
        }
    }

    private static bool TryDeleteDirectory(string path)
    {
        try
        {
            if (Directory.Exists(path))
            {
                Directory.Delete(path, recursive: true);
            }
            return true;
        }
        catch (IOException)
        {
            return false;
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
    }

    internal sealed record SavepointEntry(
        string Name,
        DateTime CreatedAtUtc,
        long SizeBytes,
        string WorkbookPath,
        string SnapshotPath)
    {
        internal WorkbookSavepointDescriptor ToDescriptor() =>
            new(Name, CreatedAtUtc, SizeBytes, WorkbookPath);
    }

    private readonly record struct ReservationKey(string SessionId, string Name);

    private readonly record struct TransientEntry(string SessionId, long SizeBytes);
}
