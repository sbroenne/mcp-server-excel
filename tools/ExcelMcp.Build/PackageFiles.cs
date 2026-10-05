using System.Diagnostics;

namespace Sbroenne.ExcelMcp.Build;

public static class PackageFiles
{
    public static void AssertOutput(string path, string root, IEnumerable<string?>? inputs = null)
    {
        var output = Path.TrimEndingDirectorySeparator(Path.GetFullPath(path, root));
        root = Path.TrimEndingDirectorySeparator(Path.GetFullPath(root));
        var allowed = new[] { Path.Combine(root, "artifacts"), Path.Combine(root, "plugins"), Path.Combine(root, "mcpb", "artifacts") };
        if (output == Path.TrimEndingDirectorySeparator(Path.GetPathRoot(output)!) ||
            Contains(output, root) || Contains(root, output) && !allowed.Any(directory => Contains(directory, output)))
        {
            throw new ArgumentException($"Unsafe package output directory: {output}");
        }
        foreach (var input in inputs ?? [])
        {
            if (string.IsNullOrWhiteSpace(input)) { continue; }
            var full = Path.GetFullPath(input, root);
            if (Contains(output, full) || Contains(full, output))
            {
                throw new ArgumentException($"Package output overlaps a prepared input: {output}");
            }
        }
        AssertNoLinks(output);
    }

    public static void AssertNoLinks(string path)
    {
        for (var current = Path.GetFullPath(path); current is not null; current = Path.GetDirectoryName(current))
        {
            if (Exists(current) && File.GetAttributes(current).HasFlag(FileAttributes.ReparsePoint))
            {
                throw new ArgumentException($"Package output must not traverse a link: {current}");
            }
        }
    }

    public static void AssertArchitecture(string path, string architecture)
    {
        var expected = architecture switch
        {
            "x64" => 0x8664,
            "arm64" => 0xaa64,
            _ => throw new ArgumentException($"Unknown runtime architecture: {architecture}.")
        };
        using var stream = File.OpenRead(path);
        using var reader = new BinaryReader(stream);
        if (stream.Length < 64 || reader.ReadUInt16() != 0x5a4d)
        {
            throw new InvalidOperationException($"Runtime is not a Windows executable: {path}");
        }
        stream.Position = 0x3c;
        var offset = reader.ReadUInt32();
        if (offset < 64 || offset > stream.Length - 6)
        {
            throw new InvalidOperationException($"Runtime has an invalid executable header: {path}");
        }
        stream.Position = offset;
        if (reader.ReadUInt32() != 0x00004550)
        {
            throw new InvalidOperationException($"Runtime has an invalid PE signature: {path}");
        }
        var machine = reader.ReadUInt16();
        if (machine != expected)
        {
            throw new InvalidOperationException($"Runtime machine type 0x{machine:x4} does not match {architecture}: {path}");
        }
    }

    public static void Copy(string source, string destination)
    {
        if (File.GetAttributes(source).HasFlag(FileAttributes.ReparsePoint))
        {
            throw new IOException($"Package source must not be a link: {source}");
        }
        if (!Directory.Exists(source))
        {
            Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
            File.Copy(source, destination, overwrite: true);
            return;
        }
        Directory.CreateDirectory(destination);
        foreach (var entry in Directory.EnumerateFileSystemEntries(source))
        {
            Copy(entry, Path.Combine(destination, Path.GetFileName(entry)));
        }
    }

    public static void Install(string source, string destination,
        Action<string, string>? move = null, Action<string>? delete = null, Action<string>? warning = null)
    {
        AssertNoLinks(destination);
        move ??= Move;
        delete ??= Delete;
        warning ??= Console.Error.WriteLine;
        var temporary = $"{destination}.{Guid.NewGuid():N}.tmp";
        var backup = $"{temporary}.bak";
        try
        {
            Copy(source, temporary);
            if (Exists(destination)) { move(destination, backup); }
            try { move(temporary, destination); }
            catch (Exception primary) when (primary is IOException or UnauthorizedAccessException)
            {
                if (!Exists(backup)) { throw; }
                try { move(backup, destination); }
                catch (Exception recovery) when (recovery is IOException or UnauthorizedAccessException)
                {
                    throw new AggregateException(
                        $"Package installation failed: {primary.Message}. Recovery also failed: {recovery.Message}. " +
                        $"Destination is {(Exists(destination) ? "present" : "absent")} at '{destination}'. " +
                        $"Recovery backup {(Exists(backup) ? "retained" : "is absent")} at '{backup}'.", primary, recovery);
                }
                throw;
            }
            if (Exists(backup))
            {
                try { delete(backup); }
                catch (Exception cleanup) when (cleanup is IOException or UnauthorizedAccessException)
                {
                    warning($"Package output installed successfully at '{destination}', but backup cleanup failed: {cleanup.Message}. Backup retained at '{backup}'.");
                }
            }
        }
        finally
        {
            if (Exists(temporary))
            {
                try { delete(temporary); }
                catch (Exception cleanup) when (cleanup is IOException or UnauthorizedAccessException)
                {
                    warning($"Temporary package output cleanup failed: {cleanup.Message}. Temporary output retained at '{temporary}'.");
                }
            }
        }
    }

    public static void Delete(string path)
    {
        AssertNoLinks(path);
        if (Directory.Exists(path)) { Directory.Delete(path, recursive: true); }
        else if (File.Exists(path)) { File.Delete(path); }
    }

    public static void RemoveStaging(string path, TimeSpan? timeout = null, TimeSpan? retryInterval = null,
        Action<string>? remove = null)
    {
        var limit = timeout ?? TimeSpan.FromMinutes(2);
        var interval = retryInterval ?? TimeSpan.FromMilliseconds(500);
        if (limit < TimeSpan.Zero || interval < TimeSpan.Zero) { throw new ArgumentOutOfRangeException(nameof(timeout)); }
        remove ??= Delete;
        var clock = Stopwatch.StartNew();
        Exception? last = null;
        var attempts = 0;
        while (Directory.Exists(path))
        {
            attempts++;
            try { last = null; remove(path); }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { last = exception; }
            if (!Directory.Exists(path)) { return; }
            var remaining = limit - clock.Elapsed;
            if (remaining <= TimeSpan.Zero)
            {
                throw new IOException(
                    $"Failed to remove MCPB staging directory '{path}' after {attempts} {(attempts == 1 ? "attempt" : "attempts")} within {Math.Round(limit.TotalMilliseconds)} ms. " +
                    $"The verified bundle was preserved, but stale staging remains. Last error: {last?.Message ?? "the directory still exists"}", last);
            }
            if (interval > TimeSpan.Zero) { Thread.Sleep(interval < remaining ? interval : remaining); }
        }
    }

    public static bool Contains(string directory, string path) =>
        Path.TrimEndingDirectorySeparator(directory).Equals(Path.TrimEndingDirectorySeparator(path), StringComparison.OrdinalIgnoreCase) ||
        path.StartsWith(Path.TrimEndingDirectorySeparator(directory) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);

    private static bool Exists(string path) => File.Exists(path) || Directory.Exists(path);
    private static void Move(string source, string destination)
    {
        if (Directory.Exists(source)) { Directory.Move(source, destination); }
        else { File.Move(source, destination); }
    }
}
