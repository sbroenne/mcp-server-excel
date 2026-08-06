// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

namespace Sbroenne.ExcelMcp.Service.Safety;

/// <summary>
/// Keeps durable safety evidence on a local path without traversing reparse points.
/// </summary>
internal static class SafetyStatePathPolicy
{
    public static string PrepareRoot(string root) => PrepareRoot(
        root,
        EnsureNoExistingReparsePoint,
        path => Directory.CreateDirectory(path));

    internal static string PrepareRoot(
        string root,
        Action<string> validateExistingAncestors,
        Action<string> createDirectory)
    {
        if (string.IsNullOrWhiteSpace(root))
        {
            throw new ArgumentException("A safety-state root is required.", nameof(root));
        }

        ArgumentNullException.ThrowIfNull(validateExistingAncestors);
        ArgumentNullException.ThrowIfNull(createDirectory);

        if (IsNetworkPath(root))
        {
            throw new InvalidOperationException("The safety-state root must be on a local filesystem.");
        }

        var fullPath = Path.GetFullPath(root);
        if (IsNetworkPath(fullPath))
        {
            throw new InvalidOperationException("The safety-state root must be on a local filesystem.");
        }

        // Validate before creating anything. Otherwise an absent final directory below
        // an existing junction could be created through that junction before rejection.
        validateExistingAncestors(fullPath);
        createDirectory(fullPath);
        // Validate again to fail closed if the filesystem changed during creation.
        validateExistingAncestors(fullPath);
        return fullPath;
    }

    public static void EnsureSafePath(string root, string path)
    {
        var fullRoot = Path.GetFullPath(root);
        var fullPath = Path.GetFullPath(path);
        var relative = Path.GetRelativePath(fullRoot, fullPath);
        if (relative == ".." ||
            relative.StartsWith($"..{Path.DirectorySeparatorChar}", StringComparison.Ordinal) ||
            Path.IsPathRooted(relative))
        {
            throw new InvalidOperationException("Safety-state path escaped the configured local root.");
        }

        EnsureNoExistingReparsePoint(fullRoot);

        var current = fullPath;
        while (!string.IsNullOrWhiteSpace(current))
        {
            if (TryGetExistingAttributes(current, out var attributes))
            {
                if (IsReparsePoint(attributes))
                {
                    throw new InvalidOperationException("Safety-state paths cannot contain symbolic links, junctions, or other reparse points.");
                }
            }

            if (string.Equals(
                    Path.TrimEndingDirectorySeparator(current),
                    Path.TrimEndingDirectorySeparator(fullRoot),
                    StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            current = Path.GetDirectoryName(current);
        }

        throw new InvalidOperationException("Safety-state path could not be validated beneath the configured local root.");
    }

    internal static bool IsReparsePoint(FileAttributes attributes) =>
        (attributes & FileAttributes.ReparsePoint) != 0;

    internal static bool IsNetworkDriveType(DriveType driveType) =>
        driveType == DriveType.Network;

    /// <summary>
    /// Returns false only for a path that is definitely absent. Access and I/O failures
    /// propagate so callers fail closed instead of mistaking an inaccessible reparse
    /// point for a missing directory.
    /// </summary>
    internal static bool TryGetExistingAttributes(string path, out FileAttributes attributes)
    {
        try
        {
            attributes = File.GetAttributes(path);
            return true;
        }
        catch (FileNotFoundException)
        {
            attributes = default;
            return false;
        }
        catch (DirectoryNotFoundException)
        {
            attributes = default;
            return false;
        }
    }

    private static void EnsureNoExistingReparsePoint(string path)
    {
        string? current = path;
        while (!string.IsNullOrWhiteSpace(current))
        {
            if (TryGetExistingAttributes(current, out var attributes))
            {
                if (IsReparsePoint(attributes))
                {
                    throw new InvalidOperationException("Safety-state paths cannot contain symbolic links, junctions, or other reparse points.");
                }
            }

            var parent = Path.GetDirectoryName(Path.TrimEndingDirectorySeparator(current));
            if (string.IsNullOrWhiteSpace(parent) || string.Equals(parent, current, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            current = parent;
        }
    }

    private static bool IsNetworkPath(string path) =>
        path.StartsWith("\\\\", StringComparison.Ordinal) ||
        path.StartsWith("//", StringComparison.Ordinal) ||
        IsMappedNetworkDrive(path);

    private static bool IsMappedNetworkDrive(string path)
    {
        var root = Path.GetPathRoot(path);
        if (string.IsNullOrWhiteSpace(root))
        {
            return false;
        }

        try
        {
            return IsNetworkDriveType(new DriveInfo(root).DriveType);
        }
        catch (ArgumentException)
        {
            return true;
        }
        catch (IOException)
        {
            return true;
        }
        catch (UnauthorizedAccessException)
        {
            return true;
        }
    }
}
