using System.Runtime.InteropServices;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal static class MacPathCanonicalizer
{
    public static string Normalize(string path)
    {
        var fullPath = Path.GetFullPath(path);
        if (!OperatingSystem.IsMacOS())
        {
            return fullPath;
        }

        var suffix = new Stack<string>();
        var existingPath = fullPath;
        while (!File.Exists(existingPath) && !Directory.Exists(existingPath))
        {
            var name = Path.GetFileName(existingPath);
            if (string.IsNullOrEmpty(name))
            {
                return fullPath;
            }

            suffix.Push(name);
            var parent = Path.GetDirectoryName(existingPath);
            if (string.IsNullOrEmpty(parent) || string.Equals(parent, existingPath, StringComparison.Ordinal))
            {
                return fullPath;
            }
            existingPath = parent;
        }

        var canonical = ResolveExistingPath(existingPath);
        while (suffix.TryPop(out var component))
        {
            canonical = Path.Combine(canonical, component);
        }
        return UseMacPresentationAlias(canonical);
    }

    private static string UseMacPresentationAlias(string path)
    {
        if (string.Equals(path, "/private/var", StringComparison.Ordinal))
        {
            return "/var";
        }
        if (path.StartsWith("/private/var/", StringComparison.Ordinal))
        {
            return path["/private".Length..];
        }
        if (string.Equals(path, "/private/tmp", StringComparison.Ordinal))
        {
            return "/tmp";
        }
        if (path.StartsWith("/private/tmp/", StringComparison.Ordinal))
        {
            return path["/private".Length..];
        }
        return path;
    }

    private static string ResolveExistingPath(string path)
    {
        var nativePath = Marshal.StringToCoTaskMemUTF8(path);
        IntPtr result;
        try
        {
            result = RealPath(nativePath, IntPtr.Zero);
        }
        finally
        {
            Marshal.FreeCoTaskMem(nativePath);
        }

        if (result == IntPtr.Zero)
        {
            throw new IOException(
                $"Could not resolve the canonical macOS path '{path}'.",
                new System.ComponentModel.Win32Exception(Marshal.GetLastPInvokeError()));
        }

        try
        {
            return Marshal.PtrToStringUTF8(result)
                ?? throw new IOException($"macOS returned an invalid canonical path for '{path}'.");
        }
        finally
        {
            Free(result);
        }
    }

    [DllImport("/usr/lib/libSystem.B.dylib", EntryPoint = "realpath", SetLastError = true)]
    private static extern IntPtr RealPath(IntPtr path, IntPtr resolvedPath);

    [DllImport("/usr/lib/libSystem.B.dylib", EntryPoint = "free")]
    private static extern void Free(IntPtr pointer);
}
