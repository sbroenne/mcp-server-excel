using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

[SupportedOSPlatform("macos")]
internal static class MacDaemonConsole
{
    public static TextWriter Redirect(string pipeName)
    {
        var logPath = Path.ChangeExtension(DaemonProcessTracker.GetTrackingFilePath(pipeName), ".log");
        Directory.CreateDirectory(Path.GetDirectoryName(logPath)!);
        using var input = File.OpenRead("/dev/null");
        var log = new FileStream(logPath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
        // Replace the inherited descriptors, not just Console writers: retaining the
        // caller's pipe would prevent its JSON reader from ever observing EOF.
        try
        {
            Duplicate(input.SafeFileHandle.DangerousGetHandle().ToInt32(), 0);
            Duplicate(log.SafeFileHandle.DangerousGetHandle().ToInt32(), 1);
            Duplicate(log.SafeFileHandle.DangerousGetHandle().ToInt32(), 2);
            var writer = TextWriter.Synchronized(new StreamWriter(log) { AutoFlush = true });
            Console.SetOut(writer);
            Console.SetError(writer);
            return writer;
        }
        catch
        {
            log.Dispose();
            throw;
        }
    }

    private static void Duplicate(int source, int destination)
    {
        if (dup2(source, destination) < 0)
        {
            throw new Win32Exception(Marshal.GetLastPInvokeError(), "Could not redirect macOS daemon standard streams.");
        }
    }

    [DllImport("/usr/lib/libSystem.B.dylib", SetLastError = true)]
    private static extern int dup2(int source, int destination);
}
