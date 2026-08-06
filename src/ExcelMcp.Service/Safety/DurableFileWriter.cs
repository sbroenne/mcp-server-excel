using System.Text;

namespace Sbroenne.ExcelMcp.Service.Safety;

/// <summary>
/// Publishes safety evidence only after its bytes have been flushed to stable storage.
/// Temporary and destination files always share a directory so publication is atomic.
/// </summary>
internal static class DurableFileWriter
{
    private static readonly UTF8Encoding Utf8WithoutBom = new(
        encoderShouldEmitUTF8Identifier: false,
        throwOnInvalidBytes: true);

    public static void WriteUtf8Atomically(string destinationPath, string content)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(destinationPath);
        ArgumentNullException.ThrowIfNull(content);

        var directory = Path.GetDirectoryName(destinationPath) ??
            throw new ArgumentException("A durable destination must have a parent directory.", nameof(destinationPath));
        var temporaryPath = Path.Combine(
            directory,
            $".{Path.GetFileName(destinationPath)}.{Guid.NewGuid():N}.tmp");

        try
        {
            var bytes = Utf8WithoutBom.GetBytes(content);
            using (var stream = new FileStream(
                       temporaryPath,
                       FileMode.CreateNew,
                       FileAccess.Write,
                       FileShare.None,
                       bufferSize: 16 * 1024,
                       FileOptions.WriteThrough | FileOptions.SequentialScan))
            {
                stream.Write(bytes);
                stream.Flush(flushToDisk: true);
            }

            PublishFlushedFileAtomically(temporaryPath, destinationPath);
        }
        finally
        {
            TryDeleteTemporaryFile(temporaryPath);
        }
    }

    public static void FlushExistingFile(string path)
    {
        using var stream = new FileStream(
            path,
            FileMode.Open,
            FileAccess.ReadWrite,
            FileShare.Read,
            bufferSize: 16 * 1024,
            FileOptions.WriteThrough | FileOptions.SequentialScan);
        stream.Flush(flushToDisk: true);
    }

    public static void PublishFlushedFileAtomically(string temporaryPath, string destinationPath)
    {
        var temporaryDirectory = Path.GetFullPath(Path.GetDirectoryName(temporaryPath)!);
        var destinationDirectory = Path.GetFullPath(Path.GetDirectoryName(destinationPath)!);
        if (!string.Equals(temporaryDirectory, destinationDirectory, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("Atomic publication requires temporary and destination files in the same directory.");
        }

        if (File.Exists(destinationPath))
        {
            File.Replace(temporaryPath, destinationPath, destinationBackupFileName: null, ignoreMetadataErrors: true);
        }
        else
        {
            File.Move(temporaryPath, destinationPath);
        }
    }

    private static void TryDeleteTemporaryFile(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (IOException)
        {
            // Preserve the original write/publication exception.
        }
        catch (UnauthorizedAccessException)
        {
            // Preserve the original write/publication exception.
        }
    }
}
