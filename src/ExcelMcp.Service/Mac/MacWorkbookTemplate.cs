using System.Reflection;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal static class MacWorkbookTemplate
{
    private const string ResourceName = "Sbroenne.ExcelMcp.Service.Mac.Blank.xlsx";

    public static void Copy(string filePath, bool macroEnabled)
    {
        if (macroEnabled)
        {
            throw new PlatformNotSupportedException(
                "Creating macro-enabled workbooks on macOS requires an opaque " +
                "Excel-authored .xlsm template, which is not currently bundled.");
        }

        var directory = Path.GetDirectoryName(filePath)
            ?? throw new ArgumentException(
                "Workbook path has no parent directory.",
                nameof(filePath));
        var temporaryPath = Path.Combine(
            directory,
            $".excelmcp-template-{Guid.NewGuid():N}.tmp");
        try
        {
            using var source = Assembly.GetExecutingAssembly()
                .GetManifestResourceStream(ResourceName)
                ?? throw new InvalidOperationException(
                    $"Embedded workbook template '{ResourceName}' is missing.");
            using (var destination = new FileStream(
                       temporaryPath,
                       FileMode.CreateNew,
                       FileAccess.Write,
                       FileShare.None))
            {
                source.CopyTo(destination);
                destination.Flush(flushToDisk: true);
            }

            File.Move(temporaryPath, filePath);
        }
        finally
        {
            File.Delete(temporaryPath);
        }
    }
}
