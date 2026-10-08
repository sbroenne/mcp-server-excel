using System.Diagnostics;
using System.Text.Json.Nodes;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal static class MacNativeWorkbook
{
    internal static bool IsOpen(string filePath, TimeSpan timeout)
    {
        var target = MacPathCanonicalizer.Normalize(filePath);
        return ReadNames(fullPaths: true, timeout).Any(path => MatchesPath(path, target));
    }

    internal static void Attach(string filePath, bool show, TimeSpan timeout)
    {
        var started = Stopwatch.GetTimestamp();
        while (true)
        {
            using var workbook = Find(filePath, MacAppleEvents.Remaining(timeout, started));
            if (workbook is not null)
            {
                using var first = MacAppleEvents.Create(MacAppleEvents.Code("long"), BitConverter.GetBytes(1));
                using var window = MacAppleEvents.Object(MacExcelDictionary.WindowClass, workbook, MacAppleEvents.Code("indx"), first);
                using var visible = MacAppleEvents.Property(window, MacExcelDictionary.Visible);
                using var value = MacAppleEvents.Create(MacAppleEvents.Code("bool"), [show ? (byte)1 : (byte)0]);
                using var appleEvent = MacAppleEvents.Event(MacAppleEvents.Code("core"), MacAppleEvents.Code("setd"));
                MacAppleEvents.Put(appleEvent, MacAppleEvents.Code("----"), visible);
                MacAppleEvents.Put(appleEvent, MacAppleEvents.Code("data"), value);
                MacAppleEvents.SendCommand(appleEvent, MacAppleEvents.Remaining(timeout, started));
                return;
            }
            if (Stopwatch.GetElapsedTime(started) >= TimeSpan.FromSeconds(10))
            {
                throw new InvalidOperationException("LaunchServices did not open the requested workbook within ten seconds.");
            }
            Thread.Sleep(TimeSpan.FromMilliseconds(Math.Min(100, MacAppleEvents.Remaining(timeout, started).TotalMilliseconds)));
        }
    }

    internal static void Close(string filePath, bool save, TimeSpan timeout)
    {
        var started = Stopwatch.GetTimestamp();
        using var workbook = Find(filePath, timeout)
            ?? throw new InvalidOperationException("Workbook is not open in this ExcelMcp session.");
        using var saving = MacAppleEvents.Create(MacAppleEvents.Code("enum"),
            BitConverter.GetBytes(MacAppleEvents.Code(save ? "yes " : "no  ")));
        using var appleEvent = MacAppleEvents.Event(MacExcelDictionary.CloseClass, MacExcelDictionary.CloseId);
        MacAppleEvents.Put(appleEvent, MacAppleEvents.Code("----"), workbook);
        MacAppleEvents.Put(appleEvent, MacAppleEvents.Code("savo"), saving);
        MacAppleEvents.SendCommand(appleEvent, MacAppleEvents.Remaining(timeout, started));
    }

    private static MacAppleEvents.Descriptor? Find(string filePath, TimeSpan timeout)
    {
        var target = MacPathCanonicalizer.Normalize(filePath);
        var started = Stopwatch.GetTimestamp();
        if (!ReadNames(fullPaths: true, timeout).Any(path => MatchesPath(path, target)))
        {
            return null;
        }
        using var application = MacAppleEvents.Create(MacAppleEvents.Code("null"), []);
        using var name = MacAppleEvents.Text(Path.GetFileName(target));
        var workbook = MacAppleEvents.Object(MacExcelDictionary.WorkbookClass, application, MacAppleEvents.Code("name"), name);
        try
        {
            using var fullName = MacAppleEvents.Property(workbook, MacExcelDictionary.FullName);
            using var appleEvent = MacAppleEvents.Event(MacAppleEvents.Code("core"), MacAppleEvents.Code("getd"));
            MacAppleEvents.Put(appleEvent, MacAppleEvents.Code("----"), fullName);
            var actualPath = MacAppleEvents.Send(appleEvent, MacAppleEvents.Remaining(timeout, started))?.GetValue<string>()
                ?? throw new InvalidOperationException("Excel did not return the exact workbook identity.");
            if (!MatchesPath(actualPath, target))
            {
                throw new InvalidOperationException("The workbook identity changed during native Excel discovery.");
            }
            return workbook;
        }
        catch
        {
            workbook.Dispose();
            throw;
        }
    }

    private static bool MatchesPath(string path, string target) =>
        Path.IsPathFullyQualified(path)
        && string.Equals(MacPathCanonicalizer.Normalize(path), target, StringComparison.Ordinal);

    internal static MacAppleEvents.Descriptor Resolve(string filePath, TimeSpan timeout) =>
        Find(filePath, timeout)
        ?? throw new InvalidOperationException("Workbook is not open in this ExcelMcp session.");

    internal static void PrepareOpen(string filePath, TimeSpan timeout)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        var target = MacPathCanonicalizer.Normalize(filePath);
        var started = Stopwatch.GetTimestamp();
        var paths = ReadNames(fullPaths: true, timeout);
        if (paths.Any(path => MatchesPath(path, target)))
        {
            throw new InvalidOperationException(
                "Workbook is already open in shared Excel. Reuse its owning session or close it before opening a new session.");
        }

        var names = ReadNames(fullPaths: false, timeout - Stopwatch.GetElapsedTime(started));
        if (names.Contains(Path.GetFileName(target), StringComparer.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                "A workbook with the same name is already open in shared Excel. Close it before opening this workbook.");
        }
    }

    internal static string[] ReadNames(bool fullPaths, TimeSpan timeout)
    {
        if (timeout <= TimeSpan.Zero)
        {
            throw new TimeoutException("Mac Excel workbook discovery timed out; the operation did not complete.");
        }
        var started = Stopwatch.GetTimestamp();
        using var application = MacAppleEvents.Create(MacAppleEvents.Code("null"), []);
        if (MacNativeRange.Count(application, MacExcelDictionary.WorkbookClass, timeout) == 0)
            return [];
        using var all = MacAppleEvents.Create(MacAppleEvents.Code("enum"), BitConverter.GetBytes(MacAppleEvents.Code("all ")));
        using var workbooks = MacAppleEvents.Object(MacExcelDictionary.WorkbookClass, application, MacAppleEvents.Code("indx"), all);
        using var names = MacAppleEvents.Property(workbooks, fullPaths ? MacExcelDictionary.FullName : MacExcelDictionary.Name);
        using var appleEvent = MacAppleEvents.Event(MacAppleEvents.Code("core"), MacAppleEvents.Code("getd"));
        MacAppleEvents.Put(appleEvent, MacAppleEvents.Code("----"), names);
        if (MacAppleEvents.Send(appleEvent, MacAppleEvents.Remaining(timeout, started)) is not JsonArray values)
        {
            throw new InvalidOperationException("Excel did not return the requested workbook-name list.");
        }
        return values.Select(value => value?.GetValue<string>()
            ?? throw new InvalidOperationException("Excel returned an empty workbook name.")).ToArray();
    }
}
