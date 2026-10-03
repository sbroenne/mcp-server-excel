using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration;

internal static class SessionTestCleanup
{
    internal static void AssertExitedAndDelete(
        OwnedExcelProcessScope owned, IEnumerable<string> files, string? directory = null)
    {
        var failures = new List<Exception>();
        Capture(() => owned.AssertAllExited(expectProcess: false));
        owned.Dispose();
        foreach (var file in files.Where(File.Exists))
        {
            Capture(() => File.Delete(file));
        }
        if (directory is not null && Directory.Exists(directory))
        {
            Capture(() => Directory.Delete(directory, recursive: true));
        }
        if (failures.Count > 0)
        {
            throw new AggregateException("Session test cleanup failed.", failures);
        }

        void Capture(Action action)
        {
            var failure = Record.Exception(action);
            if (failure is not null)
            {
                failures.Add(failure);
            }
        }
    }
}
