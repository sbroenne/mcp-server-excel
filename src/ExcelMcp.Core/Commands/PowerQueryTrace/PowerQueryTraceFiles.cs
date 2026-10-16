using System.Diagnostics;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PowerQueryTrace;

internal static class PowerQueryTraceFiles
{
    public static int Count(string folder)
    {
        if (!Directory.Exists(folder))
        {
            return 0;
        }

        return EnumerateTraceFiles(folder).Count();
    }

    public static IEnumerable<FileInfo> EnumerateTraceFiles(string folder)
    {
        if (!Directory.Exists(folder))
        {
            return [];
        }

        return Directory.EnumerateFiles(folder)
            .Select(path => new FileInfo(path))
            .Where(file => PowerQueryTraceFileName.TryParse(file.Name, out _, out _, out _));
    }
}

/// <summary>
/// Reads local Power Query trace files and produces file lists and summaries.
/// </summary>
internal sealed class PowerQueryTraceReader
{
    private const int ContainerHostPeekLineCount = 50;
    private readonly string _traceFolder;
    private readonly Func<int, bool> _isProcessRunning;

    public PowerQueryTraceReader(string traceFolder, Func<int, bool>? isProcessRunning = null)
    {
        _traceFolder = traceFolder ?? throw new ArgumentNullException(nameof(traceFolder));
        _isProcessRunning = isProcessRunning ?? IsProcessRunning;
    }

    public PowerQueryTraceListResult List(
        int? processId,
        DateTime? sinceUtc,
        DateTime? processStartedUtc = null)
    {
        var files = new List<PowerQueryTraceFileInfo>();
        foreach (var file in PowerQueryTraceFiles.EnumerateTraceFiles(_traceFolder))
        {
            if (!PowerQueryTraceFileName.TryParse(file.Name, out string kind, out int fileProcessId, out DateTime? startedUtc))
            {
                continue;
            }

            if (processStartedUtc is DateTime processStart &&
                (startedUtc is not DateTime fileStart || fileStart < processStart))
            {
                continue;
            }

            int? hostProcessId = null;
            if (processId is int requestedProcessId)
            {
                if (kind == PowerQueryTraceFileName.ExcelKind)
                {
                    if (fileProcessId != requestedProcessId)
                    {
                        continue;
                    }
                }
                else
                {
                    hostProcessId = FindHostProcessId(file.FullName, requestedProcessId);
                    if (hostProcessId != requestedProcessId)
                    {
                        continue;
                    }
                }
            }
            else if (kind == PowerQueryTraceFileName.ContainerKind && sinceUtc.HasValue)
            {
                hostProcessId = FindHostProcessId(file.FullName, expectedProcessId: null);
            }

            DateTime lastWriteUtc = file.LastWriteTimeUtc;
            if (sinceUtc is DateTime since &&
                lastWriteUtc < since &&
                (startedUtc is not DateTime started || started < since) &&
                !_isProcessRunning(fileProcessId) &&
                (hostProcessId is not int hostPid || !_isProcessRunning(hostPid)))
            {
                continue;
            }

            files.Add(new PowerQueryTraceFileInfo
            {
                FileName = file.Name,
                Kind = kind,
                ProcessId = fileProcessId,
                HostProcessId = hostProcessId,
                StartedUtc = startedUtc,
                LastWriteUtc = lastWriteUtc,
                SizeBytes = file.Length
            });
        }

        files.Sort(static (left, right) => right.LastWriteUtc.CompareTo(left.LastWriteUtc));
        return new PowerQueryTraceListResult
        {
            Success = true,
            TraceFolder = _traceFolder,
            ProcessId = processId,
            Files = files,
            TotalSizeBytes = files.Sum(file => file.SizeBytes)
        };
    }

    public PowerQueryTraceSummaryResult Summarize(PowerQueryTraceFilter filter)
    {
        var analyzer = new PowerQueryTraceAnalyzer(filter);
        foreach (var file in SelectFilesForSummary(filter))
        {
            analyzer.AddFileRead();
            using var stream = new FileStream(
                file.FullName,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete);
            using var reader = new StreamReader(stream);
            while (reader.ReadLine() is { } line)
            {
                if (PowerQueryTraceLineParser.TryParse(line, out var entry) && entry != null)
                {
                    analyzer.Add(entry);
                }
                else
                {
                    analyzer.AddUnreadableLine();
                }
            }
        }

        return analyzer.Build();
    }

    private IEnumerable<FileInfo> SelectFilesForSummary(PowerQueryTraceFilter filter)
    {
        foreach (var file in PowerQueryTraceFiles.EnumerateTraceFiles(_traceFolder))
        {
            if (!PowerQueryTraceFileName.TryParse(file.Name, out string kind, out int fileProcessId, out DateTime? startedUtc))
            {
                continue;
            }

            if (filter.UntilUtc is DateTime until && startedUtc is DateTime start && start > until)
            {
                continue;
            }

            if (filter.ProcessStartedUtc is DateTime processStart &&
                (startedUtc is not DateTime fileStart || fileStart < processStart))
            {
                continue;
            }

            int? hostProcessId = null;
            if (filter.ProcessId is int processId)
            {
                if (kind == PowerQueryTraceFileName.ExcelKind && fileProcessId != processId)
                {
                    continue;
                }

                if (kind == PowerQueryTraceFileName.ContainerKind)
                {
                    hostProcessId = FindHostProcessId(file.FullName, processId);
                    if (hostProcessId != processId)
                    {
                        continue;
                    }
                }
            }

            if (filter.SinceUtc is DateTime since &&
                file.LastWriteTimeUtc < since &&
                (startedUtc is not DateTime started || started < since))
            {
                hostProcessId ??= kind == PowerQueryTraceFileName.ContainerKind
                    ? FindHostProcessId(file.FullName, expectedProcessId: null)
                    : null;
                if (!_isProcessRunning(fileProcessId) &&
                    (hostProcessId is not int hostPid || !_isProcessRunning(hostPid)))
                {
                    continue;
                }
            }

            yield return file;
        }
    }

    private static int? FindHostProcessId(string filePath, int? expectedProcessId)
    {
        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var reader = new StreamReader(stream);
        for (int i = 0; i < ContainerHostPeekLineCount && reader.ReadLine() is { } line; i++)
        {
            if (PowerQueryTraceLineParser.TryParse(line, out var entry) &&
                entry?.HostProcessId is int hostProcessId &&
                (expectedProcessId is null || hostProcessId == expectedProcessId))
            {
                return hostProcessId;
            }
        }

        return null;
    }

    private static bool IsProcessRunning(int processId)
    {
        try
        {
            using var process = Process.GetProcessById(processId);
            return !process.HasExited;
        }
        catch (ArgumentException)
        {
            return false;
        }
    }
}
