using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Service;

namespace Sbroenne.ExcelMcp.Benchmarks;

internal sealed class BenchmarkContext : IDisposable
{
    private readonly string _workingDirectory;

    public BenchmarkContext(BenchmarkOptions options)
    {
        Options = options;
        Directory.CreateDirectory(options.OutputDirectory);
        _workingDirectory = Path.Combine(
            Path.GetTempPath(),
            "excelmcp-benchmarks",
            $"{DateTime.UtcNow:yyyyMMddHHmmss}-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_workingDirectory);
    }

    public BenchmarkOptions Options { get; }

    public BenchmarkConfiguration Configuration => Options.Configuration;

    public string CreateWorkingPath(string prefix, string extension = ".xlsx") =>
        Path.Combine(_workingDirectory, $"{prefix}-{Guid.NewGuid():N}{extension}");

    public string CreateSafetyRoot(string prefix)
    {
        var path = Path.Combine(_workingDirectory, $"{prefix}-safety-{Guid.NewGuid():N}");
        Directory.CreateDirectory(path);
        return path;
    }

    public static void CreateEmptyWorkbook(string path)
    {
        _ = ExcelSession.CreateNew(path, isMacroEnabled: false, static (_, _) => 0);
    }

    public static void CreateDataWorkbook(string path, int rows, int columns, bool includeTable = false)
    {
        CreateEmptyWorkbook(path);
        using var batch = ExcelSession.BeginBatch(show: false, operationTimeout: null, path);
        batch.Execute((context, _) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? tables = null;
            dynamic? table = null;
            try
            {
                sheet = context.Book.Worksheets[1];
                sheet.Name = "Data";
                var values = (object[,])Array.CreateInstance(typeof(object), [rows, columns], [1, 1]);
                for (var row = 1; row <= rows; row++)
                {
                    for (var column = 1; column <= columns; column++)
                    {
                        values[row, column] = row == 1
                            ? $"Column{column}"
                            : (row - 1) * 1000d + column;
                    }
                }

                var address = $"A1:{ToColumnName(columns)}{rows}";
                range = sheet.Range[address];
                range.Value2 = values;
                if (includeTable)
                {
                    tables = sheet.ListObjects;
                    table = tables.Add(1, range, Type.Missing, 1, Type.Missing);
                    table.Name = "BenchmarkTable";
                }

                return 0;
            }
            finally
            {
                ComUtilities.Release(ref table);
                ComUtilities.Release(ref tables);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
        batch.Save();
    }

    public string CopyWorkbook(string sourcePath, string prefix)
    {
        var destination = CreateWorkingPath(prefix, Path.GetExtension(sourcePath));
        File.Copy(sourcePath, destination);
        return destination;
    }

    public static double ElapsedMilliseconds(long startedTimestamp) =>
        Stopwatch.GetElapsedTime(startedTimestamp).TotalMilliseconds;

    public static bool WaitForProcessExit(int processId, TimeSpan timeout, out double elapsedMilliseconds)
    {
        var started = Stopwatch.GetTimestamp();
        try
        {
            using var process = Process.GetProcessById(processId);
            var exited = process.WaitForExit((int)Math.Min(int.MaxValue, timeout.TotalMilliseconds));
            elapsedMilliseconds = ElapsedMilliseconds(started);
            return exited;
        }
        catch (ArgumentException)
        {
            elapsedMilliseconds = ElapsedMilliseconds(started);
            return true;
        }
    }

    public static bool IsProcessAlive(int processId)
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

    public static long GetWorkingSetBytes()
    {
        using var process = Process.GetCurrentProcess();
        process.Refresh();
        return process.WorkingSet64;
    }

    public static double EstimateTokensFromUtf8Bytes(long bytes) => Math.Ceiling(bytes / 4d);

    public static string Sha256(string value) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();

    public static string ToColumnName(int column)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(column, 1);
        var builder = new StringBuilder();
        var value = column;
        while (value > 0)
        {
            value--;
            builder.Insert(0, (char)('A' + (value % 26)));
            value /= 26;
        }

        return builder.ToString();
    }

    public static string GetRequiredString(string? json, string propertyName)
    {
        using var document = JsonDocument.Parse(json ?? throw new InvalidDataException("Service returned no JSON result."));
        return document.RootElement.GetProperty(propertyName).GetString()
            ?? throw new InvalidDataException($"Service result property '{propertyName}' was null.");
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_workingDirectory))
            {
                Directory.Delete(_workingDirectory, recursive: true);
            }
        }
        catch (IOException)
        {
            // Keep diagnostic work files when Excel or antivirus still has a transient handle.
        }
        catch (UnauthorizedAccessException)
        {
            // Keep diagnostic work files when the environment prevents cleanup.
        }
    }
}

internal static class ServiceBenchmarkHelpers
{
    public static async Task<string> CreateSessionAsync(
        ExcelMcpService service,
        string workbookPath,
        bool showExcel,
        int? timeoutSeconds = null)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(new
            {
                filePath = workbookPath,
                show = showExcel,
                timeoutSeconds
            }, ServiceProtocol.JsonOptions),
            Source = "benchmark"
        });
        EnsureSuccess(response, "session.open");
        return BenchmarkContext.GetRequiredString(response.Result, "sessionId");
    }

    public static async Task ConfigureSafetyAsync(
        ExcelMcpService service,
        string sessionId,
        string reviewMode,
        string checkpointMode,
        string journalMode = "on",
        string verificationMode = "on")
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.configure-safety",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new
            {
                reviewMode,
                checkpointMode,
                journalMode,
                verificationMode,
                abnormalShutdownPolicy = "discardWithRecoveryEvidence"
            }, ServiceProtocol.JsonOptions),
            Source = "benchmark"
        });
        EnsureSuccess(response, "session.configure-safety");
    }

    public static async Task CloseSessionAsync(ExcelMcpService service, string sessionId, bool save = false)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new { save }, ServiceProtocol.JsonOptions),
            Source = "benchmark"
        });
        EnsureSuccess(response, "session.close");
    }

    public static void EnsureSuccess(ServiceResponse response, string operation)
    {
        if (!response.Success)
        {
            throw new InvalidOperationException($"{operation} failed: {response.ErrorCategory}: {response.ErrorMessage}");
        }
    }

    public static long SerializedPayloadBytes(ServiceRequest request, ServiceResponse response) =>
        Encoding.UTF8.GetByteCount(JsonSerializer.Serialize(request, ServiceProtocol.JsonOptions)) +
        Encoding.UTF8.GetByteCount(JsonSerializer.Serialize(response, ServiceProtocol.JsonOptions));
}
