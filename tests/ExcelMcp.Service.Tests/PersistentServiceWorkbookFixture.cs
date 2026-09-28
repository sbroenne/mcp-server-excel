using System.Collections.Concurrent;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Tests.Infrastructure;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public class PersistentServiceWorkbookFixture : IAsyncLifetime, IDisposable
{
    private readonly Action<string>? _createWorkbook;
    private readonly bool _show;
    private readonly string _workbookFileName = "RangeValuesFixture.xlsx";
    private readonly ConcurrentDictionary<ExcelProcessIdentity, byte> _launchedIdentities = new();
    private readonly IExcelBatch _batchToken = new ServiceBatchToken();
    private readonly string _tempDirectory = Path.Combine(
        Path.GetTempPath(),
        $"ExcelMcpServiceFixture_{Guid.NewGuid():N}");
    private ExcelMcpService? _service;
    private int _sheetCounter;
    private string? _sessionId;

    public int ExcelLaunchCount => _launchedIdentities.Count;

    internal IExcelBatch BatchToken => _batchToken;
    internal string WorkbookPath { get; private set; } = string.Empty;

    public PersistentServiceWorkbookFixture()
    {
    }

    protected PersistentServiceWorkbookFixture(bool show)
    {
        _show = show;
    }

    protected PersistentServiceWorkbookFixture(
        string workbookFileName,
        bool show = false)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(workbookFileName);
        _workbookFileName = workbookFileName;
        _show = show;
    }

    protected PersistentServiceWorkbookFixture(
        Action<string> createWorkbook,
        string workbookFileName = "RangeValuesFixture.xlsx",
        bool show = false)
    {
        ArgumentNullException.ThrowIfNull(createWorkbook);
        ArgumentException.ThrowIfNullOrWhiteSpace(workbookFileName);
        _createWorkbook = createWorkbook;
        _workbookFileName = workbookFileName;
        _show = show;
    }

    public async Task InitializeAsync()
    {
        if (SessionManager.GetTrackedExcelProcesses().Count != 0)
        {
            throw new InvalidOperationException(
                $"Tracked Excel identities existed before fixture startup: " +
                $"{string.Join(", ", SessionManager.GetTrackedExcelProcesses())}");
        }

        Directory.CreateDirectory(_tempDirectory);
        var workbookPath = Path.Combine(_tempDirectory, _workbookFileName);
        WorkbookPath = workbookPath;
        _service = new ExcelMcpService();
        ServiceResponse response;
        if (_createWorkbook is null)
        {
            response = await SendAsync(
                "session.create",
                new { filePath = workbookPath, show = _show },
                sessionId: null);
        }
        else
        {
            _createWorkbook(workbookPath);
            response = await SendAsync(
                "session.open",
                new { filePath = workbookPath, show = _show },
                sessionId: null);
        }

        using var document = ParseResult(response);
        _sessionId = document.RootElement.GetProperty("sessionId").GetString();
        if (string.IsNullOrWhiteSpace(_sessionId))
        {
            throw new InvalidOperationException("session.create returned no session ID.");
        }

        foreach (var identity in SessionManager.GetTrackedExcelProcesses())
        {
            _launchedIdentities.TryAdd(identity, 0);
        }

        if (_service.SessionCount != 1 || _launchedIdentities.Count != 1)
        {
            throw new InvalidOperationException(
                $"Fixture startup expected one session and one Excel identity, but found " +
                $"{_service.SessionCount} session(s) and {_launchedIdentities.Count} identity/identities.");
        }
    }

    public async Task DisposeAsync()
    {
        var service = _service;
        Exception? failure = null;
        try
        {
            if (service is not null && _sessionId is not null)
            {
                await SendAsync("session.close", new { save = false }, _sessionId);
                _sessionId = null;
            }
        }
        catch (Exception ex)
        {
            failure = PersistentServiceCleanupFailures.Combine(failure, ex);
        }

        try
        {
            if (service is not null && service.SessionCount != 0)
            {
                failure = PersistentServiceCleanupFailures.Combine(
                    failure,
                    new InvalidOperationException(
                        $"Fixture shutdown left {service.SessionCount} Service session(s)."));
            }
        }
        catch (Exception ex)
        {
            failure = PersistentServiceCleanupFailures.Combine(failure, ex);
        }

        try
        {
            service?.Dispose();
        }
        catch (Exception ex)
        {
            failure = PersistentServiceCleanupFailures.Combine(failure, ex);
        }
        finally
        {
            _service = null;
            _sessionId = null;
        }

        if (!SpinWait.SpinUntil(
                () => SessionManager.GetTrackedExcelProcesses().Count == 0,
                TimeSpan.FromSeconds(15)))
        {
            failure = PersistentServiceCleanupFailures.Combine(
                failure,
                new InvalidOperationException(
                    $"Fixture shutdown left tracked Excel identities: " +
                    $"{string.Join(", ", SessionManager.GetTrackedExcelProcesses())}"));
        }

        try
        {
            if (Directory.Exists(_tempDirectory))
            {
                Directory.Delete(_tempDirectory, recursive: true);
            }
        }
        catch (Exception ex)
        {
            failure = PersistentServiceCleanupFailures.Combine(
                failure,
                new InvalidOperationException(
                    $"Fixture temporary directory cleanup failed: {ex.Message}",
                    ex));
        }

        if (failure is not null)
        {
            throw failure;
        }
    }

    public void Dispose()
    {
        var service = _service;
        try
        {
            service?.Dispose();
        }
        finally
        {
            _service = null;
            _sessionId = null;
            GC.SuppressFinalize(this);
        }
    }

    public string CreateSheetName(string scenario)
    {
        var counter = Interlocked.Increment(ref _sheetCounter);
        var prefix = $"S{counter:D2}_";
        var availableLength = 31 - prefix.Length;
        var suffix = scenario.Length <= availableLength
            ? scenario
            : scenario[..availableLength];
        return $"{prefix}{suffix}";
    }

    internal PersistentServiceWorkbookTestScope CreateScope() => new(this);

    internal T CreateCommands<T>()
        where T : class =>
        ServiceCommandProxy.Create<T>(this);

    internal async Task SaveAndReopenAsync()
    {
        if (_service is null || string.IsNullOrWhiteSpace(_sessionId))
        {
            throw new InvalidOperationException(
                "The persistent Service fixture is not initialized.");
        }

        var sessionId = _sessionId;
        await SendAsync("session.close", new { save = true }, sessionId);
        _sessionId = null;

        var response = await SendAsync(
            "session.open",
            new { filePath = WorkbookPath, show = false },
            sessionId: null);
        using var document = ParseResult(response);
        _sessionId = document.RootElement.GetProperty("sessionId").GetString();
        if (string.IsNullOrWhiteSpace(_sessionId))
        {
            throw new InvalidOperationException("session.open returned no session ID.");
        }

        foreach (var identity in SessionManager.GetTrackedExcelProcesses())
        {
            _launchedIdentities.TryAdd(identity, 0);
        }

        if (_service.SessionCount != 1)
        {
            throw new InvalidOperationException(
                $"Fixture reopen expected one session, but found {_service.SessionCount}.");
        }
    }

    internal void ExecuteRawVerification(
        Action<ExcelContext, CancellationToken> operation)
    {
        ExecuteRawVerification<object?>((context, cancellationToken) =>
        {
            operation(context, cancellationToken);
            return null;
        });
    }

    internal T ExecuteRawVerification<T>(
        Func<ExcelContext, CancellationToken, T> operation)
    {
        if (_service is null || string.IsNullOrWhiteSpace(_sessionId))
        {
            throw new InvalidOperationException(
                "The persistent Service fixture is not initialized.");
        }
        if (!_service.SessionManager.TryBeginOperation(
                _sessionId,
                out var batch,
                out var errorMessage))
        {
            throw new InvalidOperationException(
                $"Could not acquire the Service workbook for raw test verification: {errorMessage}");
        }

        try
        {
            return batch.Execute(operation);
        }
        finally
        {
            _service.SessionManager.EndOperation(_sessionId);
        }
    }

    internal void ValidateBatchToken(object? value)
    {
        if (!ReferenceEquals(value, _batchToken))
        {
            throw new InvalidOperationException(
                "Service-backed command received an unexpected Excel batch. " +
                "Direct Core batches are not valid in this fixture.");
        }
    }

    public Task<ServiceResponse> SendAsync(string command, object args) =>
        SendAsync(command, args, _sessionId);

    internal async Task<ServiceResponse> SendForFailureAsync(
        string command,
        object args)
    {
        if (_service is null)
        {
            throw new InvalidOperationException(
                "The persistent Service fixture is not initialized.");
        }

        var response = await _service.ProcessAsync(new ServiceRequest
        {
            Command = command,
            SessionId = _sessionId,
            Args = JsonSerializer.Serialize(args, ServiceProtocol.JsonOptions),
            Source = "persistent-service-failure-fixture"
        });

        if (response.Success || string.IsNullOrWhiteSpace(response.ErrorMessage))
        {
            throw new InvalidOperationException(
                $"{command} unexpectedly returned a successful response.");
        }

        return response;
    }

    internal string CreateInputFile(string extension, string content)
    {
        var path = Path.Combine(_tempDirectory, $"{Guid.NewGuid():N}{extension}");
        File.WriteAllText(path, content);
        return path;
    }

    internal string CreateInputFile(string extension, byte[] content)
    {
        var path = Path.Combine(_tempDirectory, $"{Guid.NewGuid():N}{extension}");
        File.WriteAllBytes(path, content);
        return path;
    }

    internal string CreateBlankWorkbook(string scenario, string suffix)
    {
        var safeScenario = string.Concat(
            scenario.Select(character =>
                Path.GetInvalidFileNameChars().Contains(character) ? '_' : character));
        var path = Path.Combine(
            _tempDirectory,
            $"{safeScenario}_{suffix}_{Guid.NewGuid():N}.xlsx");
        return SavedWorkbookTemplates.CopyBlankTo(path);
    }

    private async Task<ServiceResponse> SendAsync(
        string command,
        object args,
        string? sessionId)
    {
        if (_service is null)
        {
            throw new InvalidOperationException("The persistent Service fixture is not initialized.");
        }

        var response = await _service.ProcessAsync(new ServiceRequest
        {
            Command = command,
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(args, ServiceProtocol.JsonOptions),
            Source = "persistent-service-range-fixture"
        });

        if (!response.Success || !string.IsNullOrEmpty(response.ErrorMessage))
        {
            throw CreateServiceException(command, response);
        }
        if (response.Command is not null
            && !string.Equals(response.Command, command, StringComparison.Ordinal))
        {
            throw new InvalidOperationException(
                $"{command} returned mismatched command '{response.Command}'.");
        }

        return response;
    }

    private static Exception CreateServiceException(
        string command,
        ServiceResponse response)
    {
        var message =
            $"{command} failed [{response.ErrorCategory}/{response.ExceptionType}]: " +
            $"{response.ErrorMessage}; inner={response.InnerError}";
        return response.ExceptionType switch
        {
            nameof(ArgumentException) => new ArgumentException(message),
            nameof(ArgumentNullException) => new ArgumentNullException(null, message),
            nameof(ArgumentOutOfRangeException) => new ArgumentOutOfRangeException(null, message),
            nameof(DirectoryNotFoundException) => new DirectoryNotFoundException(message),
            nameof(FileNotFoundException) => new FileNotFoundException(message),
            nameof(IOException) => new IOException(message),
            nameof(KeyNotFoundException) => new KeyNotFoundException(message),
            nameof(NotSupportedException) => new NotSupportedException(message),
            nameof(OperationCanceledException) => new OperationCanceledException(message),
            nameof(System.Reflection.TargetParameterCountException) =>
                new System.Reflection.TargetParameterCountException(message),
            nameof(TimeoutException) => new TimeoutException(message),
            nameof(UnauthorizedAccessException) => new UnauthorizedAccessException(message),
            _ => new InvalidOperationException(message)
        };
    }

    private static JsonDocument ParseResult(ServiceResponse response)
    {
        if (string.IsNullOrWhiteSpace(response.Result))
        {
            throw new InvalidOperationException("Service response contained no result.");
        }
        return JsonDocument.Parse(response.Result);
    }

    private sealed class ServiceBatchToken : IExcelBatch
    {
        private const string Message =
            "The Service fixture batch token cannot execute COM directly.";

        public string WorkbookPath => throw new InvalidOperationException(Message);
        public ILogger Logger => NullLogger.Instance;
        public IReadOnlyDictionary<string, Excel.Workbook> Workbooks =>
            throw new InvalidOperationException(Message);
        public bool HasTimedOutOperation => false;
        public int? ExcelProcessId => throw new InvalidOperationException(Message);
        public TimeSpan OperationTimeout => throw new InvalidOperationException(Message);
        public bool IsExcelVisible => throw new InvalidOperationException(Message);

        public void Dispose()
        {
        }

        public void Execute(
            Action<ExcelContext, CancellationToken> operation,
            CancellationToken cancellationToken = default) =>
            throw new InvalidOperationException(Message);

        public T Execute<T>(
            Func<ExcelContext, CancellationToken, T> operation,
            CancellationToken cancellationToken = default) =>
            throw new InvalidOperationException(Message);

        public Excel.Workbook GetWorkbook(string filePath) =>
            throw new InvalidOperationException(Message);

        public bool IsExcelProcessAlive() =>
            throw new InvalidOperationException(Message);

        public void Save(CancellationToken cancellationToken = default) =>
            throw new InvalidOperationException(Message);

        public void UpdateWorkbookPath(string workbookPath) =>
            throw new InvalidOperationException(Message);
    }
}

internal static class PersistentServiceCleanupFailures
{
    internal static Exception Combine(
        Exception? existing,
        Exception next)
    {
        ArgumentNullException.ThrowIfNull(next);
        return existing is null
            ? next
            : new AggregateException(
                "Persistent Service fixture cleanup failed.",
                existing,
                next);
    }
}

public sealed class PersistentServiceWorkbookTestScope(
    PersistentServiceWorkbookFixture fixture) : IAsyncDisposable
{
    private readonly List<string> _sheets = [];
    private readonly List<string> _powerQueries = [];
    private readonly List<string> _xmlMaps = [];
    private readonly List<string> _connections = [];
    private readonly List<string> _namedRanges = [];
    private readonly List<string> _dataModelTables = [];
    private readonly List<string> _dataModelMeasures = [];
    private readonly List<string> _tables = [];
    private readonly List<(string PivotTableName, string MemberName)> _calculatedMembers = [];
    private readonly List<string> _vbaModules = [];

    internal IExcelBatch BatchToken => fixture.BatchToken;
    internal string WorkbookPath => fixture.WorkbookPath;

    internal void ExecuteRawVerification(
        Action<ExcelContext, CancellationToken> operation) =>
        fixture.ExecuteRawVerification(operation);

    internal T ExecuteRawVerification<T>(
        Func<ExcelContext, CancellationToken, T> operation) =>
        fixture.ExecuteRawVerification(operation);

    internal T CreateCommands<T>()
        where T : class =>
        fixture.CreateCommands<T>();

    internal Task SaveAndReopenAsync() =>
        fixture.SaveAndReopenAsync();

    internal string CreateTestSheet(
        IExcelBatch batch,
        [System.Runtime.CompilerServices.CallerMemberName] string testName = "")
    {
        fixture.ValidateBatchToken(batch);
        var sheetName = fixture.CreateSheetName(testName);
        return CreateNamedTestSheet(batch, sheetName);
    }

    internal string CreateNamedTestSheet(IExcelBatch batch, string sheetName)
    {
        fixture.ValidateBatchToken(batch);
        fixture.SendAsync("sheet.create", new { sheetName }).GetAwaiter().GetResult();
        _sheets.Add(sheetName);
        return sheetName;
    }

    internal ServiceResponse Send(string command, object args) =>
        fixture.SendAsync(command, args).GetAwaiter().GetResult();

    internal Task<ServiceResponse> SendForFailureAsync(
        string command,
        object args) =>
        fixture.SendForFailureAsync(command, args);

    internal void RegisterSheetForCleanup(string sheetName) =>
        _sheets.Add(sheetName);

    internal void RenameTrackedSheet(string oldName, string newName)
    {
        var index = _sheets.IndexOf(oldName);
        if (index < 0)
        {
            throw new InvalidOperationException(
                $"Sheet '{oldName}' is not registered for cleanup.");
        }

        _sheets[index] = newName;
    }

    internal void ForgetSheet(string sheetName)
    {
        if (!_sheets.Remove(sheetName))
        {
            throw new InvalidOperationException(
                $"Sheet '{sheetName}' is not registered for cleanup.");
        }
    }

    internal string CreateInputFile(string extension, string content) =>
        fixture.CreateInputFile(extension, content);

    internal string CreateInputFile(string extension, byte[] content) =>
        fixture.CreateInputFile(extension, content);

    internal string CreateBlankWorkbook(string scenario, string suffix) =>
        fixture.CreateBlankWorkbook(scenario, suffix);

    internal void RegisterXmlMapForCleanup(string mapName) =>
        _xmlMaps.Add(mapName);

    internal void RegisterPowerQueryForCleanup(string queryName) =>
        _powerQueries.Add(queryName);

    internal void RegisterConnectionForCleanup(string connectionName)
    {
        if (!_connections.Contains(connectionName, StringComparer.Ordinal))
        {
            _connections.Add(connectionName);
        }
    }

    internal void RegisterNamedRangeForCleanup(string namedRange)
    {
        if (!_namedRanges.Contains(namedRange, StringComparer.Ordinal))
        {
            _namedRanges.Add(namedRange);
        }
    }

    internal void RegisterDataModelTableForCleanup(string tableName)
    {
        if (!_dataModelTables.Contains(tableName, StringComparer.Ordinal))
        {
            _dataModelTables.Add(tableName);
        }
    }

    internal void RegisterDataModelMeasureForCleanup(string measureName)
    {
        if (!_dataModelMeasures.Contains(measureName, StringComparer.Ordinal))
        {
            _dataModelMeasures.Add(measureName);
        }
    }

    internal void RegisterTableForCleanup(string tableName)
    {
        if (!_tables.Contains(tableName, StringComparer.Ordinal))
        {
            _tables.Add(tableName);
        }
    }

    internal void RegisterCalculatedMemberForCleanup(
        string pivotTableName,
        string memberName)
    {
        var item = (pivotTableName, memberName);
        if (!_calculatedMembers.Contains(item))
        {
            _calculatedMembers.Add(item);
        }
    }

    internal void RegisterVbaModuleForCleanup(string moduleName)
    {
        if (!_vbaModules.Contains(moduleName, StringComparer.Ordinal))
        {
            _vbaModules.Add(moduleName);
        }
    }

    internal void ForgetVbaModule(string moduleName)
    {
        if (!_vbaModules.Remove(moduleName))
        {
            throw new InvalidOperationException(
                $"VBA module '{moduleName}' is not registered for cleanup.");
        }
    }

    internal void ForgetCalculatedMember(
        string pivotTableName,
        string memberName)
    {
        if (!_calculatedMembers.Remove((pivotTableName, memberName)))
        {
            throw new InvalidOperationException(
                $"Calculated member '{memberName}' is not registered for cleanup.");
        }
    }

    internal void ForgetTable(string tableName)
    {
        if (!_tables.Remove(tableName))
        {
            throw new InvalidOperationException(
                $"Table '{tableName}' is not registered for cleanup.");
        }
    }

    internal void ForgetDataModelMeasure(string measureName)
    {
        if (!_dataModelMeasures.Remove(measureName))
        {
            throw new InvalidOperationException(
                $"Data Model measure '{measureName}' is not registered for cleanup.");
        }
    }

    internal void RenameTrackedPowerQuery(string oldName, string newName)
    {
        var index = _powerQueries.IndexOf(oldName);
        if (index < 0)
        {
            throw new InvalidOperationException(
                $"Power Query '{oldName}' is not registered for cleanup.");
        }

        _powerQueries[index] = newName;
    }

    internal void ForgetPowerQuery(string queryName)
    {
        if (!_powerQueries.Remove(queryName))
        {
            throw new InvalidOperationException(
                $"Power Query '{queryName}' is not registered for cleanup.");
        }
    }

    internal void ForgetXmlMap(string mapName)
    {
        if (!_xmlMaps.Remove(mapName))
        {
            throw new InvalidOperationException(
                $"XML map '{mapName}' is not registered for cleanup.");
        }
    }

    internal void ForgetConnection(string connectionName)
    {
        if (!_connections.Remove(connectionName))
        {
            throw new InvalidOperationException(
                $"Connection '{connectionName}' is not registered for cleanup.");
        }
    }

    internal void ForgetNamedRange(string namedRange)
    {
        if (!_namedRanges.Remove(namedRange))
        {
            throw new InvalidOperationException(
                $"Named range '{namedRange}' is not registered for cleanup.");
        }
    }

    public async ValueTask DisposeAsync()
    {
        List<Exception>? failures = null;
        for (var index = _vbaModules.Count - 1; index >= 0; index--)
        {
            try
            {
                await fixture.SendAsync(
                    "vba.delete",
                    new { moduleName = _vbaModules[index] });
            }
            catch (Exception ex)
            {
                failures ??= [];
                failures.Add(ex);
            }
        }

        for (var index = _calculatedMembers.Count - 1; index >= 0; index--)
        {
            try
            {
                var item = _calculatedMembers[index];
                await fixture.SendAsync(
                    "pivottablecalc.delete-calculated-member",
                    new
                    {
                        pivotTableName = item.PivotTableName,
                        memberName = item.MemberName
                    });
            }
            catch (Exception ex)
            {
                failures ??= [];
                failures.Add(ex);
            }
        }

        for (var index = _xmlMaps.Count - 1; index >= 0; index--)
        {
            try
            {
                await fixture.SendAsync(
                    "xmlmap.delete",
                    new { mapName = _xmlMaps[index] });
            }
            catch (Exception ex)
            {
                failures ??= [];
                failures.Add(ex);
            }
        }

        for (var index = _powerQueries.Count - 1; index >= 0; index--)
        {
            try
            {
                await fixture.SendAsync(
                    "powerquery.delete",
                    new { queryName = _powerQueries[index] });
            }
            catch (Exception ex)
            {
                failures ??= [];
                failures.Add(ex);
            }
        }

        for (var index = _connections.Count - 1; index >= 0; index--)
        {
            try
            {
                await fixture.SendAsync(
                    "connection.delete",
                    new { connectionName = _connections[index] });
            }
            catch (Exception ex)
            {
                failures ??= [];
                failures.Add(ex);
            }
        }

        for (var index = _dataModelMeasures.Count - 1; index >= 0; index--)
        {
            try
            {
                await fixture.SendAsync(
                    "datamodel.delete-measure",
                    new { measureName = _dataModelMeasures[index] });
            }
            catch (Exception ex)
            {
                failures ??= [];
                failures.Add(ex);
            }
        }

        for (var index = _dataModelTables.Count - 1; index >= 0; index--)
        {
            try
            {
                await fixture.SendAsync(
                    "datamodel.delete-table",
                    new { tableName = _dataModelTables[index] });
            }
            catch (Exception ex)
            {
                failures ??= [];
                failures.Add(ex);
            }
        }

        for (var index = _namedRanges.Count - 1; index >= 0; index--)
        {
            try
            {
                await fixture.SendAsync(
                    "namedrange.delete",
                    new { name = _namedRanges[index] });
            }
            catch (Exception ex)
            {
                failures ??= [];
                failures.Add(ex);
            }
        }

        for (var index = _tables.Count - 1; index >= 0; index--)
        {
            try
            {
                await fixture.SendAsync(
                    "table.delete",
                    new { tableName = _tables[index] });
            }
            catch (Exception ex)
            {
                failures ??= [];
                failures.Add(ex);
            }
        }

        for (var index = _sheets.Count - 1; index >= 0; index--)
        {
            try
            {
                await fixture.SendAsync(
                    "sheetstyle.show",
                    new { sheetName = _sheets[index] });
                await fixture.SendAsync("sheet.delete", new { sheetName = _sheets[index] });
            }
            catch (Exception ex)
            {
                failures ??= [];
                failures.Add(ex);
            }
        }

        if (failures is not null)
        {
            throw new AggregateException("Service workbook cleanup failed.", failures);
        }
    }
}
