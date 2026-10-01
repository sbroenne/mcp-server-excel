using System.Diagnostics;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

/// <summary>
/// Integration tests for ExcelBatch - verifies batch operations and COM cleanup.
/// Tests that Excel instances are reused across operations and properly cleaned up.
///
/// LAYER RESPONSIBILITY:
/// - ✅ Test ExcelBatch.Execute() reuses Excel instance
/// - ✅ Test ExcelBatch.Dispose() COM cleanup
/// - ✅ Test ExcelBatch.Save() functionality
/// - ✅ Verify Excel.exe process termination (no leaks)
///
/// NOTE: ExcelBatch.Dispose() handles all GC cleanup automatically.
/// Tests only need to wait for async disposal and process termination timing.
///
/// These tests run sequentially and verify only their owned Excel processes.
/// Startup-hook and configured native-capability controls run OnDemand.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "ExcelBatch")]
[Collection("Sequential")] // Disable parallelization to avoid COM interference
[Trait("RequiresExcel", "true")]
public class ExcelBatchTests : IAsyncLifetime, IDisposable
{
    private readonly ITestOutputHelper _output;
    private static string? _staticTestFile;
    private string? _testFileCopy;
    private readonly OwnedExcelProcessScope _owned = new();
    private readonly List<string> _temporaryFiles = new();
    private readonly List<ExcelProcessIdentity> _startupIdentities = new();

    private static string? GetConfiguredIrmTestFilePath()
    {
        var irmTestFile = Environment.GetEnvironmentVariable("TEST_IRM_FILE");
        return !string.IsNullOrWhiteSpace(irmTestFile) && File.Exists(irmTestFile)
            ? Path.GetFullPath(irmTestFile)
            : null;
    }

    public ExcelBatchTests(ITestOutputHelper output)
    {
        _output = output;
    }

    public Task InitializeAsync()
    {
        // Use static test file from TestFiles folder (must be pre-created)
        if (_staticTestFile == null)
        {
            var testFolder = Path.Join(AppContext.BaseDirectory, "Integration", "Session", "TestFiles");
            _staticTestFile = Path.Join(testFolder, "batch-test-static.xlsx");

            // Verify the static file exists
            if (!File.Exists(_staticTestFile))
            {
                throw new FileNotFoundException($"Static test file not found at {_staticTestFile}. " +
                    "Please create the batch-test-static.xlsx file in the TestFiles folder.");
            }
        }

        // Create a fresh copy for this test instance (in temp folder)
        _testFileCopy = Path.Join(Path.GetTempPath(), $"batch-test-{Guid.NewGuid():N}.xlsx");
        _temporaryFiles.Add(_testFileCopy);
        File.Copy(_staticTestFile, _testFileCopy, overwrite: true);

        return Task.CompletedTask;
    }

    public Task DisposeAsync()
    {
        Dispose();
        return Task.CompletedTask;
    }

    public void Dispose()
    {
        GC.SuppressFinalize(this);
        var failures = new List<Exception>();
        foreach (var identity in _startupIdentities)
        {
            var failure = Record.Exception(() =>
            {
                Assert.True(ExcelBatch.TryTerminateOwnedProcess(identity,
                    TimeSpan.Zero, TimeSpan.FromSeconds(5)));
                Assert.True(OwnedProcessGuard.TryConfirmExited(identity));
                SessionManager.UntrackExcelProcess(identity);
            });
            if (failure is not null) { failures.Add(failure); }
        }
        var cleanup = Record.Exception(() =>
            SessionTestCleanup.AssertExitedAndDelete(_owned, _temporaryFiles));
        if (cleanup is not null) { failures.Add(cleanup); }
        if (failures.Count > 0)
        {
            throw new AggregateException("ExcelBatch test cleanup failed.", failures);
        }
    }

    [Fact]
    public void ExecuteAsync_MultipleOperations_ReusesExcelInstance()
    {
        // Arrange
        int operationCount = 0;

        // Act - Use batching for multiple operations
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);
        SessionWorkbookAssertions.AssertIdentity(batch, _testFileCopy!);
        var originalIdentity = ReadNativeIdentity(batch);
        var originalProcessId = batch.ExcelProcessId;
        SessionWorkbookAssertions.WriteMarker(batch, "initial");

        for (int i = 0; i < 5; i++)
        {
            batch.Execute((ctx, ct) =>
            {
                operationCount++;
                _output.WriteLine($"Batch operation {operationCount}");

                return operationCount;
            });
            Assert.Equal(originalIdentity, ReadNativeIdentity(batch));
            Assert.Equal(originalProcessId, batch.ExcelProcessId);
            Assert.Equal(i == 0 ? "initial" : $"operation-{i}",
                SessionWorkbookAssertions.ReadMarker(batch));
            SessionWorkbookAssertions.WriteMarker(batch, $"operation-{i + 1}");
        }

        // Assert
        Assert.Equal(5, operationCount);
        _output.WriteLine($"✓ Completed {operationCount} batch operations");
    }

    [Fact]
    public void Dispose_CleansUpComObjects_NoProcessLeak()
    {
        // Arrange
        using (var batch = ExcelSession.BeginBatch(_testFileCopy!))
        {
            SessionWorkbookAssertions.AssertIdentity(batch, _testFileCopy!);
            SessionWorkbookAssertions.WriteMarker(batch, "disposed-workbook");
        }

        _owned.AssertAllExited();
    }

    [Fact]
    public void Save_PersistsChanges_ToWorkbook()
    {
        // Arrange
        string testValue = $"Test-{Guid.NewGuid():N}";

        // Act - Write and save
        using (var batch = ExcelSession.BeginBatch(_testFileCopy!))
        {
            SessionWorkbookAssertions.AssertIdentity(batch, _testFileCopy!);
            SessionWorkbookAssertions.WriteMarker(batch, testValue);

            batch.Save();
        }

        _owned.AssertAllExited();

        // Verify - Read back the value in a new batch session
        string readValue;
        using (var batch = ExcelSession.BeginBatch(_testFileCopy!))
        {
            SessionWorkbookAssertions.AssertIdentity(batch, _testFileCopy!);
            readValue = Assert.IsType<string>(SessionWorkbookAssertions.ReadMarker(batch));
        }

        // Assert
        Assert.Equal(testValue, readValue);
        _output.WriteLine($"✓ Value persisted correctly: {testValue}");
    }

    [Fact]
    public void Save_ReadOnlyWorkbook_RejectsEvenWhenAlreadySaved()
    {
        using var batch = ExcelSession.BeginReadOnlyValidation(
            _testFileCopy!, TimeSpan.FromSeconds(30));
        var originalMarker = SessionWorkbookAssertions.ReadMarker(batch);
        Assert.True(batch.Execute((context, _) => context.Book.ReadOnly));
        Assert.True(batch.Execute((context, _) => context.Book.Saved));

        var exception = Assert.Throws<InvalidOperationException>(() => batch.Save());

        Assert.Contains("read-only", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(originalMarker, SessionWorkbookAssertions.ReadMarker(batch));
        Assert.True(batch.Execute((context, _) => context.Book.ReadOnly));
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    public void BeginBatch_ProtectedDetection_RequestsEditableAccess()
    {
        var syntheticProtectedPath = Path.Join(
            Path.GetTempPath(), $"batch-irm-access-{Guid.NewGuid():N}.xlsx");
        _temporaryFiles.Add(syntheticProtectedPath);
        OleDataSpaceTestFile.Write(syntheticProtectedPath, "\tDRMDataSpace");
        Assert.True(FileAccessValidator.IsIrmProtected(syntheticProtectedPath));
        bool reachedOpen = false;
        ExcelBatch.BeforeWorkbookOpenHook = (path, _) =>
        {
            Assert.Equal(syntheticProtectedPath, path, ignoreCase: true);
            reachedOpen = true;
            // Replace only synthetic metadata with a valid fixture after detection.
            // This exercises the actual COM open options without enterprise credentials.
            File.Copy(_testFileCopy!, path, overwrite: true);
        };

        try
        {
            using var batch = ExcelSession.BeginBatch(
                show: true, operationTimeout: TimeSpan.FromSeconds(30), syntheticProtectedPath);
            Assert.True(reachedOpen);
            Assert.True(batch.Execute((context, _) => context.App.Visible));
            Assert.False(batch.Execute((context, _) => context.Book.ReadOnly));
            Assert.Equal(
                Path.GetFullPath(syntheticProtectedPath),
                batch.Execute((context, _) => context.Book.FullName),
                ignoreCase: true);
        }
        finally
        {
            ExcelBatch.BeforeWorkbookOpenHook = null;
        }
    }

    [JapaneseLocaleFact]
    [Trait("RunType", "OnDemand")]
    [Trait("RequiresExcel", "true")]
    [Trait("Locale", "ja-JP")]
    public void OpenAndSave_JapaneseTableDateFormat_PreservesTableColumnDxf()
    {
        string workbookPath = Path.Join(Path.GetTempPath(), $"table-dxf-{Guid.NewGuid():N}.xlsx");
        _temporaryFiles.Add(workbookPath);
        CreateJapaneseTableWorkbook(workbookPath);
        string expectedFormat = ReadTableColumnFormatCode(workbookPath, "Date");
        Assert.Equal("yyyy/m/d", expectedFormat);

        using (var batch = ExcelSession.BeginBatch(workbookPath))
        {
            batch.Save();
        }

        Assert.Equal(expectedFormat, ReadTableColumnFormatCode(workbookPath, "Date"));
    }

    [Fact]
    public void WorkbookPath_ReturnsCorrectPath()
    {
        // Arrange & Act
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);

        // Assert
        SessionWorkbookAssertions.AssertIdentity(batch, _testFileCopy!);
    }

    [Fact]
    public void Execute_DependentOperations_RetainCreatedWorksheetAndNamedRange()
    {
        const string sheetName = "TestData";
        const string namedRangeName = "TestRange";
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);
        var identity = ReadNativeIdentity(batch);
        batch.Execute((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets.Add();
                sheet.Name = sheetName;
                Assert.Equal(sheetName, sheet.Name);
            }
            finally
            {
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        WithWorksheet(batch, sheetName, sheet =>
        {
            Excel.Range? values = null;
            Excel.Range? formula = null;
            try
            {
                values = sheet.Range["A1:B2"];
                values.Value2 = new object[,] { { "Header1", "Header2" }, { "Value1", 0 } };
                formula = sheet.Range["B2"];
                formula.Formula = "=LEN(A2)";
            }
            finally
            {
                ComUtilities.Release(ref formula);
                ComUtilities.Release(ref values);
            }
        });
        AssertWorkflowValues("Value1", 6);
        batch.Execute((context, _) =>
        {
            Excel.Names? names = null;
            Excel.Name? name = null;
            try
            {
                names = context.Book.Names;
                name = names.Add(namedRangeName, $"={sheetName}!$A$1:$B$2");
                Assert.Equal(namedRangeName, name.Name);
                Assert.Equal($"={sheetName}!$A$1:$B$2", name.RefersTo);
            }
            finally
            {
                ComUtilities.Release(ref name);
                ComUtilities.Release(ref names);
            }
        });
        WithWorksheet(batch, sheetName, sheet =>
        {
            Excel.Range? cell = null;
            try
            {
                cell = sheet.Range["A2"];
                cell.Value2 = "Modified";
            }
            finally { ComUtilities.Release(ref cell); }
        });
        AssertWorkflowValues("Modified", 8);
        batch.Execute((context, _) =>
        {
            Excel.Names? names = null;
            Excel.Name? name = null;
            Excel.Range? target = null;
            try
            {
                names = context.Book.Names;
                name = names.Item(namedRangeName);
                target = name.RefersToRange;
                Assert.Equal($"={sheetName}!$A$1:$B$2", name.RefersTo);
                Assert.Equal("$A$1:$B$2", target.Address);
                var values = Assert.IsType<object[,]>(target.Value2);
                Assert.Equal("Modified", values[2, 1]);
                Assert.Equal(8d, values[2, 2]);
            }
            finally
            {
                ComUtilities.Release(ref target);
                ComUtilities.Release(ref name);
                ComUtilities.Release(ref names);
            }
        });
        Assert.Equal(identity, ReadNativeIdentity(batch));

        void AssertWorkflowValues(string expectedValue, double expectedLength)
        {
            WithWorksheet(batch, sheetName, sheet =>
            {
                Excel.Range? range = null;
                try
                {
                    range = sheet.Range["A1:B2"];
                    var values = Assert.IsType<object[,]>(range.Value2);
                    Assert.Equal("Header1", values[1, 1]);
                    Assert.Equal("Header2", values[1, 2]);
                    Assert.Equal(expectedValue, values[2, 1]);
                    Assert.Equal(expectedLength, values[2, 2]);
                }
                finally { ComUtilities.Release(ref range); }
            });
        }
    }

    private static (int Window, IntPtr Workbook) ReadNativeIdentity(IExcelBatch batch) =>
        batch.Execute((context, _) =>
        {
            var workbook = Marshal.GetIUnknownForObject(context.Book);
            try { return (context.App.Hwnd, workbook); }
            finally { Marshal.Release(workbook); }
        });

    private static void WithWorksheet(IExcelBatch batch, object identifier, Action<Excel.Worksheet> action)
    {
        batch.Execute((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[identifier];
                action(sheet);
            }
            finally
            {
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }

    [Fact]
    public async Task ParallelBatches_TwoConcurrentBatches_NoExcelProcessLeak()
    {
        // Arrange
        const int batchCount = 2;
        var testFileCopies = new List<string>();

        // Create fresh copies for parallel test
        for (int i = 0; i < batchCount; i++)
        {
            string copy = Path.Join(Path.GetTempPath(), $"batch-test-parallel-{i}-{Guid.NewGuid():N}.xlsx");
            _temporaryFiles.Add(copy);
            File.Copy(_staticTestFile!, copy, overwrite: true);
            testFileCopies.Add(copy);
        }

        var tasks = testFileCopies.Select((testFile, index) =>
        {
            return Task.Run(() =>
            {
                using var batch = ExcelSession.BeginBatch(testFile);
                SessionWorkbookAssertions.AssertIdentity(batch, testFile);
                var identity = ReadNativeIdentity(batch);

                // Perform multiple operations per batch
                for (int op = 0; op < 3; op++)
                {
                    WithWorksheet(batch, 1, sheet =>
                    {
                        Excel.Range? cell = null;
                        try
                        {
                            cell = sheet.Range[$"A{op + 1}"];
                            cell.Value2 = $"Batch{index}-Op{op}";
                            Assert.Equal($"Batch{index}-Op{op}", cell.Value2);
                        }
                        finally { ComUtilities.Release(ref cell); }
                    });
                    Assert.Equal(identity, ReadNativeIdentity(batch));
                }
                WithWorksheet(batch, 1, sheet =>
                {
                    Excel.Range? cells = null;
                    try
                    {
                        cells = sheet.Range["A1:A3"];
                        var values = Assert.IsType<object[,]>(cells.Value2);
                        for (int row = 1; row <= 3; row++)
                        {
                            Assert.Equal($"Batch{index}-Op{row - 1}", values[row, 1]);
                        }
                    }
                    finally { ComUtilities.Release(ref cells); }
                });

                _output.WriteLine($"✓ Batch {index} completed");

                return (Index: index, ProcessId: batch.ExcelProcessId);
            });
        }).ToArray();

        // Wait for all batches to complete
        var results = await Task.WhenAll(tasks);

        Assert.Equal(batchCount, results.Length);
        Assert.Equal([0, 1], results.Select(result => result.Index).Order());
        Assert.All(results, result => Assert.NotNull(result.ProcessId));
        Assert.Equal(batchCount, results.Select(result => result.ProcessId).Distinct().Count());
        _output.WriteLine($"✓ All {batchCount} parallel batches completed");

        _owned.AssertAllExited();
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    public void BeginBatch_IrmWorkbook_ShowFalse_FailsFastBeforeOpen()
    {
        string fakeIrmFile = Path.Join(Path.GetTempPath(), $"batch-irm-headless-{Guid.NewGuid():N}.xlsx");
        _temporaryFiles.Add(fakeIrmFile);
        OleDataSpaceTestFile.Write(fakeIrmFile, "\tDRMDataSpace");
        var originalBytes = File.ReadAllBytes(fakeIrmFile);

        bool openAttempted = false;
        ExcelBatch.BeforeWorkbookOpenHook = (_, _) => openAttempted = true;

        try
        {
            var ex = Assert.Throws<InvalidOperationException>(() =>
                ExcelSession.BeginBatch(show: false, operationTimeout: TimeSpan.FromSeconds(15), fakeIrmFile));

            Assert.False(openAttempted);
            Assert.Contains("IRM/AIP-protected workbook", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("show=true", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(originalBytes, File.ReadAllBytes(fakeIrmFile));
        }
        finally
        {
            ExcelBatch.BeforeWorkbookOpenHook = null;

        }
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    [Trait("RequiresExcel", "true")]
    public void BeginBatch_StartupFailureWithConfirmedExit_UntracksExactIdentity()
    {
        ExcelProcessIdentity? capturedIdentity = null;
        void CaptureIdentity(ExcelProcessIdentity identity)
        {
            capturedIdentity = identity;
            _startupIdentities.Add(identity);
        }

        SessionManager.ExcelProcessIdentityTracked += CaptureIdentity;
        ExcelBatch.BeforeWorkbookOpenHook = (_, _) =>
            throw new InvalidOperationException("synthetic startup failure");
        ExcelBatch.FailedStartupTerminationHook = _ => false;
        ExcelBatch.FailedStartupExitConfirmationHook = _ => true;
        try
        {
            var exception = Assert.Throws<InvalidOperationException>(
                () => ExcelSession.BeginBatch(_testFileCopy!));

            Assert.Equal("synthetic startup failure", exception.Message);
            var identity = Assert.IsType<ExcelProcessIdentity>(capturedIdentity);
            Assert.DoesNotContain(identity, SessionManager.GetTrackedExcelProcesses());
        }
        finally
        {
            SessionManager.ExcelProcessIdentityTracked -= CaptureIdentity;
            ExcelBatch.BeforeWorkbookOpenHook = null;
            ExcelBatch.FailedStartupTerminationHook = null;
            ExcelBatch.FailedStartupExitConfirmationHook = null;
        }
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    [Trait("RequiresExcel", "true")]
    public void BeginBatch_GenericExcel1004_PreservesOriginalDiagnostic()
    {
#pragma warning disable CA2201 // A real COMException is required to exercise Excel HRESULT classification.
        ExcelBatch.BeforeWorkbookOpenHook = (_, _) =>
            throw new COMException("Synthetic non-lock Excel failure", unchecked((int)0x800A03EC));
#pragma warning restore CA2201

        try
        {
            var exception = Assert.Throws<COMException>(
                () => ExcelSession.BeginBatch(_testFileCopy!));

            Assert.Equal(unchecked((int)0x800A03EC), exception.HResult);
            Assert.Contains("Synthetic non-lock Excel failure", exception.Message, StringComparison.Ordinal);
            Assert.DoesNotContain("locked by another process", exception.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            ExcelBatch.BeforeWorkbookOpenHook = null;
        }
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    [Trait("RequiresExcel", "true")]
    public void BeginBatch_StartupFailureWithUnconfirmedLiveIdentity_RetainsOwnership()
    {
        ExcelProcessIdentity? capturedIdentity = null;
        void CaptureIdentity(ExcelProcessIdentity identity)
        {
            capturedIdentity = identity;
            _startupIdentities.Add(identity);
        }

        SessionManager.ExcelProcessIdentityTracked += CaptureIdentity;
        ExcelBatch.BeforeWorkbookOpenHook = (_, _) =>
            throw new InvalidOperationException("synthetic startup failure");
        ExcelBatch.FailedStartupTerminationHook = _ => false;
        ExcelBatch.FailedStartupExitConfirmationHook = _ => false;
        try
        {
            var exception = Assert.Throws<InvalidOperationException>(
                () => ExcelSession.BeginBatch(_testFileCopy!));

            var identity = Assert.IsType<ExcelProcessIdentity>(capturedIdentity);
            Assert.Contains(identity, SessionManager.GetTrackedExcelProcesses());
            Assert.Contains("remains tracked", exception.Message, StringComparison.OrdinalIgnoreCase);
            var failures = Assert.IsType<AggregateException>(exception.InnerException);
            Assert.Contains(
                failures.InnerExceptions,
                failure => failure.Message == "synthetic startup failure");
        }
        finally
        {
            SessionManager.ExcelProcessIdentityTracked -= CaptureIdentity;
            ExcelBatch.BeforeWorkbookOpenHook = null;
            ExcelBatch.FailedStartupTerminationHook = null;
            ExcelBatch.FailedStartupExitConfirmationHook = null;
        }
    }

    [ConfiguredIrmFact]
    [Trait("RunType", "OnDemand")]
    public async Task BeginBatch_RealIrmWorkbook_CompletesStartupWithinBudget_WhenConfigured()
    {
        // Real IRM startup depends on interactive auth/enterprise policy and cannot run in CI.
        var irmTestFile = GetConfiguredIrmTestFilePath()
            ?? throw new InvalidOperationException("Configured IRM test fixture was unavailable after test discovery.");

        using var owned = new OwnedExcelProcessScope();

        var stopwatch = Stopwatch.StartNew();
        IExcelBatch? batch = null;
        var openTask = Task.Run(() =>
            ExcelSession.BeginBatch(show: true, operationTimeout: TimeSpan.FromSeconds(15), irmTestFile));
        var primary = await Record.ExceptionAsync(async () =>
        {
            batch = await openTask.WaitAsync(TimeSpan.FromSeconds(20));
            stopwatch.Stop();
            SessionWorkbookAssertions.AssertIdentity(batch, irmTestFile);
            Assert.True(batch.Execute((context, _) => context.App.Visible));
            Assert.True(stopwatch.Elapsed <= TimeSpan.FromSeconds(20));
            _output.WriteLine($"Opened IRM workbook in {stopwatch.Elapsed.TotalSeconds:F1}s");
        });
        var completion = await Record.ExceptionAsync(async () =>
        {
            if (batch is null && !openTask.IsFaulted && !openTask.IsCanceled)
            {
                batch = await openTask.WaitAsync(TimeSpan.FromSeconds(60));
            }
        });
        var disposal = Record.Exception(() => batch?.Dispose());
        var exit = Record.Exception(() => owned.AssertAllExited());
        var failures = new[] { primary, completion, disposal, exit }.OfType<Exception>().ToArray();
        if (failures.Length > 0)
        {
            throw new AggregateException("Configured IRM startup or cleanup failed.", failures);
        }
    }

    private static void CreateJapaneseTableWorkbook(string workbookPath)
    {
        Excel.Application? excel = null;
        Excel.Workbooks? workbooks = null;
        Excel.Workbook? workbook = null;
        Excel.Sheets? worksheets = null;
        Excel.Worksheet? worksheet = null;
        Excel.Range? sourceRange = null;
        Excel.ListObjects? listObjects = null;
        Excel.ListObject? table = null;
        Excel.ListColumns? listColumns = null;
        Excel.ListColumn? dateColumn = null;
        Excel.Range? dataBodyRange = null;
        Exception? primary = null;
        Exception? cleanup = null;

        try
        {
            primary = Record.Exception(() =>
            {
                excel = new Excel.Application { DisplayAlerts = false };
                workbooks = excel.Workbooks;
                workbook = workbooks.Add();
                worksheets = workbook.Worksheets;
                worksheet = (Excel.Worksheet)worksheets[1];
                sourceRange = worksheet.Range["A1:B3"];
                sourceRange.Value2 = new object[,] { { "Amount", "Date" }, { 1, 46000 }, { 2, 46001 } };
                listObjects = worksheet.ListObjects;
                table = listObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, sourceRange,
                    Type.Missing, Excel.XlYesNoGuess.xlYes);
                listColumns = table.ListColumns;
                dateColumn = listColumns["Date"];
                dataBodyRange = dateColumn.DataBodyRange;
                dataBodyRange.NumberFormatLocal = "yyyy/m/d";
                workbook.SaveAs(workbookPath, Excel.XlFileFormat.xlOpenXMLWorkbook);
            });
            cleanup = Record.Exception(() => CloseWorkbookAndQuitExcel(workbook, excel));
        }
        finally
        {
            ComUtilities.Release(ref dataBodyRange);
            ComUtilities.Release(ref dateColumn);
            ComUtilities.Release(ref listColumns);
            ComUtilities.Release(ref table);
            ComUtilities.Release(ref listObjects);
            ComUtilities.Release(ref sourceRange);
            ComUtilities.Release(ref worksheet);
            ComUtilities.Release(ref worksheets);
            ComUtilities.Release(ref workbook);
            ComUtilities.Release(ref workbooks);
            ComUtilities.Release(ref excel);
        }
        var failures = new[] { primary, cleanup }.OfType<Exception>().ToArray();
        if (failures.Length > 0)
        {
            throw new AggregateException("Japanese native fixture creation or cleanup failed.", failures);
        }
    }

    private static string ReadTableColumnFormatCode(string workbookPath, string columnName)
    {
        using var archive = ZipFile.OpenRead(workbookPath);
        var tableEntry = archive.Entries.Single(entry => entry.FullName.StartsWith("xl/tables/", StringComparison.OrdinalIgnoreCase));
        using var tableStream = tableEntry.Open();
        var tableDocument = XDocument.Load(tableStream);
        XNamespace spreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var column = tableDocument.Descendants(spreadsheetNamespace + "tableColumn")
            .Single(element => element.Attribute("name")?.Value == columnName);
        int dxfId = int.Parse(column.Attribute("dataDxfId")!.Value, CultureInfo.InvariantCulture);

        var stylesEntry = archive.GetEntry("xl/styles.xml")
            ?? throw new InvalidOperationException("Workbook styles were not found.");
        using var stylesStream = stylesEntry.Open();
        var stylesDocument = XDocument.Load(stylesStream);
        return stylesDocument.Root!
            .Element(spreadsheetNamespace + "dxfs")!
            .Elements(spreadsheetNamespace + "dxf")
            .ElementAt(dxfId)
            .Element(spreadsheetNamespace + "numFmt")!
            .Attribute("formatCode")!
            .Value;
    }

    private static void CloseWorkbookAndQuitExcel(Excel.Workbook? workbook, Excel.Application? excel)
    {
        var failures = new List<Exception>();
        if (workbook != null)
        {
            var failure = Record.Exception(() => workbook.Close(false));
            if (failure is not null) { failures.Add(failure); }
        }

        if (excel != null)
        {
            var failure = Record.Exception(excel.Quit);
            if (failure is not null) { failures.Add(failure); }
        }
        if (failures.Count > 0)
        {
            throw new AggregateException("Native Excel fixture shutdown failed.", failures);
        }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "FileLocking")]
    public void Constructor_FileLockedByAnotherProcess_ThrowsInvalidOperationException()
    {
        // Arrange - Create a separate test file for locking test
        var lockedTestFile = Path.Join(Path.GetTempPath(), $"batch-test-locked-{Guid.NewGuid():N}.xlsx");
        _temporaryFiles.Add(lockedTestFile);
        File.Copy(_staticTestFile!, lockedTestFile, overwrite: true);
        var originalBytes = File.ReadAllBytes(lockedTestFile);

        {
            // Lock the file by opening with exclusive access (simulating Excel or another process)
            using var fileLock = new FileStream(
                lockedTestFile,
                FileMode.Open,
                FileAccess.ReadWrite,
                FileShare.None);

            // Act & Assert - Attempting to create ExcelBatch should fail immediately
            var ex = Assert.Throws<InvalidOperationException>(() =>
            {
                var batch = ExcelSession.BeginBatch(lockedTestFile);
                batch.Dispose();
            });

            // Verify error message is clear and actionable
            Assert.Contains("already open", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("close the file", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("exclusive access", ex.Message, StringComparison.OrdinalIgnoreCase);

            _output.WriteLine($"✓ File locking detected successfully");
            _output.WriteLine($"Error message: {ex.Message}");
        }
        Assert.Equal(originalBytes, File.ReadAllBytes(lockedTestFile));
        using var recovered = ExcelSession.BeginBatch(lockedTestFile);
        SessionWorkbookAssertions.AssertIdentity(recovered, lockedTestFile);
        SessionWorkbookAssertions.WriteMarker(recovered, "recovered-after-file-lock");
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "FileLocking")]
    public void Constructor_FileLockedByAnotherProcess_DoesNotLeakExcelProcess()
    {
        var lockedTestFile = Path.Join(Path.GetTempPath(), $"batch-test-locked-leak-{Guid.NewGuid():N}.xlsx");
        _temporaryFiles.Add(lockedTestFile);
        File.Copy(_staticTestFile!, lockedTestFile, overwrite: true);
        var originalBytes = File.ReadAllBytes(lockedTestFile);

        using var owned = new OwnedExcelProcessScope();

        {
            using var fileLock = new FileStream(
                lockedTestFile,
                FileMode.Open,
                FileAccess.ReadWrite,
                FileShare.None);

            var ex = Assert.Throws<InvalidOperationException>(() =>
            {
                using var batch = ExcelSession.BeginBatch(lockedTestFile);
            });

            Assert.Contains("already open", ex.Message, StringComparison.OrdinalIgnoreCase);

            owned.AssertAllExited(expectProcess: false);
        }
        Assert.Equal(originalBytes, File.ReadAllBytes(lockedTestFile));
    }
}
