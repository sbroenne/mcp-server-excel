using System.Collections.Concurrent;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Tests.Infrastructure;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceSheetTests
{
    private static readonly ConcurrentDictionary<(string? Sheet, bool DuplicateGuard), SavedWorkbookTemplateStore>
        MarkedWorkbookTemplates = new();

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, true)]
    [InlineData(true, false)]
    public void CrossFile_CancelledSave_ReportsExactPersistedState(bool move, bool cancelSource)
    {
        var (sourceFile, targetFile) = CreateMarkedWorkbookPair(
            nameof(CrossFile_CancelledSave_ReportsExactPersistedState), "TransferSheet");
        var sourceBytes = File.ReadAllBytes(sourceFile);
        var targetBytes = File.ReadAllBytes(targetFile);
        var cancelledPath = cancelSource ? sourceFile : targetFile;
        var eventCount = 0;
        var originalHook = ExcelBatch.AfterWorkbookOpenHookForTests;
        try
        {
            ExcelBatch.AfterWorkbookOpenHookForTests = (application, openedWorkbook) =>
            {
                var app = (Excel.Application)application;
                var workbook = (Excel.Workbook)openedWorkbook;
                if (!string.Equals(workbook.FullName, targetFile, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                Excel.AppEvents_WorkbookBeforeSaveEventHandler? cancelSave = null;
                cancelSave = (Excel.Workbook savingWorkbook, bool _, ref bool cancel) =>
                {
                    if (string.Equals(savingWorkbook.FullName, cancelledPath, StringComparison.OrdinalIgnoreCase))
                    {
                        eventCount++;
                        cancel = true;
                        app.WorkbookBeforeSave -= cancelSave;
                    }
                };
                app.WorkbookBeforeSave += cancelSave;
            };
            var sourceSaved = move && !cancelSource;
            var command = move ? "sheet.move-to-file" : "sheet.copy-to-file";
            RunRetainedCrossFileFailure(command,
                new { sourceFile, sourceSheet = "TransferSheet", targetFile },
                (service, response, sessionId) =>
                {
                    ExcelBatch.AfterWorkbookOpenHookForTests = originalHook;
                    Assert.Equal(1, eventCount);
                    Assert.Contains("save was not confirmed", response.ErrorMessage, StringComparison.Ordinal);
                    Assert.Contains("was not saved", response.InnerError, StringComparison.Ordinal);
                    Assert.Contains(Path.GetFileName(cancelledPath), response.InnerError, StringComparison.Ordinal);
                    if (move)
                    {
                        Assert.Contains(
                            cancelSource ? "source save was not confirmed" : "Excel confirmed the source save",
                            response.ErrorMessage, StringComparison.Ordinal);
                        Assert.Contains("No rollback was performed", response.ErrorMessage, StringComparison.Ordinal);
                    }
                    AssertRetainedCrossFileState(service, response, sessionId, command,
                        sourceFile, move ? ["Sheet1"] : ["TransferSheet", "Sheet1"],
                        sourceSaved: !move || sourceSaved,
                        targetFile, ["Sheet1", "TransferSheet"], targetSaved: false);
                });
            AssertWorkbookStates(
                sourceFile, sourceSaved ? ["Sheet1"] : ["TransferSheet", "Sheet1"],
                sourceSaved ? null : "TransferSheet",
                targetFile, ["Sheet1"]);
            if (!sourceSaved)
            {
                Assert.Equal(sourceBytes, File.ReadAllBytes(sourceFile));
            }
            else
            {
                Assert.NotEqual(sourceBytes, File.ReadAllBytes(sourceFile));
            }
            Assert.Equal(targetBytes, File.ReadAllBytes(targetFile));
        }
        finally
        {
            ExcelBatch.AfterWorkbookOpenHookForTests = originalHook;
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, true)]
    [InlineData(true, false)]
    public void CrossFile_ReadOnlyWorkbook_RejectsBeforeChangingEitherFile(
        bool move, bool readOnlySource)
    {
        var (sourceFile, targetFile) = CreateMarkedWorkbookPair(
            nameof(CrossFile_ReadOnlyWorkbook_RejectsBeforeChangingEitherFile),
            "TransferSheet");
        var sourceBytes = File.ReadAllBytes(sourceFile);
        var targetBytes = File.ReadAllBytes(targetFile);
        var readOnlyFile = readOnlySource ? sourceFile : targetFile;
        var originalAttributes = File.GetAttributes(readOnlyFile);
        var originalOpenHook = ExcelBatch.BeforeWorkbookOpenHook;
        Exception? failure = null;
        try
        {
            // Change permissions after filesystem preflight so Excel determines read-only access.
            ExcelBatch.BeforeWorkbookOpenHook = (path, _) =>
            {
                if (string.Equals(path, readOnlyFile, StringComparison.OrdinalIgnoreCase))
                {
                    File.SetAttributes(readOnlyFile, originalAttributes | FileAttributes.ReadOnly);
                }
            };
            using (var batch = ExcelSession.BeginBatch(sourceFile, targetFile))
            {
                batch.Execute((_, _) =>
                {
                    Assert.Equal(readOnlySource, batch.GetWorkbook(sourceFile).ReadOnly);
                    Assert.Equal(!readOnlySource, batch.GetWorkbook(targetFile).ReadOnly);
                    Assert.True(batch.GetWorkbook(sourceFile).Saved);
                    Assert.True(batch.GetWorkbook(targetFile).Saved);
                });
            }
            File.SetAttributes(readOnlyFile, originalAttributes);

            var command = move ? "sheet.move-to-file" : "sheet.copy-to-file";
            RunRetainedCrossFileFailure(command,
                new { sourceFile, sourceSheet = "TransferSheet", targetFile },
                (service, response, sessionId) =>
                {
                    ExcelBatch.BeforeWorkbookOpenHook = originalOpenHook;
                    File.SetAttributes(readOnlyFile, originalAttributes);
                    Assert.Contains("Cannot change this workbook", response.ErrorMessage, StringComparison.Ordinal);
                    Assert.Contains("read-only", response.ErrorMessage, StringComparison.Ordinal);
                    Assert.Contains(
                        "This operation has not changed the workbook.", response.ErrorMessage, StringComparison.Ordinal);
                    AssertRetainedCrossFileState(service, response, sessionId, command,
                        sourceFile, ["TransferSheet", "Sheet1"], sourceSaved: true,
                        targetFile, ["Sheet1"], targetSaved: true,
                        sourceReadOnly: readOnlySource, targetReadOnly: !readOnlySource);
                });
            AssertWorkbookState(sourceFile, ["TransferSheet", "Sheet1"], "TransferSheet");
            AssertWorkbookState(targetFile, ["Sheet1"]);
            Assert.Equal(sourceBytes, File.ReadAllBytes(sourceFile));
            Assert.Equal(targetBytes, File.ReadAllBytes(targetFile));
        }
        catch (Exception ex)
        {
            failure = ex;
        }
        finally
        {
            ExcelBatch.BeforeWorkbookOpenHook = originalOpenHook;
            try
            {
                File.SetAttributes(readOnlyFile, originalAttributes);
            }
            catch (Exception ex)
            {
                failure = PersistentServiceCleanupFailures.Combine(failure, ex);
            }
        }
        PersistentServiceCleanupFailures.ThrowIfFailed(failure);
    }

    [Fact]
    public void CopyToFile_WithTargetName_CopiesAndRenames()
    {
        var (sourceFile, targetFile) = CreateMarkedWorkbookPair(
            nameof(CopyToFile_WithTargetName_CopiesAndRenames), "SourceSheet");

        RequireSuccess(_sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile, "CopiedSheet"));

        AssertWorkbookStates(
            sourceFile, ["SourceSheet", "Sheet1"], "SourceSheet",
            targetFile, ["Sheet1", "CopiedSheet"], "CopiedSheet");
    }

    [Fact]
    public void CopyToFile_NoTargetName_CopiesWithOriginalName()
    {
        var (sourceFile, targetFile) = CreateMarkedWorkbookPair(
            nameof(CopyToFile_NoTargetName_CopiesWithOriginalName), "SourceSheet");

        RequireSuccess(_sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile));

        AssertWorkbookStates(
            sourceFile, ["SourceSheet", "Sheet1"], "SourceSheet",
            targetFile, ["Sheet1", "SourceSheet"], "SourceSheet");
    }

    [Fact]
    public void CopyToFile_WithBeforeSheet_PositionsCorrectly()
    {
        var (sourceFile, targetFile) = CreateMarkedWorkbookPair(
            nameof(CopyToFile_WithBeforeSheet_PositionsCorrectly), "SourceSheet");

        RequireSuccess(_sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile, "Copied", beforeSheet: "Sheet1"));

        AssertWorkbookStates(
            sourceFile, ["SourceSheet", "Sheet1"], "SourceSheet",
            targetFile, ["Copied", "Sheet1"], "Copied");
    }

    [Fact]
    public void CopyToFile_WithAfterSheet_PositionsCorrectly()
    {
        var (sourceFile, targetFile) = CreateMarkedWorkbookPair(
            nameof(CopyToFile_WithAfterSheet_PositionsCorrectly), "SourceSheet");

        RequireSuccess(_sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile, "Copied", afterSheet: "Sheet1"));

        AssertWorkbookStates(
            sourceFile, ["SourceSheet", "Sheet1"], "SourceSheet",
            targetFile, ["Sheet1", "Copied"], "Copied");
    }

    [Fact]
    public void CopyToFile_BothBeforeAndAfter_ThrowsException()
    {
        var (sourceFile, targetFile) = CreateMarkedWorkbookPair(
            nameof(CopyToFile_BothBeforeAndAfter_ThrowsException), "SourceSheet");
        var sourceBytes = File.ReadAllBytes(sourceFile);
        var targetBytes = File.ReadAllBytes(targetFile);

        var exception = Assert.Throws<ArgumentException>(() =>
            _sheetCommands.CopyToFile(
                sourceFile,
                "SourceSheet",
                targetFile,
                "Copied",
                beforeSheet: "Sheet1",
                afterSheet: "Sheet1"));
        Assert.Contains(
            "both beforeSheet and afterSheet",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
        AssertWorkbookStates(
            sourceFile, ["SourceSheet", "Sheet1"], "SourceSheet",
            targetFile, ["Sheet1"]);
        Assert.Equal(sourceBytes, File.ReadAllBytes(sourceFile));
        Assert.Equal(targetBytes, File.ReadAllBytes(targetFile));
    }

    [Fact]
    public void CopyToFile_SameFile_ThrowsException()
    {
        var testFile = CreateMarkedWorkbook(nameof(CopyToFile_SameFile_ThrowsException), "Test");
        var before = File.ReadAllBytes(testFile);

        var exception = Assert.Throws<ArgumentException>(() =>
            _sheetCommands.CopyToFile(testFile, "Sheet1", testFile, "Copied"));
        Assert.Contains("same-file copy", exception.Message, StringComparison.OrdinalIgnoreCase);
        AssertWorkbookState(testFile, ["Sheet1"]);
        Assert.Equal(before, File.ReadAllBytes(testFile));
    }

    [Fact]
    public void CopyToFile_SourceFileNotFound_ThrowsException()
    {
        var targetFile = CreateMarkedWorkbook(nameof(CopyToFile_SourceFileNotFound_ThrowsException), "Target");
        var before = File.ReadAllBytes(targetFile);
        var nonExistentSource = Path.Combine(Path.GetDirectoryName(targetFile)!, "NonExistent.xlsx");

        var exception = Assert.Throws<FileNotFoundException>(() =>
            _sheetCommands.CopyToFile(nonExistentSource, "Sheet1", targetFile));
        Assert.Contains("Source file not found", exception.Message);
        AssertWorkbookState(targetFile, ["Sheet1"]);
        Assert.Equal(before, File.ReadAllBytes(targetFile));
        Assert.False(File.Exists(nonExistentSource));
    }

    [Fact]
    public void CopyToFile_TargetFileNotFound_ThrowsException()
    {
        var sourceFile = CreateMarkedWorkbook(nameof(CopyToFile_TargetFileNotFound_ThrowsException), "Source");
        var before = File.ReadAllBytes(sourceFile);
        var nonExistentTarget = Path.Combine(Path.GetDirectoryName(sourceFile)!, "NonExistent.xlsx");

        var exception = Assert.Throws<FileNotFoundException>(() =>
            _sheetCommands.CopyToFile(sourceFile, "Sheet1", nonExistentTarget));
        Assert.Contains("Target file not found", exception.Message);
        AssertWorkbookState(sourceFile, ["Sheet1"]);
        Assert.Equal(before, File.ReadAllBytes(sourceFile));
        Assert.False(File.Exists(nonExistentTarget));
    }

    [Theory]
    [InlineData("History")]
    [InlineData("Bad/Name")]
    [InlineData("12345678901234567890123456789012")]
    [InlineData("   ")]
    public void CopyToFile_InvalidTargetSheetName_RejectsBeforeChangingEitherFile(string invalidName)
    {
        var (sourceFile, targetFile) = CreateMarkedWorkbookPair(
            nameof(CopyToFile_InvalidTargetSheetName_RejectsBeforeChangingEitherFile), "SourceSheet");
        var sourceBytes = File.ReadAllBytes(sourceFile);
        var targetBytes = File.ReadAllBytes(targetFile);

        var error = Assert.Throws<ArgumentException>(() =>
            _sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile, invalidName));

        Assert.Contains("Worksheet names must", error.Message, StringComparison.Ordinal);
        AssertWorkbookStates(
            sourceFile, ["SourceSheet", "Sheet1"], "SourceSheet",
            targetFile, ["Sheet1"]);
        Assert.Equal(sourceBytes, File.ReadAllBytes(sourceFile));
        Assert.Equal(targetBytes, File.ReadAllBytes(targetFile));
    }

    [Fact]
    public void CopyToFile_ExistingTargetSheetName_RejectsWithoutChangingEitherFile()
    {
        var (sourceFile, targetFile) = CreateMarkedWorkbookPair(
            nameof(CopyToFile_ExistingTargetSheetName_RejectsWithoutChangingEitherFile),
            "SourceSheet", "CopiedSheet", targetDuplicateGuard: true);
        var sourceBytes = File.ReadAllBytes(sourceFile);
        var targetBytes = File.ReadAllBytes(targetFile);

        RunRetainedCrossFileFailure("sheet.copy-to-file",
            new { sourceFile, sourceSheet = "SourceSheet", targetFile, targetSheetName = "CopiedSheet" },
            (service, response, sessionId) =>
            {
                Assert.Contains("already exists", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
                AssertRetainedCrossFileState(service, response, sessionId, "sheet.copy-to-file",
                    sourceFile, ["SourceSheet", "Sheet1"], sourceSaved: true,
                    targetFile, ["CopiedSheet", "Sheet1"], targetSaved: true, targetDuplicateGuard: true);
            });
        AssertWorkbookStates(
            sourceFile, ["SourceSheet", "Sheet1"], "SourceSheet",
            targetFile, ["CopiedSheet", "Sheet1"], "CopiedSheet", targetDuplicateGuard: true);
        Assert.Equal(sourceBytes, File.ReadAllBytes(sourceFile));
        Assert.Equal(targetBytes, File.ReadAllBytes(targetFile));
    }

    [Fact]
    public void MoveToFile_Default_MovesSheetSuccessfully()
    {
        var (sourceFile, targetFile) = CreateMarkedWorkbookPair(
            nameof(MoveToFile_Default_MovesSheetSuccessfully), "MoveMe");

        RequireSuccess(_sheetCommands.MoveToFile(sourceFile, "MoveMe", targetFile));

        AssertWorkbookStates(
            sourceFile, ["Sheet1"], null,
            targetFile, ["Sheet1", "MoveMe"], "MoveMe");
    }

    [Fact]
    public void MoveToFile_WithBeforeSheet_PositionsCorrectly()
    {
        var (sourceFile, targetFile) = CreateMarkedWorkbookPair(
            nameof(MoveToFile_WithBeforeSheet_PositionsCorrectly), "MoveMe");

        RequireSuccess(_sheetCommands.MoveToFile(sourceFile, "MoveMe", targetFile, beforeSheet: "Sheet1"));

        AssertWorkbookStates(
            sourceFile, ["Sheet1"], null,
            targetFile, ["MoveMe", "Sheet1"], "MoveMe");
    }

    [Fact]
    public void MoveToFile_WithAfterSheet_PositionsCorrectly()
    {
        var (sourceFile, targetFile) = CreateMarkedWorkbookPair(
            nameof(MoveToFile_WithAfterSheet_PositionsCorrectly), "MoveMe");

        RequireSuccess(_sheetCommands.MoveToFile(sourceFile, "MoveMe", targetFile, afterSheet: "Sheet1"));

        AssertWorkbookStates(
            sourceFile, ["Sheet1"], null,
            targetFile, ["Sheet1", "MoveMe"], "MoveMe");
    }

    [Fact]
    public void MoveToFile_BothBeforeAndAfter_ThrowsException()
    {
        var (sourceFile, targetFile) = CreateMarkedWorkbookPair(
            nameof(MoveToFile_BothBeforeAndAfter_ThrowsException), "MoveMe");
        var sourceBytes = File.ReadAllBytes(sourceFile);
        var targetBytes = File.ReadAllBytes(targetFile);

        var exception = Assert.Throws<ArgumentException>(() =>
            _sheetCommands.MoveToFile(
                sourceFile,
                "MoveMe",
                targetFile,
                beforeSheet: "Sheet1",
                afterSheet: "Sheet1"));
        Assert.Contains(
            "both beforeSheet and afterSheet",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
        AssertWorkbookStates(
            sourceFile, ["MoveMe", "Sheet1"], "MoveMe",
            targetFile, ["Sheet1"]);
        Assert.Equal(sourceBytes, File.ReadAllBytes(sourceFile));
        Assert.Equal(targetBytes, File.ReadAllBytes(targetFile));
    }

    [Fact]
    public void MoveToFile_SameFile_ThrowsException()
    {
        var testFile = CreateMarkedWorkbook(nameof(MoveToFile_SameFile_ThrowsException), "Test");
        var before = File.ReadAllBytes(testFile);

        var exception = Assert.Throws<ArgumentException>(() =>
            _sheetCommands.MoveToFile(testFile, "Sheet1", testFile));
        Assert.Contains("same-file move", exception.Message, StringComparison.OrdinalIgnoreCase);
        AssertWorkbookState(testFile, ["Sheet1"]);
        Assert.Equal(before, File.ReadAllBytes(testFile));
    }

    private static void RunRetainedCrossFileFailure(
        string command, object arguments, Action<ExcelMcpService, ServiceResponse, string> assertions)
    {
        // The shared fixture proxy always selects its own workbook and throws away
        // failure responses. This owner can inspect and close only the reported batch.
        var service = new ExcelMcpService();
        ServiceResponse? response = null;
        Exception? failure = null;
        try
        {
            response = service.ProcessAsync(new ServiceRequest
            {
                Command = command,
                Args = JsonSerializer.Serialize(arguments, ServiceProtocol.JsonOptions)
            }).GetAwaiter().GetResult();
            Assert.False(response.Success);
            Assert.Equal(nameof(InvalidOperationException), response.ExceptionType);
            Assert.Null(response.ErrorCategory);
            Assert.Null(response.HResult);
            Assert.Equal(command, response.Command);
            Assert.Null(response.Result);
            Assert.False(string.IsNullOrWhiteSpace(response.SessionId));
            assertions(service, response, response.SessionId!);
        }
        catch (Exception ex)
        {
            failure = ex;
        }
        finally
        {
            if (response?.SessionId is { } sessionId)
            {
                try
                {
                    var closed = service.ProcessAsync(new ServiceRequest
                    {
                        Command = "session.close",
                        SessionId = sessionId,
                        Args = JsonSerializer.Serialize(new { save = false }, ServiceProtocol.JsonOptions)
                    }).GetAwaiter().GetResult();
                    Assert.True(closed.Success, closed.ErrorMessage);
                    using var listed = ReadCrossFileServiceResult(service, "session.list", sessionId: null);
                    Assert.Empty(listed.RootElement.GetProperty("sessions").EnumerateArray());
                }
                catch (Exception ex)
                {
                    failure = PersistentServiceCleanupFailures.Combine(failure, ex);
                }
            }
            try
            {
                service.Dispose();
            }
            catch (Exception ex)
            {
                failure = PersistentServiceCleanupFailures.Combine(failure, ex);
            }
        }
        PersistentServiceCleanupFailures.ThrowIfFailed(failure);
    }

    private static void AssertRetainedCrossFileState(
        ExcelMcpService service, ServiceResponse response, string sessionId, string command,
        string sourceFile, string[] sourceSheets, bool sourceSaved,
        string targetFile, string[] targetSheets, bool targetSaved,
        bool sourceReadOnly = false, bool targetReadOnly = false, bool targetDuplicateGuard = false)
    {
        using (var partial = JsonDocument.Parse(response.PartialState!))
        {
            Assert.True(partial.RootElement.GetProperty("retained").GetBoolean());
            Assert.Equal(sessionId, partial.RootElement.GetProperty("sessionId").GetString());
            Assert.Equal(new[] { sourceFile, targetFile },
                partial.RootElement.GetProperty("filePaths").EnumerateArray().Select(path => path.GetString()));
            var operation = partial.RootElement.GetProperty("operation");
            Assert.Equal(command, operation.GetProperty("command").GetString());
            Assert.Equal("failed", operation.GetProperty("state").GetString());
            Assert.False(operation.GetProperty("executionPending").GetBoolean());
            Assert.False(operation.GetProperty("waitExpired").GetBoolean());
            Assert.False(string.IsNullOrWhiteSpace(operation.GetProperty("errorMessage").GetString()));
            Assert.Contains(operation.GetProperty("errorMessage").GetString()!, response.ErrorMessage,
                StringComparison.Ordinal);
        }
        using (var list = ReadCrossFileServiceResult(service, "session.list", sessionId: null))
        {
            var retained = Assert.Single(list.RootElement.GetProperty("sessions").EnumerateArray());
            Assert.Equal(sessionId, retained.GetProperty("sessionId").GetString());
            Assert.Equal(sourceFile, retained.GetProperty("filePath").GetString());
            Assert.Equal(new[] { sourceFile, targetFile },
                retained.GetProperty("filePaths").EnumerateArray().Select(path => path.GetString()));
            Assert.False(retained.GetProperty("isAttached").GetBoolean());
            Assert.False(retained.GetProperty("operationPending").GetBoolean());
        }
        foreach (var (path, expectedSheets) in new[] { (sourceFile, sourceSheets), (targetFile, targetSheets) })
        {
            using var sheets = ReadCrossFileServiceResult(service, "sheet.list", sessionId, new { filePath = path });
            Assert.Equal(path, sheets.RootElement.GetProperty("filePath").GetString());
            Assert.Equal(expectedSheets,
                sheets.RootElement.GetProperty("worksheets").EnumerateArray()
                    .Select(sheet => sheet.GetProperty("name").GetString()));
        }
        using (var info = ReadCrossFileServiceResult(service, "workbook.get-info", sessionId))
        {
            Assert.Equal(sourceFile, info.RootElement.GetProperty("filePath").GetString());
            Assert.Equal(sourceFile, info.RootElement.GetProperty("fullName").GetString());
            Assert.Equal(sourceSaved, info.RootElement.GetProperty("saved").GetBoolean());
            Assert.Equal(sourceReadOnly, info.RootElement.GetProperty("readOnly").GetBoolean());
        }
        AssertCrossFileServiceCells(service, sessionId, "Sheet1",
            "retained workbook marker", 71, 79, "=B1+8");
        foreach (var sheet in sourceSheets.Where(name => name != "Sheet1"))
            AssertCrossFileServiceCells(service, sessionId, sheet, "transferred marker", 3, 6, "=B1*2");
        using (var discovery = ReadCrossFileServiceResult(service, "session.discover", sessionId: null))
        {
            Assert.True(discovery.RootElement.GetProperty("complete").GetBoolean());
            Assert.Empty(discovery.RootElement.GetProperty("errors").EnumerateArray());
            var workbooks = discovery.RootElement.GetProperty("workbooks").EnumerateArray().ToArray();
            foreach (var (path, saved) in new[] { (sourceFile, sourceSaved), (targetFile, targetSaved) })
            {
                var workbook = Assert.Single(workbooks, entry => entry.GetProperty("location").GetString() == path);
                Assert.Equal(sessionId, workbook.GetProperty("sessionId").GetString());
                Assert.Equal(Path.GetFileName(path), workbook.GetProperty("name").GetString());
                Assert.Equal(saved, workbook.GetProperty("saved").GetBoolean());
            }
        }
        AssertRetainedTargetThroughService(targetFile, targetSheets, targetSaved, targetReadOnly, targetDuplicateGuard);
        using var primary = ReadCrossFileServiceResult(service, "workbook.get-info", sessionId);
        Assert.Equal(sourceFile, primary.RootElement.GetProperty("fullName").GetString());
    }

    private static void AssertRetainedTargetThroughService(
        string targetFile, string[] expectedSheets, bool saved, bool readOnly, bool duplicateGuard)
    {
        // A retained batch selects its source. A separate verification connection
        // reads the live target; disposal detaches, never closes it.
        var verification = new ExcelMcpService();
        Exception? failure = null;
        try
        {
            using var discovery = ReadCrossFileServiceResult(verification, "session.discover", sessionId: null);
            var target = Assert.Single(discovery.RootElement.GetProperty("workbooks").EnumerateArray(),
                workbook => workbook.GetProperty("location").GetString() == targetFile);
            var attached = verification.ProcessAsync(new ServiceRequest
            {
                Command = "session.attach",
                Args = JsonSerializer.Serialize(new
                {
                    workbookId = target.GetProperty("workbookId").GetString()
                }, ServiceProtocol.JsonOptions)
            }).GetAwaiter().GetResult();
            Assert.True(attached.Success, attached.ErrorMessage);
            Assert.False(string.IsNullOrWhiteSpace(attached.SessionId));
            using var result = JsonDocument.Parse(attached.Result!);
            Assert.Equal(attached.SessionId, result.RootElement.GetProperty("sessionId").GetString());
            Assert.Equal(targetFile, result.RootElement.GetProperty("filePath").GetString());
            using var info = ReadCrossFileServiceResult(verification, "workbook.get-info", attached.SessionId);
            Assert.Equal(targetFile, info.RootElement.GetProperty("fullName").GetString());
            Assert.Equal(saved, info.RootElement.GetProperty("saved").GetBoolean());
            Assert.Equal(readOnly, info.RootElement.GetProperty("readOnly").GetBoolean());
            using var sheets = ReadCrossFileServiceResult(verification, "sheet.list", attached.SessionId);
            Assert.Equal(expectedSheets, sheets.RootElement.GetProperty("worksheets").EnumerateArray()
                .Select(sheet => sheet.GetProperty("name").GetString()));
            AssertCrossFileServiceCells(verification, attached.SessionId!, "Sheet1",
                "retained workbook marker", 71, 79, "=B1+8");
            foreach (var sheet in expectedSheets.Where(name => name != "Sheet1"))
            {
                AssertCrossFileServiceCells(verification, attached.SessionId!, sheet,
                    duplicateGuard ? "duplicate guard" : "transferred marker",
                    duplicateGuard ? 12 : 3, duplicateGuard ? 15 : 6,
                    duplicateGuard ? "=B1+3" : "=B1*2");
            }
        }
        catch (Exception ex)
        {
            failure = ex;
        }
        finally
        {
            try
            {
                verification.Dispose();
            }
            catch (Exception ex)
            {
                failure = PersistentServiceCleanupFailures.Combine(failure, ex);
            }
        }
        PersistentServiceCleanupFailures.ThrowIfFailed(failure);
    }

    private static void AssertCrossFileServiceCells(
        ExcelMcpService service, string sessionId, string sheetName,
        string marker, double input, double calculated, string expectedFormula)
    {
        using var values = ReadCrossFileServiceResult(service, "range.get-values", sessionId,
            new { sheetName, rangeAddress = "A1:C1" });
        var row = Assert.Single(values.RootElement.GetProperty("values").EnumerateArray());
        Assert.Equal(3, row.GetArrayLength());
        Assert.Equal(marker, row[0].GetString());
        Assert.Equal(input, row[1].GetDouble());
        Assert.Equal(calculated, row[2].GetDouble());
        using var formulas = ReadCrossFileServiceResult(service, "range.get-formulas", sessionId,
            new { sheetName, rangeAddress = "C1" });
        var formulaRow = Assert.Single(formulas.RootElement.GetProperty("formulas").EnumerateArray());
        Assert.Equal(expectedFormula, Assert.Single(formulaRow.EnumerateArray()).GetString());
    }

    private static JsonDocument ReadCrossFileServiceResult(
        ExcelMcpService service, string command, string? sessionId, object? arguments = null)
    {
        var response = service.ProcessAsync(new ServiceRequest
        {
            Command = command,
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(arguments ?? new { }, ServiceProtocol.JsonOptions)
        }).GetAwaiter().GetResult();
        Assert.True(response.Success, response.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(response.ErrorMessage), response.ErrorMessage);
        Assert.False(string.IsNullOrWhiteSpace(response.Result));
        return JsonDocument.Parse(response.Result!);
    }

    private (string Source, string Target) CreateMarkedWorkbookPair(
        string scenario,
        string sourceSheet,
        string? targetSheet = null,
        bool targetDuplicateGuard = false)
    {
        var sourceFile = CreateMarkedWorkbook(scenario, "Source", sourceSheet);
        var targetFile = CreateMarkedWorkbook(scenario, "Target", targetSheet, targetDuplicateGuard);
        return (sourceFile, targetFile);
    }

    private string CreateMarkedWorkbook(
        string scenario, string suffix, string? additionalSheet = null, bool duplicateGuard = false)
    {
        var safeScenario = string.Concat(
            scenario.Select(character =>
                Path.GetInvalidFileNameChars().Contains(character) ? '_' : character));
        var path = Path.Combine(
            Path.GetDirectoryName(_fixture.WorkbookPath)!,
            $"{safeScenario}_{suffix}_{Guid.NewGuid():N}.xlsx");
        var key = (additionalSheet, duplicateGuard);
        var templates = MarkedWorkbookTemplates.GetOrAdd(
            key,
            static key => SavedWorkbookTemplates.CreateStore(
                templatePath => CreateMarkedWorkbookTemplate(templatePath, key.Sheet, key.DuplicateGuard)));
        templates.CopyTo(path, $"{additionalSheet ?? "base"}-{duplicateGuard}");
        return path;
    }

    private static void CreateMarkedWorkbookTemplate(
        string templatePath, string? additionalSheet, bool duplicateGuard)
    {
        using var manager = new SessionManager();
        var sessionId = manager.CreateSessionForNewFile(templatePath, show: false);
        var batch = manager.GetSession(sessionId)
            ?? throw new InvalidOperationException("The marked workbook template session was not created.");
        batch.Execute((context, _) => SeedWorkbook(context.Book, additionalSheet, duplicateGuard));
        if (!manager.CloseSession(sessionId, save: true))
        {
            throw new InvalidOperationException("The marked workbook template could not be saved.");
        }
    }

    private static void SeedWorkbook(
        Excel.Workbook workbook,
        string? additionalSheet,
        bool duplicateGuard = false)
    {
        Excel.Sheets? sheets = null;
        Excel.Worksheet? sheet = null;
        Excel.Range? values = null;
        Excel.Range? formula = null;
        try
        {
            sheets = workbook.Worksheets;
            sheet = (Excel.Worksheet)sheets.Item["Sheet1"];
            values = sheet.Range["A1:B1"];
            values.Value2 = new object[,] { { "retained workbook marker", 71d } };
            formula = sheet.Range["C1"];
            formula.Formula = "=B1+8";
            if (additionalSheet is not null)
            {
                ComUtilities.Release(ref formula);
                ComUtilities.Release(ref values);
                ComUtilities.Release(ref sheet);
                sheet = (Excel.Worksheet)sheets.Add();
                sheet.Name = additionalSheet;
                values = sheet.Range["A1:B1"];
                values.Value2 = duplicateGuard
                    ? new object[,] { { "duplicate guard", 12d } }
                    : new object[,] { { "transferred marker", 3d } };
                formula = sheet.Range["C1"];
                formula.Formula = duplicateGuard ? "=B1+3" : "=B1*2";
            }
        }
        finally
        {
            ComUtilities.Release(ref formula);
            ComUtilities.Release(ref values);
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref sheets);
        }
    }

    private static void AssertWorkbookState(
        string workbookPath, string[] expectedNames, string? additionalSheet = null, bool duplicateGuard = false)
    {
        using var batch = ExcelSession.BeginBatch(workbookPath);
        batch.Execute((context, _) => AssertWorkbookState(
            context.Book, expectedNames, additionalSheet, duplicateGuard));
    }

    private static void AssertWorkbookStates(
        string sourcePath,
        string[] sourceExpectedNames,
        string? sourceSheet,
        string targetPath,
        string[] targetExpectedNames,
        string? targetSheet = null,
        bool targetDuplicateGuard = false)
    {
        using var batch = ExcelSession.BeginBatch(sourcePath, targetPath);
        batch.Execute((_, _) =>
        {
            AssertWorkbookState(batch.GetWorkbook(sourcePath), sourceExpectedNames, sourceSheet);
            AssertWorkbookState(batch.GetWorkbook(targetPath), targetExpectedNames, targetSheet, targetDuplicateGuard);
        });
    }

    private static void AssertWorkbookState(
        Excel.Workbook workbook,
        string[] expectedNames,
        string? additionalSheet = null,
        bool duplicateGuard = false)
    {
        Excel.Sheets? sheets = null;
        try
        {
            sheets = workbook.Worksheets;
            var names = new List<string>();
            for (var index = 1; index <= sheets.Count; index++)
            {
                Excel.Worksheet? sheet = null;
                try
                {
                    sheet = (Excel.Worksheet)sheets[index];
                    names.Add(sheet.Name);
                }
                finally { ComUtilities.Release(ref sheet); }
            }
            Assert.Equal(expectedNames, names);
            AssertSheetCells(sheets, "Sheet1", "retained workbook marker", 71, 79, "=B1+8");
            if (additionalSheet is not null)
            {
                if (duplicateGuard)
                {
                    AssertSheetCells(sheets, additionalSheet, "duplicate guard", 12, 15, "=B1+3");
                }
                else
                {
                    AssertSheetCells(sheets, additionalSheet, "transferred marker", 3, 6, "=B1*2");
                }
            }
        }
        finally { ComUtilities.Release(ref sheets); }
    }

    private static void AssertSheetCells(
        Excel.Sheets sheets, string sheetName, string marker, double input, double calculated, string expectedFormula)
    {
        Excel.Worksheet? sheet = null;
        Excel.Range? values = null;
        Excel.Range? formula = null;
        try
        {
            sheet = (Excel.Worksheet)sheets[sheetName];
            values = sheet.Range["A1:C1"];
            var cells = Assert.IsType<object[,]>(((object)values.Value2));
            Assert.Equal(1, cells.GetLength(0));
            Assert.Equal(3, cells.GetLength(1));
            Assert.Equal(marker, cells[1, 1]);
            Assert.Equal(input, Convert.ToDouble(cells[1, 2], System.Globalization.CultureInfo.InvariantCulture));
            Assert.Equal(calculated, Convert.ToDouble(cells[1, 3], System.Globalization.CultureInfo.InvariantCulture));
            formula = sheet.Range["C1"];
            Assert.Equal(expectedFormula, ((object)formula.Formula));
        }
        finally
        {
            ComUtilities.Release(ref formula);
            ComUtilities.Release(ref values);
            ComUtilities.Release(ref sheet);
        }
    }
}
