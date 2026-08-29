using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Commands.Workbook;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Workbook;

[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "Workbook")]
[Trait("RequiresExcel", "true")]
public sealed class WorkbookIntegrityCommandsTests : IClassFixture<FileTestsFixture>
{
    private readonly WorkbookCommands _commands = new();
    private readonly FileTestsFixture _fixture;

    public WorkbookIntegrityCommandsTests(FileTestsFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void ValidateIntegrity_CleanWorkbook_ReturnsPassedWithoutChangingWorkbook()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        SetAutomaticCalculation(batch);
        bool savedBefore = batch.Execute((context, _) => context.Book.Saved);

        var result = _commands.ValidateIntegrity(batch);

        Assert.True(result.Success);
        Assert.Equal(WorkbookIntegrityStatus.Passed, result.OverallStatus);
        Assert.Equal(0, result.FindingCount);
        Assert.Empty(result.Groups);
        Assert.False(result.FindingsTruncated);
        Assert.Equal(3, result.CheckedChecks.Count);
        Assert.False(string.IsNullOrWhiteSpace(result.CalculationMode));
        Assert.False(string.IsNullOrWhiteSpace(result.CalculationState));
        Assert.Equal(savedBefore, batch.Execute((context, _) => context.Book.Saved));
    }

    [Fact]
    public void ValidateIntegrity_FormulaErrors_UsesCanonicalErrorsAndGroupsBrokenReferences()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        WriteFormulas(batch, "A1:C1", ["=1/0", "=INDIRECT(\"A0\")", "=1+1"]);

        var result = _commands.ValidateIntegrity(
            batch,
            checks: [WorkbookIntegrityCheck.FormulaErrors],
            worksheetNames: ["Sheet1"],
            maxFindings: 1);

        Assert.True(result.Success);
        Assert.Equal(WorkbookIntegrityStatus.Failed, result.OverallStatus);
        Assert.Equal(2, result.ErrorCount);
        Assert.Equal(2, result.FindingCount);
        Assert.Equal(2, result.Groups.Sum(group => group.Count));
        Assert.True(result.FindingsTruncated);
        Assert.Single(result.CheckedWorksheets);

        var retainedFinding = Assert.Single(result.Groups.SelectMany(group => group.Findings));
        Assert.Equal(WorkbookIntegrityReliability.Deterministic, retainedFinding.Reliability);
        Assert.True(retainedFinding.ErrorName is "#DIV/0!" or "#REF!");
        Assert.StartsWith("=", retainedFinding.Formula, StringComparison.Ordinal);
        Assert.True(retainedFinding.CellAddress is "A1" or "B1");

        Assert.Contains(result.Groups, group =>
            group.Category == WorkbookIntegrityCategory.FormulaError ||
            group.Category == WorkbookIntegrityCategory.BrokenReference);
    }

    [Fact]
    public void ValidateIntegrity_WorksheetFilter_ExcludesUnselectedSheets()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        AddWorksheetWithReferenceError(batch, "Excluded");

        var result = _commands.ValidateIntegrity(
            batch,
            checks: [WorkbookIntegrityCheck.FormulaErrors],
            worksheetNames: ["Sheet1"]);

        Assert.Equal(WorkbookIntegrityStatus.Passed, result.OverallStatus);
        Assert.Equal("Sheet1", Assert.Single(result.CheckedWorksheets));
        Assert.Equal(0, result.FindingCount);
    }

    [Fact]
    public void ValidateIntegrity_BrokenReferenceHiddenByIfError_IsStillReported()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        WriteFormulas(batch, "A1:B1", ["=IFERROR(#REF!,0)", "=\"#REF!\""]);

        var result = _commands.ValidateIntegrity(
            batch,
            checks: [WorkbookIntegrityCheck.FormulaErrors],
            worksheetNames: ["Sheet1"]);

        Assert.Equal(WorkbookIntegrityStatus.Failed, result.OverallStatus);
        var finding = Assert.Single(result.Groups
            .Where(group => group.Category == WorkbookIntegrityCategory.BrokenReference)
            .SelectMany(group => group.Findings));
        Assert.Equal("broken-reference-token", finding.Code);
        Assert.Equal("A1", finding.CellAddress);
        Assert.Equal(WorkbookIntegrityReliability.Deterministic, finding.Reliability);
    }

    [Fact]
    public void ValidateIntegrity_ManualCalculation_ReportsHeuristicWarningWithoutCalculating()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        batch.Execute((context, _) =>
        {
            context.App.Calculation = Excel.XlCalculation.xlCalculationManual;
            return 0;
        });

        try
        {
            var result = _commands.ValidateIntegrity(
                batch,
                checks: [WorkbookIntegrityCheck.FormulaErrors],
                worksheetNames: ["Sheet1"]);

            Assert.Equal(WorkbookIntegrityStatus.PassedWithWarnings, result.OverallStatus);
            Assert.Equal("manual", result.CalculationMode);
            var finding = Assert.Single(result.Groups
                .Where(group => group.Category == WorkbookIntegrityCategory.CalculationState)
                .SelectMany(group => group.Findings));
            Assert.Equal("manual-calculation", finding.Code);
            Assert.Equal(WorkbookIntegrityReliability.Heuristic, finding.Reliability);
        }
        finally
        {
            SetAutomaticCalculation(batch);
        }
    }

    [Fact]
    public void ValidateIntegrity_ControlTotals_HonorsAbsoluteTolerance()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        WriteValues(batch, "B2", 100.005d);

        var passing = _commands.ValidateIntegrity(
            batch,
            checks: [WorkbookIntegrityCheck.ControlTotals],
            controlTotals:
            [
                new WorkbookControlTotalExpectation
                {
                    SheetName = "sheet1",
                    CellAddress = "B2",
                    ExpectedValue = 100d,
                    Tolerance = 0.01d
                }
            ]);
        var failing = _commands.ValidateIntegrity(
            batch,
            checks: [WorkbookIntegrityCheck.ControlTotals],
            controlTotals:
            [
                new WorkbookControlTotalExpectation
                {
                    SheetName = "Sheet1",
                    CellAddress = "B2",
                    ExpectedValue = 99d,
                    Tolerance = 0.1d
                }
            ]);

        Assert.Equal(WorkbookIntegrityStatus.Passed, passing.OverallStatus);
        Assert.Equal("Sheet1", Assert.Single(passing.CheckedWorksheets));
        Assert.Equal(WorkbookIntegrityStatus.Failed, failing.OverallStatus);
        var finding = Assert.Single(failing.Groups.SelectMany(group => group.Findings));
        Assert.Equal(WorkbookIntegrityCategory.ControlTotal, finding.Category);
        Assert.Equal(99d, finding.ExpectedValue);
        Assert.Equal(100.005d, finding.ActualValue);
        Assert.Equal(0.1d, finding.Tolerance);
        Assert.Equal("Sheet1", finding.SheetName);
        Assert.Equal("B2", finding.CellAddress);
    }

    [Fact]
    public void ValidateIntegrity_InvalidScopedArguments_AreRejectedClearly()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        var missingExpectations = Assert.Throws<ArgumentException>(() =>
            _commands.ValidateIntegrity(
                batch,
                checks: [WorkbookIntegrityCheck.ControlTotals]));
        Assert.Equal("controlTotals", missingExpectations.ParamName);

        var irrelevantWorksheetFilter = Assert.Throws<ArgumentException>(() =>
            _commands.ValidateIntegrity(
                batch,
                checks: [WorkbookIntegrityCheck.ExternalLinks],
                worksheetNames: ["Sheet1"]));
        Assert.Equal("worksheetNames", irrelevantWorksheetFilter.ParamName);

        var ignoredExpectations = Assert.Throws<ArgumentException>(() =>
            _commands.ValidateIntegrity(
                batch,
                checks: [WorkbookIntegrityCheck.FormulaErrors],
                controlTotals:
                [
                    new WorkbookControlTotalExpectation
                    {
                        SheetName = "Sheet1",
                        CellAddress = "A1",
                        ExpectedValue = 1d
                    }
                ]));
        Assert.Equal("controlTotals", ignoredExpectations.ParamName);

        var missingExpectedValue = Assert.Throws<ArgumentException>(() =>
            _commands.ValidateIntegrity(
                batch,
                checks: [WorkbookIntegrityCheck.ControlTotals],
                controlTotals:
                [
                    new WorkbookControlTotalExpectation
                    {
                        SheetName = "Sheet1",
                        CellAddress = "A1"
                    }
                ]));
        Assert.Equal("controlTotals", missingExpectedValue.ParamName);
        Assert.Contains("finite expected value", missingExpectedValue.Message, StringComparison.Ordinal);

        var invalidAddress = Assert.Throws<ArgumentException>(() =>
            _commands.ValidateIntegrity(
                batch,
                checks: [WorkbookIntegrityCheck.ControlTotals],
                controlTotals:
                [
                    new WorkbookControlTotalExpectation
                    {
                        SheetName = "Sheet1",
                        CellAddress = "not a cell",
                        ExpectedValue = 1d
                    }
                ]));
        Assert.Equal("controlTotals", invalidAddress.ParamName);
        Assert.Contains("valid cell address", invalidAddress.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ValidateIntegrity_TableCheck_FlagsCalculatedColumnOutlierAsHeuristic()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        WriteTableData(batch);
        var tableCommands = new TableCommands();
        tableCommands.Create(batch, "Sheet1", "IntegrityTable", "A1:B4");
        WriteTableFormulaOutlier(batch);

        var result = _commands.ValidateIntegrity(
            batch,
            checks: [WorkbookIntegrityCheck.Tables],
            worksheetNames: ["Sheet1"]);

        Assert.Equal(WorkbookIntegrityStatus.PassedWithWarnings, result.OverallStatus);
        var finding = Assert.Single(result.Groups
            .Where(group => group.Category == WorkbookIntegrityCategory.CalculatedColumn)
            .SelectMany(group => group.Findings));
        Assert.Equal(WorkbookIntegritySeverity.Warning, finding.Severity);
        Assert.Equal(WorkbookIntegrityReliability.Heuristic, finding.Reliability);
        Assert.Equal("IntegrityTable", finding.TableName);
        Assert.Equal("Result", finding.ColumnName);
        Assert.Equal("B4", finding.CellAddress);
    }

    [Fact]
    public void ValidateIntegrity_TableCheck_FlagsHiddenHeadersAsDeterministicWarning()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        WriteTableData(batch);
        var tableCommands = new TableCommands();
        tableCommands.Create(batch, "Sheet1", "HiddenHeaderTable", "A1:B4");
        SetTableHeadersVisible(batch, visible: false);

        var result = _commands.ValidateIntegrity(
            batch,
            checks: [WorkbookIntegrityCheck.Tables]);

        Assert.Equal(WorkbookIntegrityStatus.PassedWithWarnings, result.OverallStatus);
        var finding = Assert.Single(result.Groups
            .Where(group => group.Category == WorkbookIntegrityCategory.TableHeader)
            .SelectMany(group => group.Findings));
        Assert.Equal("table-headers-hidden", finding.Code);
        Assert.Equal(WorkbookIntegrityReliability.Deterministic, finding.Reliability);
        Assert.Equal("HiddenHeaderTable", finding.TableName);
    }

    [Fact]
    public void ValidateIntegrity_MissingExternalWorkbook_ReportsBrokenLink()
    {
        var sourcePath = _fixture.CreateTestFile();
        var targetPath = _fixture.CreateTestFile();
        WriteExternalFormula(targetPath, sourcePath);
        System.IO.File.Delete(sourcePath);

        using var batch = ExcelSession.BeginBatch(targetPath);
        var result = _commands.ValidateIntegrity(
            batch,
            checks: [WorkbookIntegrityCheck.ExternalLinks]);

        Assert.Equal(WorkbookIntegrityStatus.Failed, result.OverallStatus);
        var finding = Assert.Single(result.Groups
            .Where(group => group.Category == WorkbookIntegrityCategory.ExternalLink)
            .SelectMany(group => group.Findings));
        Assert.Equal(WorkbookIntegritySeverity.Error, finding.Severity);
        Assert.Equal(WorkbookIntegrityReliability.Deterministic, finding.Reliability);
        Assert.Equal("missing-file", finding.LinkStatus);
        Assert.Equal(Path.GetFullPath(sourcePath), finding.LinkSource, ignoreCase: true);
    }

    private static void WriteFormulas(IExcelBatch batch, string address, string[] formulas)
    {
        batch.Execute((context, _) =>
        {
            Excel.Sheets? worksheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            try
            {
                worksheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)worksheets.Item[1];
                range = sheet.Range[address];
                object[,] values = (object[,])Array.CreateInstance(
                    typeof(object),
                    [1, formulas.Length],
                    [1, 1]);
                for (int column = 1; column <= formulas.Length; column++)
                {
                    values[1, column] = formulas[column - 1];
                }

                range.Formula2 = values;
                context.App.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                context.App.CalculateFull();
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref worksheets);
            }

            return 0;
        });
    }

    private static void AddWorksheetWithReferenceError(IExcelBatch batch, string worksheetName)
    {
        batch.Execute((context, _) =>
        {
            Excel.Sheets? worksheets = null;
            Excel.Worksheet? worksheet = null;
            Excel.Range? cell = null;
            try
            {
                worksheets = context.Book.Worksheets;
                worksheet = (Excel.Worksheet)worksheets.Add();
                worksheet.Name = worksheetName;
                cell = worksheet.Range["A1"];
                cell.Formula2 = "=INDIRECT(\"A0\")";
                context.App.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                context.App.CalculateFull();
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref worksheet);
                ComUtilities.Release(ref worksheets);
            }

            return 0;
        });
    }

    private static void WriteValues(IExcelBatch batch, string address, object value)
    {
        batch.Execute((context, _) =>
        {
            Excel.Sheets? worksheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            try
            {
                worksheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)worksheets.Item[1];
                range = sheet.Range[address];
                context.App.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                range.Value2 = value;
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref worksheets);
            }

            return 0;
        });
    }

    private static void SetAutomaticCalculation(IExcelBatch batch)
    {
        batch.Execute((context, _) =>
        {
            context.App.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            return 0;
        });
    }

    private static void WriteTableData(IExcelBatch batch)
    {
        batch.Execute((context, _) =>
        {
            Excel.Sheets? worksheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            try
            {
                worksheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)worksheets.Item[1];
                range = sheet.Range["A1:B4"];
                range.Value2 = new object?[,]
                {
                    { "Input", "Result" },
                    { 1d, null },
                    { 2d, null },
                    { 3d, null }
                };
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref worksheets);
            }

            return 0;
        });
    }

    private static void WriteTableFormulaOutlier(IExcelBatch batch)
    {
        batch.Execute((context, _) =>
        {
            Excel.Sheets? worksheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? formulas = null;
            Excel.Range? outlier = null;
            try
            {
                worksheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)worksheets.Item[1];
                formulas = sheet.Range["B2:B3"];
                formulas.Formula2R1C1 = "=RC[-1]*2";
                outlier = sheet.Range["B4"];
                outlier.Value2 = 999d;
            }
            finally
            {
                ComUtilities.Release(ref outlier);
                ComUtilities.Release(ref formulas);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref worksheets);
            }

            return 0;
        });
    }

    private static void SetTableHeadersVisible(IExcelBatch batch, bool visible)
    {
        batch.Execute((context, _) =>
        {
            Excel.Sheets? worksheets = null;
            Excel.Worksheet? sheet = null;
            Excel.ListObjects? tables = null;
            Excel.ListObject? table = null;
            try
            {
                worksheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)worksheets.Item[1];
                tables = sheet.ListObjects;
                table = tables.Item[1];
                table.ShowHeaders = visible;
            }
            finally
            {
                ComUtilities.Release(ref table);
                ComUtilities.Release(ref tables);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref worksheets);
            }

            return 0;
        });
    }

    private static void WriteExternalFormula(string workbookPath, string sourcePath)
    {
        var sourceDirectory = Path.GetDirectoryName(sourcePath)!.Replace("'", "''", StringComparison.Ordinal);
        var sourceFileName = Path.GetFileName(sourcePath);
        var formula = $"='{sourceDirectory}\\[{sourceFileName}]Sheet1'!$A$1";

        using var batch = ExcelSession.BeginBatch(workbookPath);
        batch.Execute((context, _) =>
        {
            Excel.Sheets? worksheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            try
            {
                worksheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)worksheets.Item[1];
                cell = sheet.Range["A1"];
                cell.Formula = formula;
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref worksheets);
            }

            return 0;
        });
        batch.Save();
    }
}
