using System.IO.Compression;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Table;

[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Tables")]
[Trait("Speed", "Medium")]
public sealed class TableRangeConversionTests : IClassFixture<TempDirectoryFixture>
{
    private readonly TempDirectoryFixture _fixture;
    private readonly TableCommands _tableCommands = new();
    private readonly RangeCommands _rangeCommands = new();

    public TableRangeConversionTests(TempDirectoryFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void ConvertRange_WithExplicitNormalization_CreatesAndValidatesTable()
    {
        var testFile = CreateTestFile(nameof(ConvertRange_WithExplicitNormalization_CreatesAndValidatesTable));
        using var batch = ExcelSession.BeginBatch(testFile);
        InitializeSalesSheet(batch);
        SetValues(batch, "F1:H3",
        [
            ["Region", null, null],
            ["North", 10, null],
            ["South", 20, null]
        ]);
        _rangeCommands.MergeCells(batch, "Sales", "F1:G1");
        _rangeCommands.SetFormulas(batch, "Sales", "H2:H3", [["=G2*2"], ["=G3*2"]]);

        var result = _tableCommands.ConvertRange(
            batch,
            "Sales",
            "NormalizedTable",
            "F1:H3",
            tableStyle: "TableStyleLight1",
            mergedHeaderPolicy: TableMergedHeaderPolicy.UnmergeAndFill,
            headerPolicy: TableHeaderPolicy.Normalize);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal("$F$1:$H$3", result.EffectiveRange);
        Assert.Equal(["$F$1:$G$1"], result.NormalizedMergedRanges);
        Assert.Contains(
            result.HeaderChanges,
            change => change.Address == "$G$1"
                && change.NewValue == "Region_2"
                && change.Reason == TableHeaderChangeReason.Duplicate);
        Assert.Contains(
            result.HeaderChanges,
            change => change.Address == "$H$1"
                && change.NewValue == "Column3"
                && change.Reason == TableHeaderChangeReason.Blank);
        Assert.NotNull(result.Table);
        Assert.Equal("TableStyleLight1", result.Table.TableStyle);
        Assert.True(result.Validation.IsValid);
        Assert.Contains("Column3", result.Validation.CalculatedColumns);
        Assert.False(result.Validation.ShowTotals);
        Assert.False(result.Rollback.Attempted);
    }

    [Fact]
    public void ConvertRange_WithReportPolicies_RejectsBeforeMutation()
    {
        var testFile = CreateTestFile(nameof(ConvertRange_WithReportPolicies_RejectsBeforeMutation));
        using var batch = ExcelSession.BeginBatch(testFile);
        InitializeSalesSheet(batch);
        SetValues(batch, "F1:G2", [[null, "Amount"], ["Widget", 10]]);

        var exception = Assert.Throws<TableRangeConversionException>(
            () => _tableCommands.ConvertRange(
                batch,
                "Sales",
                "BlockedTable",
                "F1:G2"));

        Assert.Equal(TableConversionFailureStage.Preflight, exception.Details.FailureStage);
        Assert.False(exception.Details.Rollback.Attempted);
        Assert.False(exception.Details.Rollback.Required);
        Assert.Contains(
            exception.Details.PreflightFindings,
            finding => finding.Kind == TablePreflightFindingKind.BlankHeaders);
        Assert.DoesNotContain(_tableCommands.List(batch).Tables, table => table.Name == "BlockedTable");
        Assert.Null(ReadCellValue(batch, "F1"));
    }

    [Fact]
    public void ConvertRange_WhenStyleFailsAfterCreation_RestoresExactRangeState()
    {
        var testFile = CreateTestFile(nameof(ConvertRange_WhenStyleFailsAfterCreation_RestoresExactRangeState));
        using var batch = ExcelSession.BeginBatch(testFile);
        InitializeSalesSheet(batch);
        SetValues(batch, "F1:H3",
        [
            ["Item", "Metrics", null],
            ["One", 10, null],
            ["Two", 20, null]
        ]);
        _rangeCommands.MergeCells(batch, "Sales", "G1:H1");
        _rangeCommands.SetFormulas(batch, "Sales", "H2:H3", [["=G2*2"], ["=G3*2"]]);
        SetRollbackSentinelFormatting(batch);

        var exception = Assert.Throws<TableRangeConversionException>(
            () => _tableCommands.ConvertRange(
                batch,
                "Sales",
                "RollbackTable",
                "F1:H3",
                tableStyle: "NotARealTableStyle",
                mergedHeaderPolicy: TableMergedHeaderPolicy.UnmergeAndFill,
                headerPolicy: TableHeaderPolicy.Normalize));

        Assert.Equal(TableConversionFailureStage.Styling, exception.Details.FailureStage);
        Assert.True(exception.Details.Rollback.Required);
        Assert.True(exception.Details.Rollback.Attempted);
        Assert.True(exception.Details.Rollback.Completed, exception.Details.Rollback.ErrorMessage);
        Assert.True(exception.Details.Rollback.Verified, exception.Details.Rollback.ErrorMessage);
        Assert.NotNull(exception.InnerException);
        Assert.DoesNotContain(_tableCommands.List(batch).Tables, table => table.Name == "RollbackTable");

        var restored = ReadRollbackState(batch);
        Assert.Equal("Item", restored.ItemHeader);
        Assert.Equal("Metrics", restored.MetricsHeader);
        Assert.Null(restored.MergedHeaderTail);
        Assert.Equal("=G2*2", restored.FirstFormula);
        Assert.Equal("=G3*2", restored.SecondFormula);
        Assert.True(restored.IsMerged);
        Assert.Equal("$G$1:$H$1", restored.MergeAddress);
        Assert.Equal("0.00", restored.NumberFormat);
        Assert.True(restored.Bold);
        Assert.Equal(255, restored.FillColor);
        Assert.Equal(
            Convert.ToInt32(Excel.XlBorderWeight.xlMedium, System.Globalization.CultureInfo.InvariantCulture),
            restored.BottomBorderWeight);
        Assert.DoesNotContain(
            restored.WorksheetNames,
            name => name.StartsWith("__ExcelMcpTblRb_", StringComparison.Ordinal));
    }

    [Fact]
    public void ConvertRange_WhenFormulaEvaluatesToError_RollsBackAndReportsValidation()
    {
        var testFile = CreateTestFile(nameof(ConvertRange_WhenFormulaEvaluatesToError_RollsBackAndReportsValidation));
        using var batch = ExcelSession.BeginBatch(testFile);
        InitializeSalesSheet(batch);
        SetValues(batch, "A1:B2", [["Amount", "Ratio"], [10, null]]);
        _rangeCommands.SetFormulas(batch, "Sales", "B2", [["=1/0"]]);

        var exception = Assert.Throws<TableRangeConversionException>(
            () => _tableCommands.ConvertRange(
                batch,
                "Sales",
                "ErrorTable",
                "A1:B2"));

        Assert.Equal(TableConversionFailureStage.Validation, exception.Details.FailureStage);
        Assert.NotNull(exception.Details.Validation);
        Assert.Contains(
            exception.Details.Validation.Findings,
            finding => finding.Kind == TableConversionValidationFindingKind.FormulaError
                && finding.Addresses.Contains("$B$2"));
        Assert.True(exception.Details.Rollback.Completed);
        Assert.True(exception.Details.Rollback.Verified);
        Assert.DoesNotContain(_tableCommands.List(batch).Tables, table => table.Name == "ErrorTable");
        Assert.Equal("=1/0", ReadCellFormula(batch, "B2"));
    }

    [Fact]
    public void ConvertRange_WithBodyMergeAndUnmergePolicy_RejectsBeforeMutation()
    {
        var testFile = CreateTestFile(nameof(ConvertRange_WithBodyMergeAndUnmergePolicy_RejectsBeforeMutation));
        using var batch = ExcelSession.BeginBatch(testFile);
        InitializeSalesSheet(batch);
        SetValues(batch, "A1:B2", [["Name", "Amount"], ["Widget", 10]]);
        _rangeCommands.MergeCells(batch, "Sales", "A2:B2");

        var exception = Assert.Throws<TableRangeConversionException>(
            () => _tableCommands.ConvertRange(
                batch,
                "Sales",
                "MergedBodyTable",
                "A1:B2",
                mergedHeaderPolicy: TableMergedHeaderPolicy.UnmergeAndFill));

        Assert.Equal(TableConversionFailureStage.Preflight, exception.Details.FailureStage);
        Assert.False(exception.Details.Rollback.Attempted);
        Assert.True(_rangeCommands.GetMergeInfo(batch, "Sales", "A2:B2").IsMerged);
    }

    [Fact]
    public void ConvertRange_WithHeaderOnlyRange_RejectsWithoutShiftingFollowingData()
    {
        var testFile = CreateTestFile(nameof(ConvertRange_WithHeaderOnlyRange_RejectsWithoutShiftingFollowingData));
        using var batch = ExcelSession.BeginBatch(testFile);
        InitializeSalesSheet(batch);
        SetValues(batch, "A1:B2", [["Name", "Amount"], ["Widget", 10]]);

        var exception = Assert.Throws<TableRangeConversionException>(
            () => _tableCommands.ConvertRange(
                batch,
                "Sales",
                "HeaderOnlyTable",
                "A1:B1"));

        Assert.Equal(TableConversionFailureStage.Preflight, exception.Details.FailureStage);
        Assert.Contains(
            exception.Details.PreflightFindings,
            finding => finding.Kind == TablePreflightFindingKind.HeaderOnlyRange);
        Assert.Equal("Widget", ReadCellValue(batch, "A2"));
        Assert.Equal(10d, ReadCellValue(batch, "B2"));
    }

    [Fact]
    public void ConvertRange_WithFormulaHeader_RejectsWithoutConvertingHeader()
    {
        var testFile = CreateTestFile(nameof(ConvertRange_WithFormulaHeader_RejectsWithoutConvertingHeader));
        using var batch = ExcelSession.BeginBatch(testFile);
        InitializeSalesSheet(batch);
        SetValues(batch, "A1:B2", [["Name", "Amount"], ["Widget", 10]]);
        _rangeCommands.SetFormulas(batch, "Sales", "A1", [["=\"Name\""]]);

        var exception = Assert.Throws<TableRangeConversionException>(
            () => _tableCommands.ConvertRange(
                batch,
                "Sales",
                "FormulaHeaderTable",
                "A1:B2"));

        Assert.Equal(TableConversionFailureStage.Preflight, exception.Details.FailureStage);
        Assert.Contains(
            exception.Details.PreflightFindings,
            finding => finding.Kind == TablePreflightFindingKind.LossyHeaders
                && finding.Addresses.Contains("$A$1"));
        Assert.Equal("=\"Name\"", ReadCellFormula(batch, "A1"));
    }

    private static object? ReadCellValue(IExcelBatch batch, string address)
    {
        return batch.Execute((ctx, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, "Sales");
                cell = sheet!.Range[address];
                return cell.Value2;
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static string? ReadCellFormula(IExcelBatch batch, string address)
    {
        return batch.Execute((ctx, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, "Sales");
                cell = sheet!.Range[address];
                return Convert.ToString(cell.Formula2, System.Globalization.CultureInfo.InvariantCulture);
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private string CreateTestFile(string testName)
    {
        string path = Path.Join(
            _fixture.TempDir,
            $"{nameof(TableRangeConversionTests)}_{testName}_{Guid.NewGuid():N}.xlsx");
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        WriteEntry(
            archive,
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>""");
        WriteEntry(
            archive,
            "_rels/.rels",
            """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>""");
        WriteEntry(
            archive,
            "xl/workbook.xml",
            """<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>""");
        WriteEntry(
            archive,
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>""");
        WriteEntry(
            archive,
            "xl/worksheets/sheet1.xml",
            """<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>""");
        return path;
    }

    private static void WriteEntry(ZipArchive archive, string name, string content)
    {
        ZipArchiveEntry entry = archive.CreateEntry(name);
        using var writer = new StreamWriter(entry.Open());
        writer.Write(content);
    }

    private static void InitializeSalesSheet(IExcelBatch batch)
    {
        batch.Execute((ctx, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[1];
                sheet.Name = "Sales";
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }

    private void SetValues(IExcelBatch batch, string address, List<List<object?>> values) =>
        _rangeCommands.SetValues(batch, "Sales", address, values);

    private static void SetRollbackSentinelFormatting(IExcelBatch batch)
    {
        batch.Execute((ctx, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? header = null;
            Excel.Range? numeric = null;
            Excel.Font? font = null;
            Excel.Interior? interior = null;
            Excel.Borders? borders = null;
            Excel.Border? border = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, "Sales");
                header = sheet!.Range["F1"];
                numeric = sheet.Range["G2:H3"];
                font = header.Font;
                interior = header.Interior;
                borders = header.Borders;
                font.Bold = true;
                interior.Color = 255;
                border = borders[Excel.XlBordersIndex.xlEdgeBottom];
                border.Weight = Excel.XlBorderWeight.xlMedium;
                numeric.NumberFormat = "0.00";
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref border);
                ComUtilities.Release(ref borders);
                ComUtilities.Release(ref interior);
                ComUtilities.Release(ref font);
                ComUtilities.Release(ref numeric);
                ComUtilities.Release(ref header);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static RollbackState ReadRollbackState(IExcelBatch batch)
    {
        return batch.Execute((ctx, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? itemHeader = null;
            Excel.Range? metricsHeader = null;
            Excel.Range? mergedHeaderTail = null;
            Excel.Range? firstFormula = null;
            Excel.Range? secondFormula = null;
            Excel.Range? numeric = null;
            Excel.Range? mergeArea = null;
            Excel.Font? font = null;
            Excel.Interior? interior = null;
            Excel.Borders? borders = null;
            Excel.Border? border = null;
            Excel.Sheets? sheets = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, "Sales");
                itemHeader = sheet!.Range["F1"];
                metricsHeader = sheet.Range["G1"];
                mergedHeaderTail = sheet.Range["H1"];
                firstFormula = sheet.Range["H2"];
                secondFormula = sheet.Range["H3"];
                numeric = sheet.Range["G2:H3"];
                mergeArea = metricsHeader.MergeArea;
                font = itemHeader.Font;
                interior = itemHeader.Interior;
                borders = itemHeader.Borders;
                border = borders[Excel.XlBordersIndex.xlEdgeBottom];
                sheets = ctx.Book.Worksheets;

                var worksheetNames = new List<string>();
                for (var index = 1; index <= sheets.Count; index++)
                {
                    Excel.Worksheet? current = null;
                    try
                    {
                        current = (Excel.Worksheet)sheets[index];
                        worksheetNames.Add(current.Name);
                    }
                    finally
                    {
                        ComUtilities.Release(ref current);
                    }
                }

                return new RollbackState(
                    itemHeader.Value2,
                    metricsHeader.Value2,
                    mergedHeaderTail.Value2,
                    firstFormula.Formula2,
                    secondFormula.Formula2,
                    Convert.ToBoolean(metricsHeader.MergeCells),
                    mergeArea.Address,
                    Convert.ToString(numeric.NumberFormat, System.Globalization.CultureInfo.InvariantCulture),
                    Convert.ToBoolean(font.Bold),
                    Convert.ToInt32(interior.Color, System.Globalization.CultureInfo.InvariantCulture),
                    Convert.ToInt32(border.Weight, System.Globalization.CultureInfo.InvariantCulture),
                    worksheetNames);
            }
            finally
            {
                ComUtilities.Release(ref sheets);
                ComUtilities.Release(ref border);
                ComUtilities.Release(ref borders);
                ComUtilities.Release(ref interior);
                ComUtilities.Release(ref font);
                ComUtilities.Release(ref mergeArea);
                ComUtilities.Release(ref numeric);
                ComUtilities.Release(ref secondFormula);
                ComUtilities.Release(ref firstFormula);
                ComUtilities.Release(ref mergedHeaderTail);
                ComUtilities.Release(ref metricsHeader);
                ComUtilities.Release(ref itemHeader);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private sealed record RollbackState(
        object? ItemHeader,
        object? MetricsHeader,
        object? MergedHeaderTail,
        object? FirstFormula,
        object? SecondFormula,
        bool IsMerged,
        string MergeAddress,
        string? NumberFormat,
        bool Bold,
        int FillColor,
        int BottomBorderWeight,
        List<string> WorksheetNames);
}
