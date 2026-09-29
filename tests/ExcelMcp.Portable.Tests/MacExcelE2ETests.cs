using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacExcelTheoryAttribute : TheoryAttribute
{
    public MacExcelTheoryAttribute()
    {
        if (!OperatingSystem.IsMacOS() || Environment.GetEnvironmentVariable("EXCELMCP_MAC_E2E") != "1")
        {
            Skip = "Explicit macOS desktop Excel run required: scripts/Test-MacE2E.ps1.";
        }
    }
}

[CollectionDefinition("Mac Excel E2E", DisableParallelization = true)]
public sealed class MacExcelCollectionDefinition;

[Collection("Mac Excel E2E")]
public sealed class MacExcelE2ETests(ITestOutputHelper output)
{
    private static readonly int[][] SentinelValues = [[9876]];
    private static readonly int[][] UnsavedValues = [[999]];
    private static readonly string[][] SumFormula = [["=SUM(B2:B3)"]];
    private static readonly string[][] RangeExpansionFormula = [["=D5*2"]];
    private static readonly string[][] RangeExpansionFormats =
    [
        ["0", "0.00"],
        ["@", "#,##0"]
    ];
    private static readonly string[] InitialSheetNames = ["Data", "Spare"];

    [MacExcelTheory]
    [InlineData("cli")]
    [InlineData("mcp")]
    [Trait("Category", "Integration")]
    [Trait("RequiresExcel", "true")]
    [Trait("Feature", "MacHandoff")]
    public async Task ExistingWorkbook_RealEntryPointRoundTrip(string entryPoint)
    {
        Assert.Equal(0, MacAutomationAccess.Check());
        var root = FindRepository();
        var directory = Directory.CreateTempSubdirectory("excelmcp-mac-e2e-");
        var workbookName = $"main with spaces-{Guid.NewGuid():N}.xlsx";
        var main = Path.Combine(directory.FullName, workbookName);
        var created = Path.Combine(directory.FullName, $"created-{Guid.NewGuid():N}.XLSX");
        var createdMacro = Path.Combine(directory.FullName, $"created-{Guid.NewGuid():N}.xlsm");
        var duplicateDirectory = Directory.CreateTempSubdirectory("excelmcp-mac-e2e-duplicate-");
        var duplicate = Path.Combine(duplicateDirectory.FullName, workbookName);
        var sentinel = Path.Combine(directory.FullName, $"sentinel-{Guid.NewGuid():N}.xlsx");
        var dataFile = Path.Combine(directory.FullName, "values.json");
        CreateBlankWorkbook(main);
        File.Copy(main, duplicate);
        File.Copy(main, sentinel);
        await File.WriteAllTextAsync(dataFile, """[["Item","Amount"],["Alpha",10],["Beta",20]]""");
        using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(4));
        await using var client = await EntryPointClient.CreateAsync(root, entryPoint, output, deadline.Token);
        var completed = false;
        try
        {
            var mainSession = SessionId(await client.CallAsync("file", "open", null,
                new() { ["path"] = main }, deadline.Token));
            await InitializeWorkbookAsync(
                client,
                mainSession,
                includeSpare: true,
                deadline.Token);
            if (Environment.GetEnvironmentVariable("EXCELMCP_MAC_PYTHON_E2E") == "1")
            {
                Success(await client.CallAsync("pythoninexcel", "set-formula", mainSession,
                    new()
                    {
                        ["sheet_name"] = "Data",
                        ["range_address"] = "Z1",
                        ["code"] = "\"ExcelMcp\" + \" Python\"",
                        ["return_type"] = 0
                    }, deadline.Token));
                var pythonResult = Success(await client.CallAsync(
                    "pythoninexcel", "get-result", mainSession,
                    new()
                    {
                        ["sheet_name"] = "Data",
                        ["range_address"] = "Z1",
                        ["max_wait_seconds"] = 30
                    }, deadline.Token));
                Assert.Contains("Z1", pythonResult.GetProperty("rangeAddress").GetString(), StringComparison.Ordinal);
                var pythonFormula = pythonResult.GetProperty("formula").GetString();
                Assert.StartsWith("=PY(", pythonFormula, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("\"\"ExcelMcp\"\"", pythonFormula, StringComparison.Ordinal);
                Assert.Equal("ExcelMcp Python", pythonResult.GetProperty("value").GetString());
                Assert.False(pythonResult.GetProperty("isPythonObject").GetBoolean());
                Assert.False(pythonResult.GetProperty("isPythonError").GetBoolean());
            }
            var duplicateName = await client.CallAsync("file", "open", null,
                new() { ["path"] = duplicate }, deadline.Token);
            Assert.False(duplicateName.GetProperty("success").GetBoolean());
            Assert.Contains("same name", duplicateName.GetProperty("errorMessage").GetString(),
                StringComparison.OrdinalIgnoreCase);
            Assert.Equal("ComInterop", duplicateName.GetProperty("errorCategory").GetString());

            var sheets = Success(await client.CallAsync("sheet", "list", mainSession, new(), deadline.Token));
            Assert.Equal(InitialSheetNames, sheets.GetProperty("worksheets").EnumerateArray()
                .Select(sheet => sheet.GetProperty("name").GetString()!).ToArray());
            Success(await client.CallAsync("sheet", "rename", mainSession,
                new() { ["old_name"] = "Spare", ["new_name"] = "DeleteMe" }, deadline.Token));
            Success(await client.CallAsync("sheet", "delete", mainSession,
                new() { ["sheet_name"] = "DeleteMe" }, deadline.Token));
            sheets = Success(await client.CallAsync("sheet", "list", mainSession, new(), deadline.Token));
            Assert.Equal("Data", Assert.Single(sheets.GetProperty("worksheets").EnumerateArray()).GetProperty("name").GetString());
            Success(await client.CallAsync("sheet", "create", mainSession,
                new() { ["sheet_name"] = "Created" }, deadline.Token));
            Success(await client.CallAsync("worksheetstyle", "set-tab-color", mainSession,
                new()
                {
                    ["sheet_name"] = "Created",
                    ["red"] = 17,
                    ["green"] = 34,
                    ["blue"] = 51
                }, deadline.Token));
            var tabColor = Success(await client.CallAsync(
                "worksheetstyle", "get-tab-color", mainSession,
                new() { ["sheet_name"] = "Created" }, deadline.Token));
            Assert.True(tabColor.GetProperty("hasColor").GetBoolean());
            Assert.Equal(17, tabColor.GetProperty("red").GetInt32());
            Assert.Equal(34, tabColor.GetProperty("green").GetInt32());
            Assert.Equal(51, tabColor.GetProperty("blue").GetInt32());
            Assert.Equal("#112233", tabColor.GetProperty("hexColor").GetString());
            Success(await client.CallAsync("worksheetstyle", "clear-tab-color", mainSession,
                new() { ["sheet_name"] = "Created" }, deadline.Token));
            tabColor = Success(await client.CallAsync(
                "worksheetstyle", "get-tab-color", mainSession,
                new() { ["sheet_name"] = "Created" }, deadline.Token));
            Assert.False(tabColor.GetProperty("hasColor").GetBoolean());
            Success(await client.CallAsync("worksheetstyle", "hide", mainSession,
                new() { ["sheet_name"] = "Created" }, deadline.Token));
            var visibility = Success(await client.CallAsync(
                "worksheetstyle", "get-visibility", mainSession,
                new() { ["sheet_name"] = "Created" }, deadline.Token));
            Assert.Equal("Hidden", visibility.GetProperty("visibilityName").GetString());
            Success(await client.CallAsync("worksheetstyle", "very-hide", mainSession,
                new() { ["sheet_name"] = "Created" }, deadline.Token));
            visibility = Success(await client.CallAsync(
                "worksheetstyle", "get-visibility", mainSession,
                new() { ["sheet_name"] = "Created" }, deadline.Token));
            Assert.Equal("VeryHidden", visibility.GetProperty("visibilityName").GetString());
            Success(await client.CallAsync("worksheetstyle", "set-visibility", mainSession,
                new() { ["sheet_name"] = "Created", ["visibility"] = "visible" }, deadline.Token));
            Success(await client.CallAsync("worksheetstyle", "hide", mainSession,
                new() { ["sheet_name"] = "Created" }, deadline.Token));
            Success(await client.CallAsync("worksheetstyle", "show", mainSession,
                new() { ["sheet_name"] = "Created" }, deadline.Token));
            sheets = Success(await client.CallAsync("sheet", "list", mainSession, new(), deadline.Token));
            Assert.Contains(
                sheets.GetProperty("worksheets").EnumerateArray(),
                sheet => sheet.GetProperty("name").GetString() == "Created"
                    && sheet.GetProperty("visible").GetBoolean());
            Success(await client.CallAsync("sheet", "delete", mainSession,
                new() { ["sheet_name"] = "Created" }, deadline.Token));
            var dirtyPowerQueryRead = await client.CallAsync(
                "powerquery", "list", mainSession, new(), deadline.Token);
            Assert.False(dirtyPowerQueryRead.GetProperty("success").GetBoolean());
            Assert.Contains(
                "no supported local macOS API",
                dirtyPowerQueryRead.GetProperty("errorMessage").GetString(),
                StringComparison.Ordinal);

            var createdSession = SessionId(await client.CallAsync("file", "create", null,
                new() { ["path"] = created }, deadline.Token));
            Success(await client.CallAsync("range", "set-values", createdSession,
                RangeArgsOnSheet("Sheet1", "A1:B1", ("values", new object?[][] { ["keep", "clear"] })), deadline.Token));
            Success(await client.CallAsync("range", "clear-formats", createdSession,
                RangeArgsOnSheet("Sheet1", "A1"), deadline.Token));
            var preserved = Success(await client.CallAsync("range", "get-values", createdSession,
                RangeArgsOnSheet("Sheet1", "A1"), deadline.Token));
            Assert.Equal("keep", preserved.GetProperty("values")[0][0].GetString());
            Success(await client.CallAsync("range", "clear-contents", createdSession,
                RangeArgsOnSheet("Sheet1", "B1"), deadline.Token));
            Success(await client.CallAsync("range", "clear-all", createdSession,
                RangeArgsOnSheet("Sheet1", "A1"), deadline.Token));
            var cleared = Success(await client.CallAsync("range", "get-values", createdSession,
                RangeArgsOnSheet("Sheet1", "A1:B1"), deadline.Token));
            Assert.All(cleared.GetProperty("values").EnumerateArray().SelectMany(row => row.EnumerateArray()),
                value => Assert.True(value.ValueKind is JsonValueKind.Null
                    || value.ValueKind == JsonValueKind.String && string.IsNullOrEmpty(value.GetString())));
            Success(await client.CallAsync("file", "close", createdSession, new(), deadline.Token));
            Success(await client.CallAsync("file", "close", createdSession, new(), deadline.Token));

            var unsupportedMacroCreate = await client.CallAsync(
                "file", "create", null,
                new() { ["path"] = createdMacro }, deadline.Token);
            Assert.False(unsupportedMacroCreate.GetProperty("success").GetBoolean());
            Assert.Equal(
                "PlatformNotSupported",
                unsupportedMacroCreate.GetProperty("errorCategory").GetString());
            Assert.Contains(
                "Excel-authored .xlsm template",
                unsupportedMacroCreate.GetProperty("errorMessage").GetString(),
                StringComparison.Ordinal);
            Assert.False(File.Exists(createdMacro));

            var sentinelSession = SessionId(await client.CallAsync("file", "open", null,
                new() { ["path"] = sentinel }, deadline.Token));
            await InitializeWorkbookAsync(
                client,
                sentinelSession,
                includeSpare: false,
                deadline.Token);
            Success(await client.CallAsync("range", "set-values", sentinelSession,
                RangeArgs("A1", ("values", SentinelValues)), deadline.Token));
            Success(await client.CallAsync("range", "set-values", mainSession,
                RangeArgs("A1:B3", ("values_file", dataFile)), deadline.Token));
            Success(await client.CallAsync("range", "set-number-format", mainSession,
                RangeArgs("B2:B3", ("format_code", "0.00")), deadline.Token));
            var numberFormats = Success(await client.CallAsync("range", "get-number-formats", mainSession,
                RangeArgs("B2:B3"), deadline.Token));
            Assert.Equal(2, numberFormats.GetProperty("rowCount").GetInt32());
            Assert.Equal(1, numberFormats.GetProperty("columnCount").GetInt32());
            Assert.All(
                numberFormats.GetProperty("formats").EnumerateArray(),
                row => Assert.Equal("0.00", row[0].GetString()));
            Success(await client.CallAsync("rangeformat", "set-column-width", mainSession,
                RangeArgs("A:B", ("column_width", 14)), deadline.Token));
            Success(await client.CallAsync("rangeformat", "set-row-height", mainSession,
                RangeArgs("1:3", ("row_height", 20)), deadline.Token));
            Success(await client.CallAsync("range", "set-formulas", mainSession,
                RangeArgs("C1", ("formulas", SumFormula)), deadline.Token));
            Success(await client.CallAsync("calculation_mode", "calculate", mainSession,
                RangeArgs("C1", ("scope", "range")), deadline.Token));

            if (Environment.GetEnvironmentVariable("EXCELMCP_MAC_RANGE_EXPANSION_E2E") == "1")
            {
                var unavailableUsedRange = await client.CallAsync(
                    "range", "get-used-range", mainSession,
                    new() { ["sheet_name"] = "Data" }, deadline.Token);
                Assert.False(unavailableUsedRange.GetProperty("success").GetBoolean());
                Assert.Equal(
                    "PlatformNotSupported",
                    unavailableUsedRange.GetProperty("errorCategory").GetString());

                Success(await client.CallAsync("range", "set-values", mainSession,
                    RangeArgs("D5:E6", ("values", new object?[][]
                    {
                        [4, null],
                        ["A deliberately long value for native auto-fit", 7]
                    })), deadline.Token));
                Success(await client.CallAsync("range", "set-formulas", mainSession,
                    RangeArgs("E5", ("formulas", RangeExpansionFormula)), deadline.Token));
                Success(await client.CallAsync("range", "set-number-formats", mainSession,
                    RangeArgs("D5:E6", ("formats", RangeExpansionFormats)), deadline.Token));
                var matrixFormats = Success(await client.CallAsync(
                    "range", "get-number-formats", mainSession, RangeArgs("D5:E6"), deadline.Token));
                Assert.Equal("0", matrixFormats.GetProperty("formats")[0][0].GetString());
                Assert.Equal("0.00", matrixFormats.GetProperty("formats")[0][1].GetString());
                Assert.Equal("@", matrixFormats.GetProperty("formats")[1][0].GetString());
                Assert.Equal("#,##0", matrixFormats.GetProperty("formats")[1][1].GetString());

                var info = Success(await client.CallAsync(
                    "range", "get-info", mainSession, RangeArgs("D5:E6"), deadline.Token));
                Assert.Equal("$D$5:$E$6", info.GetProperty("address").GetString());
                Assert.Equal(2, info.GetProperty("rowCount").GetInt32());
                Assert.Equal(2, info.GetProperty("columnCount").GetInt32());
                Assert.True(info.GetProperty("width").GetDouble() > 0);
                Assert.True(info.GetProperty("height").GetDouble() > 0);

                Success(await client.CallAsync("range", "copy", mainSession,
                    new()
                    {
                        ["source_sheet"] = "Data",
                        ["source_range"] = "D5:E6",
                        ["target_sheet"] = "Data",
                        ["target_range"] = "G5"
                    }, deadline.Token));
                var copiedFormula = Success(await client.CallAsync(
                    "range", "get-formulas", mainSession, RangeArgs("H5"), deadline.Token));
                Assert.Equal("=G5*2", copiedFormula.GetProperty("formulas")[0][0].GetString());
                var copiedFormats = Success(await client.CallAsync(
                    "range", "get-number-formats", mainSession, RangeArgs("G5:H6"), deadline.Token));
                Assert.Equal("0.00", copiedFormats.GetProperty("formats")[0][1].GetString());

                Success(await client.CallAsync("range", "copy-values", mainSession,
                    new()
                    {
                        ["source_sheet"] = "Data",
                        ["source_range"] = "D5:E6",
                        ["target_sheet"] = "Data",
                        ["target_range"] = "J5"
                    }, deadline.Token));
                var copiedValues = Success(await client.CallAsync(
                    "range", "get-values", mainSession, RangeArgs("J5:K6"), deadline.Token));
                Assert.Equal(8, copiedValues.GetProperty("values")[0][1].GetDouble());
                Assert.Equal(7, copiedValues.GetProperty("values")[1][1].GetDouble());

                Success(await client.CallAsync("range", "copy-formulas", mainSession,
                    new()
                    {
                        ["source_sheet"] = "Data",
                        ["source_range"] = "D5:E6",
                        ["target_sheet"] = "Data",
                        ["target_range"] = "M5"
                    }, deadline.Token));
                var formulasOnly = Success(await client.CallAsync(
                    "range", "get-formulas", mainSession, RangeArgs("N5"), deadline.Token));
                Assert.Equal("=M5*2", formulasOnly.GetProperty("formulas")[0][0].GetString());
                var formulasOnlyFormats = Success(await client.CallAsync(
                    "range", "get-number-formats", mainSession, RangeArgs("M5:N6"), deadline.Token));
                Assert.All(
                    formulasOnlyFormats.GetProperty("formats").EnumerateArray()
                        .SelectMany(row => row.EnumerateArray()),
                    format => Assert.Equal("General", format.GetString()));

                Success(await client.CallAsync("rangeformat", "set-column-width", mainSession,
                    RangeArgs("D:E", ("column_width", 5)), deadline.Token));
                var narrowColumns = Success(await client.CallAsync(
                    "range", "get-info", mainSession, RangeArgs("D:E"), deadline.Token));
                Success(await client.CallAsync("rangeformat", "auto-fit-columns", mainSession,
                    RangeArgs("D:E"), deadline.Token));
                var fittedColumns = Success(await client.CallAsync(
                    "range", "get-info", mainSession, RangeArgs("D:E"), deadline.Token));
                Assert.True(
                    fittedColumns.GetProperty("width").GetDouble()
                    > narrowColumns.GetProperty("width").GetDouble());

                Success(await client.CallAsync("rangeformat", "set-row-height", mainSession,
                    RangeArgs("6:6", ("row_height", 5)), deadline.Token));
                var shortRow = Success(await client.CallAsync(
                    "range", "get-info", mainSession, RangeArgs("6:6"), deadline.Token));
                Success(await client.CallAsync("rangeformat", "auto-fit-rows", mainSession,
                    RangeArgs("6:6"), deadline.Token));
                var fittedRow = Success(await client.CallAsync(
                    "range", "get-info", mainSession, RangeArgs("6:6"), deadline.Token));
                Assert.True(
                    fittedRow.GetProperty("height").GetDouble()
                    > shortRow.GetProperty("height").GetDouble());

                Success(await client.CallAsync("rangeformat", "merge-cells", mainSession,
                    RangeArgs("P5:Q5"), deadline.Token));
                var unavailableMergeInfo = await client.CallAsync(
                    "rangeformat", "get-merge-info", mainSession, RangeArgs("P5:Q6"), deadline.Token);
                Assert.False(unavailableMergeInfo.GetProperty("success").GetBoolean());
                Assert.Equal(
                    "PlatformNotSupported",
                    unavailableMergeInfo.GetProperty("errorCategory").GetString());
                Success(await client.CallAsync("rangeformat", "unmerge-cells", mainSession,
                    RangeArgs("P5:Q5"), deadline.Token));

                Success(await client.CallAsync("rangelink", "set-cell-lock", mainSession,
                    RangeArgs("R5:S5", ("locked", false)), deadline.Token));
                var lockInfo = Success(await client.CallAsync(
                    "rangelink", "get-cell-lock", mainSession, RangeArgs("R5:S5"), deadline.Token));
                Assert.False(lockInfo.GetProperty("isLocked").GetBoolean());
                Success(await client.CallAsync("rangelink", "set-cell-lock", mainSession,
                    RangeArgs("R5:S5", ("locked", true)), deadline.Token));
                lockInfo = Success(await client.CallAsync(
                    "rangelink", "get-cell-lock", mainSession, RangeArgs("R5:S5"), deadline.Token));
                Assert.True(lockInfo.GetProperty("isLocked").GetBoolean());
            }

            var values = Success(await client.CallAsync("range", "get-values", mainSession,
                RangeArgs("A1:B3"), deadline.Token));
            Assert.Equal(3, values.GetProperty("rowCount").GetInt32());
            Assert.Equal(2, values.GetProperty("columnCount").GetInt32());
            Assert.Equal("Alpha", values.GetProperty("values")[1][0].GetString());
            Assert.Equal(20, values.GetProperty("values")[2][1].GetDouble());
            var formula = Success(await client.CallAsync("range", "get-formulas", mainSession,
                RangeArgs("C1"), deadline.Token));
            Assert.Equal("=SUM(B2:B3)", formula.GetProperty("formulas")[0][0].GetString());
            Assert.Equal(30, formula.GetProperty("values")[0][0].GetDouble());
            Assert.Equal(1, formula.GetProperty("rowCount").GetInt32());
            Assert.Equal(1, formula.GetProperty("columnCount").GetInt32());

            var missing = RangeArgs("A1");
            missing["sheet_name"] = "Missing";
            var error = await client.CallAsync("range", "get-values", mainSession, missing, deadline.Token);
            Assert.False(error.GetProperty("success").GetBoolean());
            Assert.False(string.IsNullOrWhiteSpace(error.GetProperty("errorMessage").GetString()));
            Success(await client.CallAsync("range", "get-values", mainSession, RangeArgs("A1"), deadline.Token));

            Success(await client.CallAsync("file", "close", mainSession, new() { ["save"] = true }, deadline.Token));
            mainSession = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = main }, deadline.Token));
            var reopenedPowerQueryRead = await client.CallAsync(
                "powerquery", "list", mainSession, new(), deadline.Token);
            Assert.False(reopenedPowerQueryRead.GetProperty("success").GetBoolean());
            Assert.Contains(
                "no supported local macOS API",
                reopenedPowerQueryRead.GetProperty("errorMessage").GetString(),
                StringComparison.Ordinal);
            var saved = Success(await client.CallAsync("range", "get-values", mainSession, RangeArgs("C1"), deadline.Token));
            Assert.Equal(30, saved.GetProperty("values")[0][0].GetDouble());
            Success(await client.CallAsync("range", "set-values", mainSession,
                RangeArgs("B2", ("values", UnsavedValues)), deadline.Token));
            Success(await client.CallAsync("file", "close", mainSession, new(), deadline.Token));
            mainSession = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = main }, deadline.Token));
            var discarded = Success(await client.CallAsync("range", "get-values", mainSession, RangeArgs("B2"), deadline.Token));
            Assert.Equal(10, discarded.GetProperty("values")[0][0].GetDouble());
            var untouched = Success(await client.CallAsync("range", "get-values", sentinelSession, RangeArgs("A1"), deadline.Token));
            Assert.Equal(9876, untouched.GetProperty("values")[0][0].GetDouble());
            Success(await client.CallAsync("file", "close", mainSession, new(), deadline.Token));
            Success(await client.CallAsync("file", "close", sentinelSession, new(), deadline.Token));
            completed = true;
            output.WriteLine($"{entryPoint}: real open/write/formula/calculate/errors/save/reopen/discard/sentinel/close passed.");
        }
        finally
        {
            if (completed)
            {
                File.Delete(main);
                File.Delete(created);
                File.Delete(createdMacro);
                File.Delete(duplicate);
                File.Delete(sentinel);
                File.Delete(dataFile);
                directory.Delete();
                duplicateDirectory.Delete();
            }
            else
            {
                output.WriteLine($"Failed run retained opaque workbook copies at {directory.FullName}.");
            }
        }
    }

    [MacExcelTheory]
    [InlineData("cli")]
    [InlineData("mcp")]
    [Trait("Category", "Integration")]
    [Trait("RequiresExcel", "true")]
    [Trait("Feature", "MacAnalysis")]
    public async Task SpecializedAnalysis_RealEntryPointRoundTrip(string entryPoint)
    {
        Assert.Equal(0, MacAutomationAccess.Check());
        var root = FindRepository();
        var directory = Directory.CreateTempSubdirectory("excelmcp-mac-analysis-");
        var workbookPath = Path.Combine(directory.FullName, $"analysis-{Guid.NewGuid():N}.xlsx");
        CreateBlankWorkbook(workbookPath);
        using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(3));
        await using var client = await EntryPointClient.CreateAsync(root, entryPoint, output, deadline.Token);
        var completed = false;
        string? sessionId = null;
        try
        {
            sessionId = SessionId(await client.CallAsync("file", "open", null,
                new() { ["path"] = workbookPath }, deadline.Token));
            await InitializeWorkbookAsync(
                client,
                sessionId,
                includeSpare: false,
                deadline.Token);
            Success(await client.CallAsync("range", "set-values", sessionId,
                RangeArgs("F1", ("values", new object?[][] { [2] })), deadline.Token));
            Success(await client.CallAsync("range", "set-formulas", sessionId,
                RangeArgs("G1", ("formulas", new object?[][] { ["=F1*F1"] })), deadline.Token));
            var goalSeek = Success(await client.CallAsync("analysis", "goal-seek", sessionId,
                new()
                {
                    ["sheet_name"] = "Data",
                    ["formula_cell"] = "G1",
                    ["goal"] = 25,
                    ["changing_cell"] = "F1"
                }, deadline.Token));
            Assert.True(goalSeek.GetProperty("converged").GetBoolean());
            Assert.InRange(goalSeek.GetProperty("formulaValue").GetDouble(), 24.999, 25.001);
            Assert.InRange(goalSeek.GetProperty("changingValue").GetDouble(), 4.999, 5.001);

            Success(await client.CallAsync("range", "set-values", sessionId,
                RangeArgs("I1:J4", ("values", new object?[][]
                {
                    [null, null],
                    [1, null],
                    [2, null],
                    [3, null]
                })), deadline.Token));
            Success(await client.CallAsync("range", "set-formulas", sessionId,
                RangeArgs("J1", ("formulas", new object?[][] { ["=$G$1"] })), deadline.Token));
            Success(await client.CallAsync("analysis", "create-data-table", sessionId,
                new()
                {
                    ["sheet_name"] = "Data",
                    ["table_range"] = "I1:J4",
                    ["column_input_cell"] = "F1"
                }, deadline.Token));
            var dataTable = Success(await client.CallAsync("range", "get-values", sessionId,
                RangeArgs("J2:J4"), deadline.Token));
            Assert.Equal([1d, 4d, 9d], dataTable.GetProperty("values").EnumerateArray()
                .Select(row => row[0].GetDouble()).ToArray());

            Success(await client.CallAsync("file", "close", sessionId, new(), deadline.Token));
            completed = true;
            output.WriteLine($"{entryPoint}: real Goal Seek and data-table round trip passed.");
        }
        finally
        {
            if (!completed && sessionId is not null)
            {
                using var cleanupDeadline = new CancellationTokenSource(TimeSpan.FromSeconds(30));
                try
                {
                    Success(await client.CallAsync(
                        "file", "close", sessionId, new() { ["save"] = false }, cleanupDeadline.Token));
                    output.WriteLine($"Closed failed-run fixture without saving: {workbookPath}");
                }
                catch (Exception cleanupError)
                {
                    output.WriteLine(
                        $"Could not close failed-run fixture '{workbookPath}' without saving: {cleanupError.Message}");
                }
            }
            if (completed)
            {
                File.Delete(workbookPath);
                directory.Delete();
            }
            else
            {
                output.WriteLine($"Failed run retained opaque workbook copy at {workbookPath}.");
            }
        }
    }

    private static Dictionary<string, object?> RangeArgs(string address, params (string Key, object? Value)[] extras)
        => RangeArgsOnSheet("Data", address, extras);

    private static Dictionary<string, object?> RangeArgsOnSheet(
        string sheetName,
        string address,
        params (string Key, object? Value)[] extras)
    {
        var result = new Dictionary<string, object?> { ["sheet_name"] = sheetName, ["range_address"] = address };
        foreach (var (key, value) in extras) { result.Add(key, value); }
        return result;
    }

    internal static JsonElement Success(JsonElement result)
    {
        Assert.True(result.GetProperty("success").GetBoolean(), result.GetRawText());
        if (result.TryGetProperty("errorMessage", out var error))
        {
            Assert.True(error.ValueKind == JsonValueKind.Null || string.IsNullOrEmpty(error.GetString()), result.GetRawText());
        }
        return result;
    }

    internal static string SessionId(JsonElement result)
    {
        Success(result);
        var property = result.TryGetProperty("session_id", out var id) ? id : result.GetProperty("sessionId");
        Assert.False(string.IsNullOrWhiteSpace(property.GetString()));
        return property.GetString()!;
    }

    internal static string FindRepository()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "Sbroenne.ExcelMcp.sln")))
        {
            directory = directory.Parent;
        }
        return directory?.FullName ?? throw new InvalidOperationException("Repository root not found.");
    }

    internal static void CreateBlankWorkbook(string path) =>
        MacWorkbookTemplate.Copy(path, macroEnabled: false);

    internal static async Task InitializeWorkbookAsync(
        EntryPointClient client,
        string sessionId,
        bool includeSpare,
        CancellationToken cancellationToken)
    {
        Success(await client.CallAsync(
            "sheet",
            "rename",
            sessionId,
            new()
            {
                ["old_name"] = "Sheet1",
                ["new_name"] = "Data"
            },
            cancellationToken));
        if (includeSpare)
        {
            Success(await client.CallAsync(
                "sheet",
                "create",
                sessionId,
                new() { ["sheet_name"] = "Spare" },
                cancellationToken));
        }
    }

    internal sealed class EntryPointClient : IAsyncDisposable
    {
        private readonly string _cli;
        private readonly string _pipe = Environment.GetEnvironmentVariable("EXCELMCP_MAC_E2E_PIPE") ?? $"em-{Guid.NewGuid():N}";
        private readonly ITestOutputHelper _output;
        private Process? _server;
        private McpClient? _mcp;
        private Task<string>? _serverError;

        private EntryPointClient(string root, ITestOutputHelper output)
        {
            _cli = Path.Combine(root, "src", "ExcelMcp.CLI", "bin", "Release", "net10.0", "excelcli");
            _output = output;
        }

        public static async Task<EntryPointClient> CreateAsync(
            string root, string entryPoint, ITestOutputHelper output, CancellationToken cancellationToken)
        {
            var client = new EntryPointClient(root, output);
            if (entryPoint == "mcp")
            {
                var dll = Path.Combine(root, "src", "ExcelMcp.McpServer", "bin", "Release", "net10.0", "Sbroenne.ExcelMcp.McpServer.dll");
                var start = CreateStart(dll);
                start.RedirectStandardInput = true;
                client._server = Process.Start(start) ?? throw new InvalidOperationException("MCP process did not start.");
                client._serverError = client._server.StandardError.ReadToEndAsync();
                try
                {
                    client._mcp = await McpClient.CreateAsync(new StreamClientTransport(
                        client._server.StandardInput.BaseStream, client._server.StandardOutput.BaseStream),
                        cancellationToken: cancellationToken);
                }
                catch
                {
                    await client.DisposeAsync();
                    throw;
                }
            }
            return client;
        }

        public async Task<JsonElement> CallAsync(
            string tool, string action, string? sessionId, Dictionary<string, object?> args, CancellationToken cancellationToken)
        {
            string json;
            if (_mcp is not null)
            {
                args["action"] = action;
                if (sessionId is not null) { args["session_id"] = sessionId; }
                if (tool == "file" && action == "open")
                {
                    args["timeout_seconds"] =
                        UsesExtendedOpenTimeout() ? 60 : 15;
                }
                var mcpTool = tool switch
                {
                    "sheet" => "worksheet",
                    "worksheetstyle" => "worksheet_style",
                    "rangeformat" => "range_format",
                    "rangelink" => "range_link",
                    _ => tool
                };
                var result = await _mcp.CallToolAsync(mcpTool, args, cancellationToken: cancellationToken);
                json = Assert.Single(result.Content.OfType<TextContentBlock>()).Text;
            }
            else
            {
                var command = new List<string> { "-q", tool == "file" ? "session" : tool == "calculation_mode" ? "calculationmode" : tool, action };
                if (sessionId is not null) { command.AddRange(["--session", sessionId]); }
                foreach (var (key, value) in args)
                {
                    if (key == "path") { command.Add((string)value!); continue; }
                    if (key == "save")
                    {
                        if (value is true) { command.Add("--save"); }
                        continue;
                    }
                    command.Add(key switch
                    {
                        "sheet_name" => "--sheet",
                        "range_address" => "--range",
                        _ => "--" + key.Replace('_', '-')
                    });
                    command.Add(value is string text ? text : JsonSerializer.Serialize(value));
                }
                if (tool == "file" && action == "open")
                {
                    command.AddRange([
                        "--timeout",
                        UsesExtendedOpenTimeout() ? "60" : "15"
                    ]);
                }
                json = await RunCliAsync(command, cancellationToken);
            }

            using var document = JsonDocument.Parse(json);
            return document.RootElement.Clone();
        }

        public async Task TryCloseAsync(string sessionId)
        {
            using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(10));
            try
            {
                var result = await CallAsync("file", "close", sessionId, new(), deadline.Token);
                if (!result.GetProperty("success").GetBoolean())
                {
                    _output.WriteLine(
                        $"Best-effort fixture workbook cleanup failed: {result.GetRawText()}");
                }
            }
            catch (Exception exception)
            {
                _output.WriteLine(
                    $"Best-effort fixture workbook cleanup threw {exception.GetType().Name}: " +
                    exception.Message);
            }
        }

        private static bool UsesExtendedOpenTimeout() =>
            Environment.GetEnvironmentVariable("EXCELMCP_MAC_PYTHON_E2E") == "1"
            || Environment.GetEnvironmentVariable("EXCELMCP_MAC_RANGE_EXPANSION_E2E") == "1";

        private static ProcessStartInfo CreateStart(string executable)
        {
            if (!File.Exists(executable))
            {
                throw new FileNotFoundException(
                    "Mac E2E build output is missing. Run scripts/Test-MacE2E.ps1 without -SkipBuild.",
                    executable);
            }
            var managedAssembly = executable.EndsWith(".dll", StringComparison.OrdinalIgnoreCase);
            var start = new ProcessStartInfo(managedAssembly ? "dotnet" : executable)
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            if (managedAssembly) { start.ArgumentList.Add(executable); }
            if (!start.Environment.TryGetValue("DOTNET_ROOT", out var root) || string.IsNullOrWhiteSpace(root))
            {
                start.Environment["DOTNET_ROOT"] = Path.GetFullPath(Path.Combine(RuntimeEnvironment.GetRuntimeDirectory(), "../../.."));
            }
            return start;
        }

        private async Task<string> RunCliAsync(IEnumerable<string> arguments, CancellationToken cancellationToken)
        {
            using var commandDeadline = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            commandDeadline.CancelAfter(TimeSpan.FromSeconds(45));
            cancellationToken = commandDeadline.Token;
            cancellationToken.ThrowIfCancellationRequested();
            var start = CreateStart(_cli);
            start.Environment["EXCELMCP_CLI_PIPE"] = _pipe;
            foreach (var argument in arguments) { start.ArgumentList.Add(argument); }
            using var process = Process.Start(start) ?? throw new InvalidOperationException("CLI did not start.");
            var stdout = process.StandardOutput.ReadToEndAsync(cancellationToken);
            var stderr = process.StandardError.ReadToEndAsync(cancellationToken);
            try
            {
                await process.WaitForExitAsync(cancellationToken);
                await Task.WhenAll(stdout, stderr).WaitAsync(cancellationToken);
            }
            catch (OperationCanceledException)
            {
                if (!process.HasExited) { process.Kill(entireProcessTree: true); }
                await process.WaitForExitAsync();
                throw;
            }
            var error = await stderr;
            if (!string.IsNullOrWhiteSpace(error)) { _output.WriteLine(error); }
            var result = await stdout;
            Assert.False(string.IsNullOrWhiteSpace(result), $"CLI exited {process.ExitCode.ToString(CultureInfo.InvariantCulture)} with no JSON.");
            return result;
        }

        public async ValueTask DisposeAsync()
        {
            if (_server is not null)
            {
                try
                {
                    if (_mcp is not null) { await _mcp.DisposeAsync().AsTask().WaitAsync(TimeSpan.FromSeconds(5)); }
                }
                finally
                {
                    if (!_server.HasExited) { _server.Kill(entireProcessTree: true); }
                    await _server.WaitForExitAsync();
                    if (_serverError is not null) { _output.WriteLine(await _serverError); }
                    _server.Dispose();
                }
            }
            else
            {
                using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(20));
                using var result = JsonDocument.Parse(await RunCliAsync(["-q", "service", "stop"], timeout.Token));
                Success(result.RootElement);
            }
        }
    }
}
