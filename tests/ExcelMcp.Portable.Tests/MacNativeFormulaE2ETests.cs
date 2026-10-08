using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.Calculation;
using Sbroenne.ExcelMcp.Service;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;
using Xunit.Abstractions;
using static Sbroenne.ExcelMcp.Portable.Tests.MacExcelE2ETests;

namespace Sbroenne.ExcelMcp.Portable.Tests;

[Collection("Mac Excel E2E")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Formulas")]
public sealed class MacNativeFormulaE2ETests(ITestOutputHelper output)
{
    private static readonly string[] FirstFormulaRow = ["=1/0", "=2007"];
    private static readonly string[][] Replacement = [["=99"]];
    private static readonly string[][] DependentFormulas = [["=1", "=A1+1"]];
    private static readonly string[][] ScopedFormulas = [["=1", "=A1+1", "=A1+2"]];
    private static readonly string[][] SheetSentinelFormula = [["=Sheet1!A1+3"]];
    private static readonly string[][] SpillFormula = [["=SEQUENCE(2,2)"]];
    private static readonly string[][] IntersectionFormula = [["=@SEQUENCE(2,2)"]];
    private static readonly string[] RejectedMergedAddresses = ["C3", "B2:C3", "A1:B2"];

    [MacExcelTheory]
    [InlineData("cli")]
    [InlineData("mcp")]
    public async Task MergedFormulaWrites_AllowOnlyTopLeftAndPreserveRejectedRangesThroughSaveReopen(string entryPoint)
    {
        Assert.Equal(0, MacAutomationAccess.Check());
        using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(2));
        await using var client = await EntryPointClient.CreateAsync(FindRepository(), entryPoint, output, deadline.Token);
        var directory = Directory.CreateTempSubdirectory("excelmcp-merged-formulas-");
        var path = Path.Combine(directory.FullName, $"merged-{Guid.NewGuid():N}.xlsx");
        CreateBlankWorkbook(path);
        string? session = null;
        var completed = false;
        try
        {
            session = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = path }, deadline.Token));
            using (var workbook = MacNativeWorkbook.Resolve(path, TimeSpan.FromSeconds(30)))
            {
                using var name = MacAppleEvents.Text("Sheet1");
                using var sheet = MacAppleEvents.Object(MacExcelDictionary.WorksheetClass, workbook, MacAppleEvents.Code("name"), name);
                using var range = MacNativeRange.Resolve(sheet, "B2:C3");
                using var merge = MacAppleEvents.Create(MacAppleEvents.Code("bool"), [1]);
                MacNativeRange.Write(range, MacExcelDictionary.MergeCells, merge, TimeSpan.FromSeconds(30));
                using var intersecting = MacNativeRange.Resolve(sheet, "A1:B2");
                output.WriteLine($"Mixed MergeCells property: {MacNativeRange.Read(intersecting, MacExcelDictionary.MergeCells, TimeSpan.FromSeconds(30))?.ToJsonString() ?? "<missing>"}");
                output.WriteLine($"Mixed range geometry: {System.Text.Json.JsonSerializer.Serialize(
                    MacNativeRange.Describe(path, "Sheet1", "A1:B2", true, TimeSpan.FromSeconds(30)))}");
            }
            Success(await client.CallAsync("range", "set-formulas", session,
                new() { ["sheet_name"] = "Sheet1", ["range_address"] = "B2", ["formulas"] = Replacement }, deadline.Token));
            Success(await client.CallAsync("calculation_mode", "calculate", session,
                new() { ["scope"] = "sheet", ["sheet_name"] = "Sheet1" }, deadline.Token));
            var before = await ReadMergedState();
            foreach (var address in RejectedMergedAddresses)
            {
                var rejected = await client.CallAsync("range", "set-formulas", session,
                    new()
                    {
                        ["sheet_name"] = "Sheet1",
                        ["range_address"] = address,
                        ["formulas"] = Replacement,
                        ["overwrite_policy"] = "allow"
                    }, deadline.Token);
                Assert.False(rejected.GetProperty("success").GetBoolean());
                Assert.True(rejected.GetProperty("errorCategory").GetString() == "Conflict",
                    $"{address}: {rejected.GetRawText()}");
                Assert.Contains($"Cannot write to range '{address}' because the write intersects merged cells.",
                    rejected.GetProperty("errorMessage").GetString(), StringComparison.Ordinal);
                Assert.Contains("Merged range: $B$2:$C$3.", rejected.GetProperty("errorMessage").GetString(), StringComparison.Ordinal);
                Assert.Equal(before.GetRawText(), (await ReadMergedState()).GetRawText());
            }
            Success(await client.CallAsync("file", "close", session, new() { ["save"] = true }, deadline.Token));
            session = null;
            session = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = path }, deadline.Token));
            Assert.Equal(before.GetRawText(), (await ReadMergedState()).GetRawText());
            Success(await client.CallAsync("file", "close", session, new(), deadline.Token));
            session = null;
            completed = true;
        }
        finally
        {
            if (session is not null) await client.TryCloseAsync(session);
            if (completed)
            {
                File.Delete(path);
                directory.Delete();
            }
            else output.WriteLine($"Failed merged formula acceptance retained its opaque fixture at {directory.FullName}.");
        }

        async Task<JsonElement> ReadMergedState()
        {
            using var workbook = MacNativeWorkbook.Resolve(path, TimeSpan.FromSeconds(30));
            using var name = MacAppleEvents.Text("Sheet1");
            using var sheet = MacAppleEvents.Object(MacExcelDictionary.WorksheetClass, workbook, MacAppleEvents.Code("name"), name);
            using var cell = MacNativeRange.Resolve(sheet, "C3");
            Assert.True(MacNativeRange.Read(cell, MacExcelDictionary.MergeCells, TimeSpan.FromSeconds(30))!.GetValue<bool>());
            using var merge = MacAppleEvents.Property(cell, MacExcelDictionary.MergeArea);
            Assert.Equal("$B$2:$C$3", MacNativeRange.Address(merge, false, TimeSpan.FromSeconds(30)));
            var read = Success(await client.CallAsync("range", "get-formulas", session,
                new() { ["sheet_name"] = "Sheet1", ["range_address"] = "A1:C3" }, deadline.Token));
            Assert.Equal("=99", read.GetProperty("formulas")[1][1].GetString());
            Assert.Equal(99, read.GetProperty("values")[1][1].GetDouble());
            Assert.Equal(JsonValueKind.Null, read.GetProperty("values")[0][0].ValueKind);
            Assert.Empty(read.GetProperty("cellErrors").EnumerateArray());
            return read;
        }
    }

    [MacExcelTheory]
    [InlineData("cli", MacExcelDictionary.CalculationAutomatic)]
    [InlineData("mcp", MacExcelDictionary.CalculationAutomatic)]
    [InlineData("cli", MacExcelDictionary.CalculationManual)]
    [InlineData("mcp", MacExcelDictionary.CalculationManual)]
    [InlineData("cli", MacExcelDictionary.CalculationSemiautomatic)]
    [InlineData("mcp", MacExcelDictionary.CalculationSemiautomatic)]
    public async Task FormulaWrites_PreserveNativeCalculationModeAndStoredValues(string entryPoint, uint mode)
    {
        Assert.Equal(0, MacAutomationAccess.Check());
        using var application = MacAppleEvents.Create(MacAppleEvents.Code("null"), []);
        uint? original = null;
        using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(2));
        await using var client = await EntryPointClient.CreateAsync(FindRepository(), entryPoint, output, deadline.Token);
        var directory = Directory.CreateTempSubdirectory("excelmcp-formula-mode-");
        var path = Path.Combine(directory.FullName, $"mode-{Guid.NewGuid():N}.xlsx");
        CreateBlankWorkbook(path);
        string? session = null;
        var completed = false;
        try
        {
            session = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = path, ["show"] = true }, deadline.Token));
            var originalMode = ReadMode();
            original = originalMode;
            WriteMode(application, mode);
            Success(await client.CallAsync("range", "set-formulas", session, new()
            {
                ["sheet_name"] = "Sheet1",
                ["range_address"] = "A1:B1",
                ["formulas"] = DependentFormulas
            }, deadline.Token));
            Assert.Equal(mode, ReadMode());
            var read = Success(await client.CallAsync("range", "get-formulas", session,
                new() { ["sheet_name"] = "Sheet1", ["range_address"] = "A1:B1" }, deadline.Token));
            Assert.Equal(mode, ReadMode());
            Assert.Equal("=A1+1", read.GetProperty("formulas")[0][1].GetString());
            using var workbook = MacNativeWorkbook.Resolve(path, TimeSpan.FromSeconds(10));
            using var sheetName = MacAppleEvents.Text("Sheet1");
            using var sheet = MacAppleEvents.Object(MacExcelDictionary.WorksheetClass, workbook, MacAppleEvents.Code("name"), sheetName);
            using var range = MacNativeRange.Resolve(sheet, "A1:B1");
            var stored = MacNativeRange.Matrix(MacNativeRange.Read(range, MacExcelDictionary.Value2, TimeSpan.FromSeconds(10)), 1, 2);
            Assert.Equal(stored[0]![0]!.GetValue<double>(), read.GetProperty("values")[0][0].GetDouble());
            Assert.Equal(stored[0]![1]!.GetValue<double>(), read.GetProperty("values")[0][1].GetDouble());
            var rejected = await client.CallAsync("range", "set-formulas", session,
                new() { ["sheet_name"] = "Sheet1", ["range_address"] = "A1", ["formulas"] = Replacement }, deadline.Token);
            Assert.False(rejected.GetProperty("success").GetBoolean());
            Assert.Equal("Conflict", rejected.GetProperty("errorCategory").GetString());
            Assert.Equal(mode, ReadMode());
            WriteMode(application, originalMode);
            Success(await client.CallAsync("file", "close", session, new(), deadline.Token));
            session = null;
            completed = true;
        }
        finally
        {
            if (original is { } restoreMode) WriteMode(application, restoreMode);
            if (session is not null) await client.TryCloseAsync(session);
            if (completed)
            {
                File.Delete(path);
                directory.Delete();
            }
            else output.WriteLine($"Failed calculation-mode acceptance retained its opaque fixture at {directory.FullName}.");
        }
    }

    [MacExcelTheory]
    [InlineData("cli", "range")]
    [InlineData("mcp", "range")]
    [InlineData("cli", "sheet")]
    [InlineData("mcp", "sheet")]
    public Task ExplicitCalculation_UpdatesOnlyRequestedScopeWithoutChangingMode(string entryPoint, string scope) =>
        VerifyScopedCalculation(entryPoint, scope, nativePrimitive: false);

    [MacExcelTheory]
    [InlineData("range")]
    [InlineData("sheet")]
    public Task NativeCalculationPrimitive_UpdatesOnlyRequestedScopeWithoutChangingMode(string scope) =>
        VerifyScopedCalculation("mcp", scope, nativePrimitive: true);

    private async Task VerifyScopedCalculation(string entryPoint, string scope, bool nativePrimitive)
    {
        Assert.Equal(0, MacAutomationAccess.Check());
        using var application = MacAppleEvents.Create(MacAppleEvents.Code("null"), []);
        uint? original = null;
        using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(2));
        await using var client = await EntryPointClient.CreateAsync(FindRepository(), entryPoint, output, deadline.Token);
        var directory = Directory.CreateTempSubdirectory("excelmcp-native-calculate-");
        var path = Path.Combine(directory.FullName, $"calculate-{Guid.NewGuid():N}.xlsx");
        var otherPath = Path.Combine(directory.FullName, $"unrelated-{Guid.NewGuid():N}.xlsx");
        CreateBlankWorkbook(path);
        CreateBlankWorkbook(otherPath);
        string? session = null;
        string? otherSession = null;
        var completed = false;
        try
        {
            session = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = path, ["show"] = true }, deadline.Token));
            var originalMode = ReadMode();
            original = originalMode;
            WriteMode(application, MacExcelDictionary.CalculationAutomatic);
            Success(await client.CallAsync("range", "set-formulas", session, new()
            {
                ["sheet_name"] = "Sheet1",
                ["range_address"] = "A1:C1",
                ["formulas"] = ScopedFormulas
            }, deadline.Token));
            Success(await client.CallAsync("sheet", "create", session,
                new() { ["sheet_name"] = "Sentinel" }, deadline.Token));
            Success(await client.CallAsync("range", "set-formulas", session,
                new() { ["sheet_name"] = "Sentinel", ["range_address"] = "A1", ["formulas"] = SheetSentinelFormula },
                deadline.Token));
            otherSession = SessionId(await client.CallAsync("file", "open", null,
                new() { ["path"] = otherPath, ["show"] = true }, deadline.Token));
            Success(await client.CallAsync("range", "set-formulas", otherSession,
                new() { ["sheet_name"] = "Sheet1", ["range_address"] = "A1:B1", ["formulas"] = DependentFormulas }, deadline.Token));
            WriteMode(application, MacExcelDictionary.CalculationManual);
            Success(await client.CallAsync("range", "set-formulas", session,
                new() { ["sheet_name"] = "Sheet1", ["range_address"] = "A1", ["formulas"] = Replacement, ["overwrite_policy"] = "allow" },
                deadline.Token));
            Success(await client.CallAsync("range", "set-formulas", otherSession,
                new() { ["sheet_name"] = "Sheet1", ["range_address"] = "A1", ["formulas"] = Replacement, ["overwrite_policy"] = "allow" },
                deadline.Token));
            var before = Success(await client.CallAsync("range", "get-formulas", session,
                new() { ["sheet_name"] = "Sheet1", ["range_address"] = "B1" }, deadline.Token));
            Assert.Equal(2, before.GetProperty("values")[0][0].GetDouble());
            await AssertValue(session, "Sheet1", "C1", 3);
            await AssertValue(session, "Sentinel", "A1", 4);
            await AssertValue(otherSession, "Sheet1", "B1", 2);
            var activeSheet = ReadActiveSheetName(path);
            var arguments = new Dictionary<string, object?> { ["scope"] = scope, ["sheet_name"] = "Sheet1" };
            if (scope == "range") arguments["range_address"] = "B1";
            var calculated = nativePrimitive
                ? Success(JsonSerializer.SerializeToElement(MacNativeCalculation.Calculate(path,
                    scope == "range" ? CalculationScope.Range : CalculationScope.Sheet, "Sheet1",
                    scope == "range" ? "B1" : null, TimeSpan.FromSeconds(30)), ServiceProtocol.JsonOptions))
                : Success(await client.CallAsync("calculation_mode", "calculate", session, arguments, deadline.Token));
            Assert.Equal(path, calculated.GetProperty("filePath").GetString());
            Assert.Equal("calculate", calculated.GetProperty("action").GetString());
            Assert.Equal($"Normal calculation completed for {(scope == "range" ? "Range" : "Sheet")}; asynchronous refresh/Python completion is not established.",
                calculated.GetProperty("message").GetString());
            Assert.Equal(MacExcelDictionary.CalculationManual, ReadMode());
            Assert.Equal(activeSheet, ReadActiveSheetName(path));
            var after = Success(await client.CallAsync("range", "get-formulas", session,
                new() { ["sheet_name"] = "Sheet1", ["range_address"] = "B1" }, deadline.Token));
            Assert.Equal(100, after.GetProperty("values")[0][0].GetDouble());
            await AssertValue(session, "Sheet1", "C1", scope == "range" ? 3 : 101);
            await AssertValue(session, "Sentinel", "A1", 4);
            await AssertValue(otherSession, "Sheet1", "B1", 2);
            if (nativePrimitive)
            {
                var unsupported = Assert.Throws<PlatformNotSupportedException>(() =>
                    MacNativeCalculation.Calculate(path, CalculationScope.Application, "", null, TimeSpan.FromSeconds(30)));
                Assert.Contains("no calculation was attempted", unsupported.Message, StringComparison.Ordinal);
            }
            else
            {
                var unsupported = await client.CallAsync("calculation_mode", "calculate", session,
                    new() { ["scope"] = "application" }, deadline.Token);
                Assert.False(unsupported.GetProperty("success").GetBoolean());
                Assert.Equal("PlatformNotSupported", unsupported.GetProperty("errorCategory").GetString());
            }
            Assert.Equal(MacExcelDictionary.CalculationManual, ReadMode());
            await AssertValue(session, "Sentinel", "A1", 4);
            await AssertValue(otherSession, "Sheet1", "B1", 2);
            WriteMode(application, originalMode);
            Success(await client.CallAsync("file", "close", otherSession, new(), deadline.Token));
            otherSession = null;
            Success(await client.CallAsync("file", "close", session, new(), deadline.Token));
            session = null;
            completed = true;
        }
        finally
        {
            if (original is { } restoreMode) WriteMode(application, restoreMode);
            if (otherSession is not null) await client.TryCloseAsync(otherSession);
            if (session is not null) await client.TryCloseAsync(session);
            if (completed)
            {
                File.Delete(path);
                File.Delete(otherPath);
                directory.Delete();
            }
            else output.WriteLine($"Failed scoped-calculation acceptance retained its opaque fixture at {directory.FullName}.");
        }

        async Task AssertValue(string sessionId, string sheetName, string address, double expected)
        {
            var result = Success(await client.CallAsync("range", "get-formulas", sessionId,
                new() { ["sheet_name"] = sheetName, ["range_address"] = address }, deadline.Token));
            Assert.Equal(expected, result.GetProperty("values")[0][0].GetDouble());
        }
    }

    private static string ReadActiveSheetName(string path)
    {
        using var workbook = MacNativeWorkbook.Resolve(path, TimeSpan.FromSeconds(30));
        using var activeSheet = MacAppleEvents.Property(workbook, MacExcelDictionary.ActiveSheet);
        return MacNativeRange.Read(activeSheet, MacExcelDictionary.Name, TimeSpan.FromSeconds(30))
            ?.GetValue<string>() ?? throw new InvalidDataException("Excel returned a missing active worksheet name.");
    }

    private static uint ReadMode()
    {
        using var application = MacAppleEvents.Create(MacAppleEvents.Code("null"), []);
        var mode = MacNativeRange.Read(application, MacExcelDictionary.Calculation, TimeSpan.FromSeconds(30))
            ?.GetValue<uint>() ?? throw new InvalidDataException("Excel returned a missing calculation mode.");
        return mode switch
        {
            MacExcelDictionary.CalculationAutomatic or MacExcelDictionary.CalculationManual
                or MacExcelDictionary.CalculationSemiautomatic => mode,
            _ => throw new InvalidDataException($"Excel returned an unsupported calculation mode: {mode}")
        };
    }

    private static void WriteMode(MacAppleEvents.Descriptor application, uint mode)
    {
        using var value = MacAppleEvents.Create(MacAppleEvents.Code("enum"), BitConverter.GetBytes(mode));
        MacNativeRange.Write(application, MacExcelDictionary.Calculation, value, TimeSpan.FromSeconds(10));
    }

    [MacExcelTheory]
    [InlineData("cli", "a1", false)]
    [InlineData("mcp", "a1", false)]
    [InlineData("cli", "r1c1", false)]
    [InlineData("mcp", "r1c1", false)]
    [InlineData("cli", "a1", true)]
    [InlineData("mcp", "a1", true)]
    [InlineData("cli", "r1c1", true)]
    [InlineData("mcp", "r1c1", true)]
    public async Task FormulaMatrix_ReturnsSharedFieldsErrorsAndPersistence(string entryPoint, string style, bool useFile)
    {
        Assert.Equal(0, MacAutomationAccess.Check());
        using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(2));
        await using var client = await EntryPointClient.CreateAsync(FindRepository(), entryPoint, output, deadline.Token);
        var directory = Directory.CreateTempSubdirectory("excelmcp-native-formulas-");
        var path = Path.Combine(directory.FullName, $"formulas-{Guid.NewGuid():N}.xlsx");
        var formulasPath = Path.Combine(directory.FullName, "formulas.json");
        CreateBlankWorkbook(path);
        string? session = null;
        var completed = false;
        try
        {
            session = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = path }, deadline.Token));
            var formulas = new[] { FirstFormulaRow, new[] { "=\"\"", style == "r1c1" ? "=R[-1]C+1" : "=C3+1" } };
            var arguments = new Dictionary<string, object?>
            {
                ["sheet_name"] = "Sheet1",
                ["range_address"] = "B3:C4",
                ["reference_style"] = style
            };
            if (useFile)
            {
                await File.WriteAllTextAsync(formulasPath, JsonSerializer.Serialize(formulas), deadline.Token);
                arguments["formulas_file"] = formulasPath;
            }
            else arguments["formulas"] = formulas;
            var written = Success(await client.CallAsync("range", "set-formulas", session, arguments, deadline.Token));
            Assert.Equal(path, written.GetProperty("filePath").GetString());
            Assert.Equal("set-formulas", written.GetProperty("action").GetString());
            var read = Success(await client.CallAsync("range", "get-formulas", session,
                new() { ["sheet_name"] = "Sheet1", ["range_address"] = "B3:C4", ["reference_style"] = style }, deadline.Token));
            Assert.Equal(path, read.GetProperty("filePath").GetString());
            Assert.Equal("Sheet1", read.GetProperty("sheetName").GetString());
            Assert.Equal("$B$3:$C$4", read.GetProperty("rangeAddress").GetString());
            Assert.Equal(2, read.GetProperty("rowCount").GetInt32());
            Assert.Equal(2, read.GetProperty("columnCount").GetInt32());
            Assert.Equal("#DIV/0!", read.GetProperty("values")[0][0].GetString());
            Assert.Equal(2007, read.GetProperty("values")[0][1].GetDouble());
            Assert.Equal("", read.GetProperty("values")[1][0].GetString());
            Assert.Equal(2008, read.GetProperty("values")[1][1].GetDouble());
            Assert.Equal(style == "r1c1" ? "=R[-1]C+1" : "=C3+1", read.GetProperty("formulas")[1][1].GetString());
            var error = Assert.Single(read.GetProperty("cellErrors").EnumerateArray());
            Assert.Equal("B3", error.GetProperty("cellAddress").GetString());
            Assert.Equal(-2146826281, error.GetProperty("errorCode").GetInt32());
            Assert.Equal(-2146826281, error.GetProperty("currentValue").GetInt32());
            Assert.Equal("=1/0", error.GetProperty("formula").GetString());
            Assert.Equal("Ensure the formula does not divide by zero.", error.GetProperty("suggestion").GetString());
            var wrongShape = await client.CallAsync("range", "set-formulas", session,
                new() { ["sheet_name"] = "Sheet1", ["range_address"] = "B3:C4", ["formulas"] = Replacement }, deadline.Token);
            Assert.False(wrongShape.GetProperty("success").GetBoolean());
            Assert.Equal("InvalidInput", wrongShape.GetProperty("errorCategory").GetString());
            Assert.Contains("Formula array row 1 column count (1) doesn't match range column count (2)",
                wrongShape.GetProperty("errorMessage").GetString(), StringComparison.Ordinal);
            var rejected = await client.CallAsync("range", "set-formulas", session,
                new() { ["sheet_name"] = "Sheet1", ["range_address"] = "B4", ["formulas"] = Replacement }, deadline.Token);
            Assert.False(rejected.GetProperty("success").GetBoolean());
            Assert.Equal("Conflict", rejected.GetProperty("errorCategory").GetString());
            Assert.Contains("$B$4", rejected.GetProperty("errorMessage").GetString(), StringComparison.Ordinal);
            var unchanged = Success(await client.CallAsync("range", "get-formulas", session,
                new() { ["sheet_name"] = "Sheet1", ["range_address"] = "B3:C4", ["reference_style"] = style }, deadline.Token));
            Assert.Equal(read.GetProperty("formulas").GetRawText(), unchanged.GetProperty("formulas").GetRawText());
            Success(await client.CallAsync("file", "close", session, new() { ["save"] = true }, deadline.Token));
            session = null;
            session = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = path }, deadline.Token));
            var reopened = Success(await client.CallAsync("range", "get-formulas", session,
                new() { ["sheet_name"] = "Sheet1", ["range_address"] = "B3:C4", ["reference_style"] = style }, deadline.Token));
            Assert.Equal(read.GetProperty("formulas").GetRawText(), reopened.GetProperty("formulas").GetRawText());
            Assert.Equal(read.GetProperty("values").GetRawText(), reopened.GetProperty("values").GetRawText());
            Assert.Equal(read.GetProperty("cellErrors").GetRawText(), reopened.GetProperty("cellErrors").GetRawText());
            Success(await client.CallAsync("file", "close", session, new(), deadline.Token));
            session = null;
            completed = true;
        }
        finally
        {
            if (session is not null) await client.TryCloseAsync(session);
            if (completed)
            {
                File.Delete(path);
                if (useFile) File.Delete(formulasPath);
                directory.Delete();
            }
            else output.WriteLine($"Failed native formula acceptance retained its opaque fixture at {directory.FullName}.");
        }
    }

    [MacExcelTheory]
    [InlineData("cli", false)]
    [InlineData("mcp", false)]
    [InlineData("cli", true)]
    [InlineData("mcp", true)]
    public async Task MixedFormulaCells_PreserveFormulasAndConstantsThroughSaveReopen(string entryPoint, bool useFile)
    {
        Assert.Equal(0, MacAutomationAccess.Check());
        using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(2));
        await using var client = await EntryPointClient.CreateAsync(FindRepository(), entryPoint, output, deadline.Token);
        var directory = Directory.CreateTempSubdirectory("excelmcp-mixed-formulas-");
        var path = Path.Combine(directory.FullName, $"mixed-{Guid.NewGuid():N}.xlsx");
        var formulasPath = Path.Combine(directory.FullName, "formulas.json");
        object?[][] cells = [["=1+1", "Label", 5.86, true, null]];
        CreateBlankWorkbook(path);
        string? session = null;
        var completed = false;
        try
        {
            session = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = path }, deadline.Token));
            var arguments = new Dictionary<string, object?>
            {
                ["sheet_name"] = "Sheet1",
                ["range_address"] = "A1:E1"
            };
            if (useFile)
            {
                await File.WriteAllTextAsync(formulasPath, JsonSerializer.Serialize(cells), deadline.Token);
                arguments["formulas_file"] = formulasPath;
            }
            else
            {
                arguments["formulas"] = cells;
            }

            Success(await client.CallAsync("range", "set-formulas", session, arguments, deadline.Token));
            var read = Success(await client.CallAsync("range", "get-formulas", session,
                new() { ["sheet_name"] = "Sheet1", ["range_address"] = "A1:E1" }, deadline.Token));
            Assert.Equal("=1+1", read.GetProperty("formulas")[0][0].GetString());
            for (var column = 1; column < 5; column++)
                Assert.Equal(string.Empty, read.GetProperty("formulas")[0][column].GetString());
            Assert.Equal(2, read.GetProperty("values")[0][0].GetDouble());
            Assert.Equal("Label", read.GetProperty("values")[0][1].GetString());
            Assert.Equal(5.86, read.GetProperty("values")[0][2].GetDouble(), 2);
            Assert.True(read.GetProperty("values")[0][3].GetBoolean());
            Assert.Equal(JsonValueKind.Null, read.GetProperty("values")[0][4].ValueKind);
            Success(await client.CallAsync("file", "close", session, new() { ["save"] = true }, deadline.Token));
            session = null;
            session = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = path }, deadline.Token));
            var reopened = Success(await client.CallAsync("range", "get-formulas", session,
                new() { ["sheet_name"] = "Sheet1", ["range_address"] = "A1:E1" }, deadline.Token));
            Assert.Equal(read.GetProperty("formulas").GetRawText(), reopened.GetProperty("formulas").GetRawText());
            Assert.Equal(read.GetProperty("values").GetRawText(), reopened.GetProperty("values").GetRawText());
            Success(await client.CallAsync("file", "close", session, new(), deadline.Token));
            session = null;
            completed = true;
        }
        finally
        {
            if (session is not null) await client.TryCloseAsync(session);
            if (completed)
            {
                File.Delete(path);
                if (useFile) File.Delete(formulasPath);
                directory.Delete();
            }
            else output.WriteLine($"Failed mixed formula acceptance retained its opaque fixture at {directory.FullName}.");
        }
    }

    [MacExcelTheory]
    [InlineData("cli")]
    [InlineData("mcp")]
    public async Task Formula2_PreservesSpillAndExplicitIntersectionThroughSaveReopen(string entryPoint)
    {
        Assert.Equal(0, MacAutomationAccess.Check());
        using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(2));
        await using var client = await EntryPointClient.CreateAsync(FindRepository(), entryPoint, output, deadline.Token);
        var directory = Directory.CreateTempSubdirectory("excelmcp-formula2-");
        var path = Path.Combine(directory.FullName, $"formula2-{Guid.NewGuid():N}.xlsx");
        CreateBlankWorkbook(path);
        string? session = null;
        var completed = false;
        try
        {
            session = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = path }, deadline.Token));
            Success(await client.CallAsync("range", "set-formulas", session,
                new() { ["sheet_name"] = "Sheet1", ["range_address"] = "A1", ["formulas"] = SpillFormula }, deadline.Token));
            Success(await client.CallAsync("range", "set-formulas", session,
                new() { ["sheet_name"] = "Sheet1", ["range_address"] = "D1", ["formulas"] = IntersectionFormula }, deadline.Token));
            Success(await client.CallAsync("calculation_mode", "calculate", session,
                new() { ["scope"] = "sheet", ["sheet_name"] = "Sheet1" }, deadline.Token));
            var before = await ReadVariants();
            var rejected = await client.CallAsync("range", "set-formulas", session,
                new() { ["sheet_name"] = "Sheet1", ["range_address"] = "B1", ["formulas"] = Replacement }, deadline.Token);
            Assert.False(rejected.GetProperty("success").GetBoolean());
            Assert.Equal("Conflict", rejected.GetProperty("errorCategory").GetString());
            Assert.Contains("$B$1", rejected.GetProperty("errorMessage").GetString(), StringComparison.Ordinal);
            Assert.Equal(before.GetRawText(), (await ReadVariants()).GetRawText());
            Success(await client.CallAsync("file", "close", session, new() { ["save"] = true }, deadline.Token));
            session = null;
            session = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = path }, deadline.Token));
            Assert.Equal(before.GetRawText(), (await ReadVariants()).GetRawText());
            Success(await client.CallAsync("file", "close", session, new(), deadline.Token));
            session = null;
            completed = true;
        }
        finally
        {
            if (session is not null) await client.TryCloseAsync(session);
            if (completed)
            {
                File.Delete(path);
                directory.Delete();
            }
            else output.WriteLine($"Failed public Formula2 acceptance retained its opaque fixture at {directory.FullName}.");
        }

        async Task<JsonElement> ReadVariants()
        {
            var read = Success(await client.CallAsync("range", "get-formulas", session,
                new() { ["sheet_name"] = "Sheet1", ["range_address"] = "A1:D2" }, deadline.Token));
            Assert.Equal(path, read.GetProperty("filePath").GetString());
            Assert.Equal("$A$1:$D$2", read.GetProperty("rangeAddress").GetString());
            Assert.Equal(2, read.GetProperty("rowCount").GetInt32());
            Assert.Equal(4, read.GetProperty("columnCount").GetInt32());
            Assert.Equal("=SEQUENCE(2,2)", read.GetProperty("formulas")[0][0].GetString());
            Assert.Equal("=@SEQUENCE(2,2)", read.GetProperty("formulas")[0][3].GetString());
            var values = read.GetProperty("values");
            Assert.Equal(1, values[0][0].GetDouble());
            Assert.Equal(2, values[0][1].GetDouble());
            Assert.Equal(3, values[1][0].GetDouble());
            Assert.Equal(4, values[1][1].GetDouble());
            Assert.Equal(1, values[0][3].GetDouble());
            Assert.Equal(JsonValueKind.Null, values[1][3].ValueKind);
            Assert.Empty(read.GetProperty("cellErrors").EnumerateArray());
            return read;
        }
    }
}
