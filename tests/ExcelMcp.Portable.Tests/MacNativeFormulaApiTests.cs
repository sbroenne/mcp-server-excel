using System.Diagnostics;
using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Portable.Tests;

[Collection("Mac Excel E2E")]
[Trait("RequiresExcel", "true")]
public sealed class MacNativeFormulaApiTests(ITestOutputHelper output)
{
    [MacExcelTheory]
    [InlineData("A1:C1", 1, 3)]
    [InlineData("A1:A3", 3, 1)]
    public Task NativeRead_ReportsTheExactVectorShape(string address, int rows, int columns) =>
        WithWorkbookAsync((path, sheet) =>
        {
            using var range = MacNativeRange.Resolve(sheet, address);
            var external = MacNativeRange.Address(range, true, TimeSpan.FromSeconds(10));
            var formula = Read(sheet, address, MacExcelDictionary.Formula2);
            var value = Read(sheet, address, MacExcelDictionary.Value2);
            var blank = MacNativeRange.Evaluate($"ISBLANK({external})", TimeSpan.FromSeconds(10));
            var error = MacNativeRange.Evaluate($"IF(ISERROR({external}),ERROR.TYPE({external}),0)", TimeSpan.FromSeconds(10));
            output.WriteLine($"Formulas: {formula}; values: {value}; blanks: {blank}; errors: {error}");
            var data = MacNativeRange.ReadData(path, "Sheet1", address,
                Core.Commands.Range.FormulaReferenceStyle.A1, TimeSpan.FromSeconds(10));
            Assert.Equal(rows, data.Values.Count);
            Assert.All(data.Values, row => Assert.Equal(columns, row!.AsArray().Count));
            return Task.CompletedTask;
        });

    [MacExcelTheory]
    [InlineData(true)]
    [InlineData(false)]
    public void NativeWorkbookNames_MatchTheNativeWorkbookCount(bool fullPaths)
    {
        Assert.Equal(0, MacAutomationAccess.Check());
        using var application = MacAppleEvents.Create(MacAppleEvents.Code("null"), []);
        var count = MacNativeRange.Count(application, MacExcelDictionary.WorkbookClass, TimeSpan.FromSeconds(10));
        output.WriteLine($"Excel reports {count} open workbooks.");
        Assert.Equal(count, MacNativeWorkbook.ReadNames(fullPaths, TimeSpan.FromSeconds(10)).Length);
    }

    [MacExcelTheory]
    [InlineData(false)]
    [InlineData(true)]
    public Task NativeFormula2_PreservesSpillsAndExplicitImplicitIntersection(bool r1c1) =>
        WithWorkbookAsync((_, sheet) =>
        {
            var formulaProperty = r1c1 ? MacExcelDictionary.Formula2R1C1 : MacExcelDictionary.Formula2;
            Write(sheet, "A1", formulaProperty, "=SEQUENCE(3)");
            Assert.Equal("=SEQUENCE(3)", Read(sheet, "A1", formulaProperty)!.GetValue<string>());
            Assert.True(JsonNode.DeepEquals(JsonNode.Parse("[[1],[2],[3]]"), Read(sheet, "A1:A3", MacExcelDictionary.Value2)));
            Write(sheet, "C1", formulaProperty, r1c1 ? "=RC[-2]:R[2]C[-2]" : "=A1:A3");
            Assert.True(JsonNode.DeepEquals(JsonNode.Parse("[[1],[2],[3]]"), Read(sheet, "C1:C3", MacExcelDictionary.Value2)));
            Write(sheet, "E1", formulaProperty, r1c1 ? "=@RC[-4]:R[2]C[-4]" : "=@A1:A3");
            Assert.Equal(r1c1 ? "=@RC[-4]:R[2]C[-4]" : "=@A1:A3", Read(sheet, "E1", formulaProperty)!.GetValue<string>());
            Assert.Equal(1, Read(sheet, "E1", MacExcelDictionary.Value2)!.GetValue<double>());
            Write(sheet, "F1", formulaProperty, r1c1 ? "=COUNTA(R[1]C[-1])" : "=COUNTA(E2)");
            Assert.Equal(0, Read(sheet, "F1", MacExcelDictionary.Value2)!.GetValue<double>());
            return Task.CompletedTask;
        });

    [MacExcelTheory]
    [InlineData(false)]
    [InlineData(true)]
    public Task NativeFormula2_MatrixPreservesRelativeReferencesAndPersistence(bool r1c1) =>
        WithWorkbookAsync(async (path, sheet) =>
        {
            var formulaProperty = r1c1 ? MacExcelDictionary.Formula2R1C1 : MacExcelDictionary.Formula2;
            string[][] formulas = r1c1
                ? [["=11", "=RC[-1]+1"], ["=R[-1]C+10", "=RC[-1]+1"]]
                : [["=11", "=A1+1"], ["=A1+10", "=A2+1"]];
            using (var matrix = MacAppleEvents.List())
            {
                foreach (var formulasRow in formulas)
                {
                    using var row = MacAppleEvents.List();
                    foreach (var formula in formulasRow)
                    {
                        using var value = MacAppleEvents.Text(formula);
                        MacAppleEvents.Append(row, value);
                    }
                    MacAppleEvents.Append(matrix, row);
                }
                Assert.True(JsonNode.DeepEquals(JsonNode.Parse(System.Text.Json.JsonSerializer.Serialize(formulas)), matrix.Decode()));
                Write(sheet, "A1:B2", formulaProperty, matrix);
            }
            var expectedFormulas = JsonNode.Parse(System.Text.Json.JsonSerializer.Serialize(formulas));
            var expectedValues = JsonNode.Parse("[[11,12],[21,22]]");
            Assert.True(JsonNode.DeepEquals(expectedFormulas, Read(sheet, "A1:B2", formulaProperty)));
            Assert.True(JsonNode.DeepEquals(expectedValues, Read(sheet, "A1:B2", MacExcelDictionary.Value2)));
            MacNativeWorkbook.Close(path, true, TimeSpan.FromSeconds(10));
            await OpenAsync(path);
            using var reopened = MacNativeWorkbook.Resolve(path, TimeSpan.FromSeconds(10));
            using var reopenedSheet = Sheet(reopened);
            Assert.True(JsonNode.DeepEquals(expectedFormulas, Read(reopenedSheet, "A1:B2", formulaProperty)));
            Assert.True(JsonNode.DeepEquals(expectedValues, Read(reopenedSheet, "A1:B2", MacExcelDictionary.Value2)));
        });

    [MacExcelTheory]
    [InlineData(false)]
    [InlineData(true)]
    public Task NativeFormula2_DistinguishesBlankFormulaFromEmptyCell(bool r1c1) =>
        WithWorkbookAsync((_, sheet) =>
        {
            var formulaProperty = r1c1 ? MacExcelDictionary.Formula2R1C1 : MacExcelDictionary.Formula2;
            Write(sheet, "A1", formulaProperty, "=\"\"");
            Assert.Equal("=\"\"", Read(sheet, "A1", formulaProperty)!.GetValue<string>());
            Assert.True(Read(sheet, "A1", MacExcelDictionary.HasFormula)!.GetValue<bool>());
            Assert.Equal("", Read(sheet, "A1", MacExcelDictionary.Value2)!.GetValue<string>());
            Assert.False(Read(sheet, "A2", MacExcelDictionary.HasFormula)!.GetValue<bool>());
            Write(sheet, "B1", formulaProperty, r1c1 ? "=COUNTA(RC[-1]:R[1]C[-1])" : "=COUNTA(A1:A2)");
            Assert.Equal(1, Read(sheet, "B1", MacExcelDictionary.Value2)!.GetValue<double>());
            return Task.CompletedTask;
        });

    [MacExcelTheory]
    [InlineData("B2:C2", "C2", 2, 2)]
    [InlineData("D5:E6", "E6", 5, 4)]
    public Task NativeMergeArea_IdentifiesTheMergedTopLeftCell(string address, string cell, int row, int column) =>
        WithWorkbookAsync((_, sheet) =>
        {
            using (var merge = MacAppleEvents.Create(MacAppleEvents.Code("bool"), [1]))
            {
                Write(sheet, address, MacExcelDictionary.MergeCells, merge);
            }
            Assert.True(Read(sheet, cell, MacExcelDictionary.MergeCells)!.GetValue<bool>());
            using var range = Range(sheet, cell);
            using var target = MacAppleEvents.Property(range, MacExcelDictionary.MergeArea);
            Assert.Equal(row, ReadProperty(target, MacExcelDictionary.FirstRowIndex)!.GetValue<int>());
            Assert.Equal(column, ReadProperty(target, MacExcelDictionary.FirstColumnIndex)!.GetValue<int>());
            return Task.CompletedTask;
        });

    [MacExcelTheory]
    [InlineData("=1/0", 2)]
    [InlineData("=NA()", 7)]
    public Task NativeRange_ReportsDimensionsAddressesAndUnambiguousErrorTypes(string formula, int expectedErrorType) =>
        WithWorkbookAsync((_, sheet) =>
        {
            using var range = MacNativeRange.Resolve(sheet, "B3:C4");
            Assert.Equal("$B$3:$C$4", MacNativeRange.Address(range, false, TimeSpan.FromSeconds(10)));
            Assert.Equal(2, MacNativeRange.Count(range, MacExcelDictionary.RowClass, TimeSpan.FromSeconds(10)));
            Assert.Equal(2, MacNativeRange.Count(range, MacExcelDictionary.ColumnClass, TimeSpan.FromSeconds(10)));
            Write(sheet, "B3", MacExcelDictionary.Formula2, formula);
            var external = MacNativeRange.Address(range, true, TimeSpan.FromSeconds(10));
            var prefix = external[..(external.LastIndexOf('!') + 1)];
            var errorType = MacNativeRange.Evaluate($"IF(ISERROR({prefix}$B$3),ERROR.TYPE({prefix}$B$3),0)", TimeSpan.FromSeconds(10));
            Assert.Equal(expectedErrorType, errorType!.GetValue<double>());
            Assert.True(MacNativeRange.Evaluate($"ISBLANK({prefix}$C$4)", TimeSpan.FromSeconds(10))!.GetValue<bool>());
            Assert.Null(Read(sheet, "B3", MacExcelDictionary.Value2));
            Assert.True(JsonNode.DeepEquals(
                JsonNode.Parse($"[[{expectedErrorType},0],[0,0]]"),
                MacNativeRange.Evaluate($"IF(ISERROR({external}),ERROR.TYPE({external}),0)", TimeSpan.FromSeconds(10))));
            Assert.True(JsonNode.DeepEquals(
                JsonNode.Parse("[[false,true],[true,true]]"),
                MacNativeRange.Evaluate($"ISBLANK({external})", TimeSpan.FromSeconds(10))));
            return Task.CompletedTask;
        });

    private async Task WithWorkbookAsync(Func<string, MacAppleEvents.Descriptor, Task> inspect)
    {
        Assert.Equal(0, MacAutomationAccess.Check());
        var directory = Directory.CreateTempSubdirectory("excelmcp-native-formula-api-");
        var path = Path.Combine(directory.FullName, $"formula-{Guid.NewGuid():N}.xlsx");
        MacWorkbookTemplate.Copy(path, macroEnabled: false);
        var completed = false;
        try
        {
            await OpenAsync(path);
            using var workbook = MacNativeWorkbook.Resolve(path, TimeSpan.FromSeconds(10));
            using var sheet = Sheet(workbook);
            await inspect(path, sheet);
            completed = true;
        }
        finally
        {
            if (completed)
            {
                MacNativeWorkbook.Close(path, false, TimeSpan.FromSeconds(10));
                File.Delete(path);
                directory.Delete();
            }
            else
            {
                output.WriteLine($"Unconfirmed formula API outcome retained its exact Excel fixture at {path}; no automatic close or retry.");
            }
        }
    }

    private static async Task OpenAsync(string path)
    {
        MacNativeWorkbook.PrepareOpen(path, TimeSpan.FromSeconds(10));
        var start = new ProcessStartInfo("/usr/bin/open");
        start.ArgumentList.Add("-g");
        start.ArgumentList.Add("-b");
        start.ArgumentList.Add("com.microsoft.Excel");
        start.ArgumentList.Add(path);
        using var process = Process.Start(start) ?? throw new InvalidOperationException("Could not open the formula API fixture.");
        using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(10));
        try
        {
            await process.WaitForExitAsync(deadline.Token);
        }
        catch (OperationCanceledException)
        {
            process.Kill(entireProcessTree: true);
            await process.WaitForExitAsync();
            throw;
        }
        Assert.Equal(0, process.ExitCode);
        MacNativeWorkbook.Attach(path, false, TimeSpan.FromSeconds(10));
    }

    private static MacAppleEvents.Descriptor Sheet(MacAppleEvents.Descriptor workbook)
    {
        using var name = MacAppleEvents.Text("Sheet1");
        return MacAppleEvents.Object(MacExcelDictionary.WorksheetClass, workbook, MacAppleEvents.Code("name"), name);
    }

    private static MacAppleEvents.Descriptor Range(MacAppleEvents.Descriptor sheet, string address)
    {
        using var name = MacAppleEvents.Text(address);
        return MacAppleEvents.Object(MacExcelDictionary.RangeClass, sheet, MacAppleEvents.Code("name"), name);
    }

    private static JsonNode? Read(MacAppleEvents.Descriptor sheet, string address, uint property)
    {
        using var range = Range(sheet, address);
        return ReadProperty(range, property);
    }

    private static JsonNode? ReadProperty(MacAppleEvents.Descriptor range, uint property)
    {
        using var target = MacAppleEvents.Property(range, property);
        using var command = MacAppleEvents.Event(MacAppleEvents.Code("core"), MacAppleEvents.Code("getd"));
        MacAppleEvents.Put(command, MacAppleEvents.Code("----"), target);
        return MacAppleEvents.Send(command, TimeSpan.FromSeconds(10));
    }

    private static void Write(MacAppleEvents.Descriptor sheet, string address, uint property, string formula)
    {
        using var value = MacAppleEvents.Text(formula);
        Write(sheet, address, property, value);
    }

    private static void Write(MacAppleEvents.Descriptor sheet, string address, uint property, MacAppleEvents.Descriptor value)
    {
        using var range = Range(sheet, address);
        using var target = MacAppleEvents.Property(range, property);
        using var command = MacAppleEvents.Event(MacAppleEvents.Code("core"), MacAppleEvents.Code("setd"));
        MacAppleEvents.Put(command, MacAppleEvents.Code("----"), target);
        MacAppleEvents.Put(command, MacAppleEvents.Code("data"), value);
        MacAppleEvents.SendCommand(command, TimeSpan.FromSeconds(10));
    }
}
