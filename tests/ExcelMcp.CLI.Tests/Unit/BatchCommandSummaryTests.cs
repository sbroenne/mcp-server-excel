using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "Batch")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class BatchCommandSummaryTests
{
    [Theory]
    [InlineData("range.get-values", """{"sheetName":"Sales Plan","rangeAddress":"A1:B2"}""", "range.get-values: 'Sales Plan'!A1:B2")]
    [InlineData("range.get-values", """{"sheetName":"O'Brien","rangeAddress":"A1"}""", "range.get-values: 'O''Brien'!A1")]
    [InlineData("range.get-values", """{"sheetName":"","rangeAddress":"SalesTotal"}""", "range.get-values: rangeAddress=SalesTotal")]
    [InlineData("range.get-values", """{"sheetName":"Sales"}""", "range.get-values: sheetName=Sales")]
    [InlineData("powerquery.refresh", """{"queryName":"SalesImport"}""", "powerquery.refresh: queryName=SalesImport")]
    [InlineData("chart.read", """{"chartName":"RevenueChart"}""", "chart.read: chartName=RevenueChart")]
    [InlineData("table.read", """{"tableName":"SalesTable"}""", "table.read: tableName=SalesTable")]
    [InlineData("pivottable.read", """{"pivotTableName":"SalesPivot"}""", "pivottable.read: pivotTableName=SalesPivot")]
    [InlineData("sheet.rename", """{"newName":"After","oldName":"Before"}""", "sheet.rename: oldName=Before, newName=After")]
    [InlineData("range.copy", """{"targetRange":"B1","targetSheet":"Archive","sourceRange":"A1:A3","sourceSheet":"Sales"}""",
        "range.copy: source=Sales!A1:A3, target=Archive!B1")]
    [InlineData("range.copy", """{"sourceRange":"SourceName","targetRange":"TargetName"}""",
        "range.copy: sourceRange=SourceName, targetRange=TargetName")]
    [InlineData("range.set-values", """{"sheetName":"Sales","rangeAddress":"A1","values":[[123]]}""",
        "range.set-values: Sales!A1 (1 row x 1 column)")]
    [InlineData("range.set-formulas", """{"sheetName":"Sales","rangeAddress":"A1:B2","formulas":[["=1",null],[true,3]]}""",
        "range.set-formulas: Sales!A1:B2 (2 rows x 2 columns)")]
    [InlineData("RANGE.SET-VALUES", """{"sheetName":"Sales","rangeAddress":"A1","values":[[123]]}""",
        "RANGE.SET-VALUES: Sales!A1 (1 row x 1 column)")]
    [InlineData("diag.ping", "{}", "diag.ping")]
    [InlineData("diag.ping", "null", "diag.ping")]
    [InlineData("diag.ping", "[]", "diag.ping")]
    [InlineData("diag.ping", "true", "diag.ping")]
    [InlineData("diag.ping", "3", "diag.ping")]
    [InlineData("range.get-values", """{"sheetName":null,"rangeAddress":["A1"],"queryName":3,"tableName":{},"chartName":false}""", "range.get-values")]
    [InlineData("range.get-values", """{"sheetName":"","rangeAddress":""}""", "range.get-values")]
    [InlineData("range.get-values", """{"SheetName":"WrongCase","RangeAddress":"A1"}""", "range.get-values")]
    public void Create_UsesOnlySupportedTargetDetails(string command, string args, string expected)
    {
        using var json = JsonDocument.Parse(args);
        Assert.Equal(expected, BatchCommandSummary.Create(command, json.RootElement));
    }

    [Fact]
    public void Create_AbsentArguments_UsesCommand()
    {
        Assert.Equal("diag.ping", BatchCommandSummary.Create("diag.ping", null));
        Assert.Equal("diag.ping", BatchCommandSummary.Create("diag.ping", default(JsonElement)));
    }

    [Theory]
    [InlineData("""{"values":[]}""")]
    [InlineData("""{"values":[[]]}""")]
    [InlineData("""{"values":[[1,2],[3]]}""")]
    [InlineData("""{"values":[[1],null]}""")]
    [InlineData("""{"values":[1,2]}""")]
    [InlineData("""{"values":"not-an-array"}""")]
    [InlineData("""{"values":null}""")]
    [InlineData("""{"valuesFile":"C:\\synthetic-private\\missing.json"}""")]
    public void Create_WithoutRectangularInlineData_DoesNotInventDimensions(string args)
    {
        using var json = JsonDocument.Parse(args);
        Assert.Equal("range.set-values", BatchCommandSummary.Create("range.set-values", json.RootElement));
    }

    [Fact]
    public void Create_DetailOrder_DoesNotDependOnInputOrder()
    {
        using var first = JsonDocument.Parse("""{"chartName":"Chart","rangeAddress":"A1","sheetName":"Sales","tableName":"Table"}""");
        using var second = JsonDocument.Parse("""{"tableName":"Table","sheetName":"Sales","rangeAddress":"A1","chartName":"Chart"}""");
        const string expected = "chart.create: Sales!A1, tableName=Table, chartName=Chart";
        Assert.Equal(expected, BatchCommandSummary.Create("chart.create", first.RootElement));
        Assert.Equal(expected, BatchCommandSummary.Create("chart.create", second.RootElement));
    }

    [Theory]
    [InlineData("diag.echo")]
    [InlineData("range.set-values")]
    [InlineData("powerquery.create")]
    [InlineData("vba.set-code")]
    [InlineData("connection.create")]
    public void Create_DoesNotCopyExcludedFieldsOrNestedArguments(string command)
    {
        using var json = JsonDocument.Parse("""
            {
              "values":[["synthetic-cell-secret"]],
              "formulas":[["=SYNTHETIC_FORMULA_SECRET()"]],
              "password":"synthetic-password-secret",
              "accessToken":"synthetic-token-secret",
              "connectionString":"Server=synthetic;Password=synthetic-connection-secret",
              "mCode":"let Secret = \"synthetic-M-secret\" in Secret",
              "code":"Sub SyntheticVbaSecret(): End Sub",
              "filePath":"C:\\synthetic-private\\workbook.xlsx",
              "message":"synthetic-message-secret",
              "nested":{"sheetName":"synthetic-nested-secret"},
              "valuesFile":"C:\\synthetic-private\\data.json"
            }
            """);
        var expected = command == "range.set-values" ? $"{command} (1 row x 1 column)" : command;
        Assert.Equal(expected, BatchCommandSummary.Create(command, json.RootElement));
    }

    [Theory]
    [InlineData(159)]
    [InlineData(160)]
    [InlineData(161)]
    [InlineData(10000)]
    public void Create_EnforcesExactLengthLimit(int length)
    {
        const string prefix = "powerquery.refresh: queryName=";
        using var json = JsonDocument.Parse(JsonSerializer.Serialize(new { queryName = new string('x', length - prefix.Length) }));
        var summary = BatchCommandSummary.Create("powerquery.refresh", json.RootElement);
        Assert.Equal(Math.Min(length, 160), summary.Length);
        Assert.Equal(length > 160 ? prefix + new string('x', 157 - prefix.Length) + "..." : prefix + new string('x', length - prefix.Length), summary);
    }

    [Fact]
    public void Create_Truncation_DoesNotSplitSurrogatePair()
    {
        const string prefix = "powerquery.refresh: queryName=";
        var name = new string('x', 156 - prefix.Length) + char.ConvertFromUtf32(0x1F600) + "tail";
        using var json = JsonDocument.Parse(JsonSerializer.Serialize(new { queryName = name }));
        var summary = BatchCommandSummary.Create("powerquery.refresh", json.RootElement);
        Assert.Equal(prefix + new string('x', 156 - prefix.Length) + "...", summary);
        Assert.Equal(159, summary.Length);
    }

    [Fact]
    public void Create_NormalizesControlsAndLineSeparators()
    {
        using var json = JsonDocument.Parse("""{"queryName":"Sales\r\n\t\u0000\u001b\u2028\u2029\u202eImport"}""");
        Assert.Equal("powerquery.refresh: queryName=Sales        Import",
            BatchCommandSummary.Create("powerquery.refresh", json.RootElement));
        Assert.Equal("diag. ping", BatchCommandSummary.Create("diag.\nping", null));
    }

    [Fact]
    public void Create_UndecodableDetail_OmitsOnlyThatDetailAndPreservesArguments()
    {
        const string args = """{"sheetName":"\uD800","rangeAddress":"A1"}""";
        using var json = JsonDocument.Parse(args);
        Assert.Equal("range.get-values: rangeAddress=A1", BatchCommandSummary.Create("range.get-values", json.RootElement));
        Assert.Equal(args, json.RootElement.GetRawText());
    }

    [Fact]
    public void Create_ValidUnicodeEscape_RemainsVisible()
    {
        using var json = JsonDocument.Parse("""{"queryName":"Sales\uD83D\uDE00"}""");
        Assert.Equal("powerquery.refresh: queryName=Sales" + char.ConvertFromUtf32(0x1F600),
            BatchCommandSummary.Create("powerquery.refresh", json.RootElement));
    }

    [Fact]
    public void Create_LongCommandWithoutDetails_IsBounded()
    {
        Assert.Equal(new string('x', 157) + "...", BatchCommandSummary.Create(new string('x', 161), null));
    }
}
