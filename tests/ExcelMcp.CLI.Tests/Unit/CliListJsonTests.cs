using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "CLI")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
[Collection("Sequential")]
public sealed class CliListJsonTests
{
    [Theory]
    [InlineData("columnIndex", "ascending")]
    [InlineData("ColumnIndex", "Ascending")]
    public async Task RangeSort_PreservesColumnOrderAndDescending(string columnKey, string ascendingKey)
    {
        var columns = $$"""
            [{"{{columnKey}}":3,"{{ascendingKey}}":false},
             {"{{columnKey}}":1,"{{ascendingKey}}":true}]
            """;
        using var captured = await CaptureArgumentsAsync(
            ["rangeedit", "sort", "--session", "owned", "--sheet-name", "Sheet1",
             "--range-address", "A1:C4", "--sort-columns", columns],
            "rangeedit.sort");

        var root = captured.RootElement;
        Assert.Equal("Sheet1", root.GetProperty("sheetName").GetString());
        Assert.Equal("A1:C4", root.GetProperty("rangeAddress").GetString());
        var values = root.GetProperty("sortColumns").EnumerateArray().ToArray();
        Assert.Equal(2, values.Length);
        Assert.Equal(3, values[0].GetProperty("columnIndex").GetInt32());
        Assert.False(values[0].GetProperty("ascending").GetBoolean());
        Assert.Equal(1, values[1].GetProperty("columnIndex").GetInt32());
        Assert.True(values[1].GetProperty("ascending").GetBoolean());
    }

    [Theory]
    [InlineData("columnName", "ascending")]
    [InlineData("ColumnName", "Ascending")]
    public async Task TableSort_PreservesNamesOrderAndDescending(string columnKey, string ascendingKey)
    {
        var columns = $$"""
            [{"{{columnKey}}":"Second value","{{ascendingKey}}":false},
             {"{{columnKey}}":"First","{{ascendingKey}}":true}]
            """;
        using var captured = await CaptureArgumentsAsync(
            ["tablecolumn", "sort-multi", "--session", "owned", "--table-name", "Records", "--sort-columns", columns],
            "tablecolumn.sort-multi");

        var root = captured.RootElement;
        Assert.Equal("Records", root.GetProperty("tableName").GetString());
        var values = root.GetProperty("sortColumns").EnumerateArray().ToArray();
        Assert.Equal(2, values.Length);
        Assert.Equal("Second value", values[0].GetProperty("columnName").GetString());
        Assert.False(values[0].GetProperty("ascending").GetBoolean());
        Assert.Equal("First", values[1].GetProperty("columnName").GetString());
        Assert.True(values[1].GetProperty("ascending").GetBoolean());
    }

    [Fact]
    public async Task RemoveDuplicates_PreservesIntegerKeysAndOrder()
    {
        using var captured = await CaptureArgumentsAsync(
            ["rangeedit", "remove-duplicates", "--session", "owned", "--sheet-name", "Sheet1",
             "--range-address", "A1:C4", "--key-columns", "[3,1]"],
            "rangeedit.remove-duplicates");

        var keys = captured.RootElement.GetProperty("keyColumns").EnumerateArray().ToArray();
        Assert.Equal(2, keys.Length);
        Assert.Equal(JsonValueKind.Number, keys[0].ValueKind);
        Assert.Equal(3, keys[0].GetInt32());
        Assert.Equal(JsonValueKind.Number, keys[1].ValueKind);
        Assert.Equal(1, keys[1].GetInt32());
    }

    [Fact]
    public async Task GroupObjects_PreservesExactStringItemsAndOrder()
    {
        string[] names = ["Shape Ω", "shape \"quoted\""];
        using var captured = await CaptureArgumentsAsync(
            ["drawing", "group-objects", "--session", "owned", "--sheet-name", "Sheet1",
             "--object-names", JsonSerializer.Serialize(names)],
            "drawing.group-objects");

        var values = captured.RootElement.GetProperty("objectNames").EnumerateArray().ToArray();
        Assert.Equal(names.Length, values.Length);
        for (int index = 0; index < names.Length; index++)
        {
            Assert.Equal(JsonValueKind.String, values[index].ValueKind);
            Assert.Equal(names[index], values[index].GetString());
        }
    }

    private static async Task<JsonDocument> CaptureArgumentsAsync(string[] arguments, string expectedCommand)
    {
        ServiceRequest? captured = null;
        int requests = 0;
        var result = await InProcessCliHelper.RunAsync(arguments, request =>
        {
            requests++;
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });

        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.Equal(1, requests);
        Assert.NotNull(captured);
        Assert.Equal(expectedCommand, captured.Command);
        Assert.Equal("owned", captured.SessionId);
        return JsonDocument.Parse(captured.Args!);
    }
}
