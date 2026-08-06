using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Table;

/// <summary>
/// Bug regression tests for TableCommands.
/// These tests reproduce known bugs and must fail before the fix and pass after.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Tables")]
[Trait("Speed", "Medium")]
public sealed class TableCommandsTests_BugRegression : IClassFixture<TempDirectoryFixture>
{
    private readonly TableCommands _tableCommands;
    private readonly TempDirectoryFixture _fixture;

    /// <summary>
    /// Initializes a new instance of the <see cref="TableCommandsTests_BugRegression"/> class.
    /// </summary>
    public TableCommandsTests_BugRegression(TempDirectoryFixture fixture)
    {
        _tableCommands = new TableCommands();
        _fixture = fixture;
    }

    /// <summary>
    /// Regression test for issue #519:
    /// table append throws COM marshalling exception when row values are JsonElement
    /// (as produced by CLI JSON deserialization of --rows parameter).
    /// Before fix: throws NotSupportedException / InvalidCastException / COMException.
    /// After fix: appends rows successfully.
    /// </summary>
    [Fact]
    public void Append_WithJsonElementValues_DoesNotThrow()
    {
        // Arrange: create a workbook with a table that has string, bool, and number columns
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(TableCommandsTests_BugRegression),
            nameof(Append_WithJsonElementValues_DoesNotThrow),
            _fixture.TempDir,
            ".xlsx");

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create data + table in the same batch (no save needed)
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[1];
            sheet.Name = "Data";
            sheet.Range["A1"].Value2 = "Label";
            sheet.Range["B1"].Value2 = "IsActive";
            sheet.Range["C1"].Value2 = "Amount";
            sheet.Range["A2"].Value2 = "Initial";
            sheet.Range["B2"].Value2 = true;
            sheet.Range["C2"].Value2 = 1.0;
            return 0;
        });
        _tableCommands.Create(batch, "Data", "DataTable", "A1:C2", true, "TableStyleLight1");

        // Act: deserialize rows the same way the CLI does — via JsonSerializer producing JsonElement
        // This is key: the values must be JsonElement (boxed as object?), not raw C# types
        var rowsJson = """[["NewRow", true, 99.5], ["AnotherRow", false, 0.0]]""";
        var deserializedRows = JsonSerializer.Deserialize<List<List<object?>>>(rowsJson)!;

        // Confirm the test is correctly structured: values must be JsonElements, not strings/bools
        Assert.IsType<JsonElement>(deserializedRows[0][0]);
        Assert.IsType<JsonElement>(deserializedRows[0][1]);
        Assert.IsType<JsonElement>(deserializedRows[0][2]);

        // Assert: should not throw — before the fix this throws a COM marshalling exception
        _tableCommands.Append(batch, "DataTable", deserializedRows);

        // Verify rows were appended
        var info = _tableCommands.Read(batch, "DataTable");
        Assert.True(info.Success, $"Read after append failed: {info.ErrorMessage}");
        Assert.Equal(3, info.Table!.RowCount); // 1 original + 2 appended
    }

    /// <summary>
    /// Appending a non-rectangular payload must fail before Excel is mutated.
    /// Silent truncation or partially filled rows makes retries unsafe and hides
    /// caller mistakes, so every row must match the table's column count.
    /// </summary>
    [Fact]
    public void Append_WithMismatchedRowWidths_FailsClosedWithoutChangingTable()
    {
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(TableCommandsTests_BugRegression),
            nameof(Append_WithMismatchedRowWidths_FailsClosedWithoutChangingTable),
            _fixture.TempDir,
            ".xlsx");

        using var batch = ExcelSession.BeginBatch(testFile);
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[1];
            sheet.Name = "Data";
            sheet.Range["A1:C2"].Value2 = new object[,]
            {
                { "Label", "IsActive", "Amount" },
                { "Initial", true, 1.0 },
            };
            return 0;
        });
        _tableCommands.Create(batch, "Data", "DataTable", "A1:C2", true, "TableStyleLight1");

        var malformedRows = new List<List<object?>>
        {
            new() { "TooShort", true },
            new() { "Too", "Many", "Values", 4 },
        };

        var exception = Assert.Throws<ArgumentException>(
            () => _tableCommands.Append(batch, "DataTable", malformedRows));
        Assert.Contains("3", exception.Message, StringComparison.Ordinal);

        var data = _tableCommands.GetData(batch, "DataTable", visibleOnly: false);
        Assert.True(data.Success, $"Read after rejected append failed: {data.ErrorMessage}");
        Assert.Single(data.Data!);
        Assert.Equal("Initial", data.Data![0][0]);
        Assert.Equal(true, data.Data[0][1]);
        Assert.Equal(1.0, Convert.ToDouble(data.Data[0][2], CultureInfo.InvariantCulture));
    }

    /// <summary>
    /// Resizing to append rows must move, rather than overwrite, an enabled totals row.
    /// </summary>
    [Fact]
    public void Append_WithTotalsRow_PreservesTotalFormulaAndAppendedValues()
    {
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(TableCommandsTests_BugRegression),
            nameof(Append_WithTotalsRow_PreservesTotalFormulaAndAppendedValues),
            _fixture.TempDir,
            ".xlsx");

        using var batch = ExcelSession.BeginBatch(testFile);
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[1];
            sheet.Name = "Data";
            sheet.Range["A1:C3"].Value2 = new object[,]
            {
                { "Label", "IsActive", "Amount" },
                { "First", true, 1.0 },
                { "Second", false, 2.0 },
            };
            return 0;
        });
        _tableCommands.Create(batch, "Data", "DataTable", "A1:C3", true, "TableStyleLight1");
        _tableCommands.ToggleTotals(batch, "DataTable", true);
        _tableCommands.SetColumnTotal(batch, "DataTable", "Amount", "Sum");

        _tableCommands.Append(
            batch,
            "DataTable",
            [new List<object?> { "Third", true, 3.0 }]);

        var info = _tableCommands.Read(batch, "DataTable");
        Assert.True(info.Success, info.ErrorMessage);
        Assert.True(info.Table!.ShowTotals);
        Assert.Equal(3, info.Table.RowCount);

        var data = _tableCommands.GetData(batch, "DataTable", visibleOnly: false);
        Assert.Equal(3, data.RowCount);
        Assert.Equal("Third", data.Data[2][0]);
        Assert.Equal(3.0, Convert.ToDouble(data.Data[2][2], CultureInfo.InvariantCulture));

        var rangeCommands = new RangeCommands();
        var total = rangeCommands.GetFormulas(batch, "Data", "C5");
        Assert.True(total.Success, total.ErrorMessage);
        Assert.Contains("SUBTOTAL", total.Formulas[0][0], StringComparison.OrdinalIgnoreCase);
        Assert.Equal(6.0, Convert.ToDouble(total.Values[0][0], CultureInfo.InvariantCulture));
    }
}
