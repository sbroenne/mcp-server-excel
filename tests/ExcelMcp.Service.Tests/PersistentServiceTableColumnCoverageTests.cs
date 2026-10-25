using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Tables")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceTableColumnCoverageTests(PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture), IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly IPersistentTableCommands _tables = fixture.CreateCommands<IPersistentTableCommands>();

    [Fact]
    public void RemoveColumn_DeletesMiddleColumnAndKeepsRemainingRecordsAligned()
    {
        var (sheet, table) = SeedTable();
        var neighborBefore = ReadNativeTable(sheet, table + "_Neighbor");
        Assert.Equal(table + "_Neighbor", neighborBefore.Name);
        Assert.Equal(["Other", "Value"], neighborBefore.Headers);
        Assert.Equal("TableStyleLight1", neighborBefore.Style);
        Assert.Equal(1, neighborBefore.Rows);
        AssertGrid(sheet, neighborBefore.Address, [["Other", "Value"], ["untouched", 77]]);
        RequireSuccess(_tables.RemoveColumn(_fixture.BatchToken, table, "Priority"));
        AssertGrid(sheet, "A1:B5",
            [["Group", "Record"], ["B", "b-low"], ["A", "a-low"], ["B", "b-high"], ["A", "a-high"]]);
        var native = ReadNativeTable(sheet, table);
        Assert.Equal(["Group", "Record"], native.Headers);
        Assert.Equal(4, native.Rows);
        var neighborAfter = ReadNativeTable(sheet, neighborBefore.Name);
        Assert.Equal(neighborBefore.Name, neighborAfter.Name);
        Assert.Equal(neighborBefore.Headers, neighborAfter.Headers);
        Assert.Equal(neighborBefore.Style, neighborAfter.Style);
        Assert.Equal(neighborBefore.Rows, neighborAfter.Rows);
        AssertGrid(sheet, neighborAfter.Address, [["Other", "Value"], ["untouched", 77]]);
    }

    [Fact]
    public void RemoveColumn_DeletesMiddleColumnAndPreservesTableInUnrelatedRows()
    {
        var (sheet, table) = SeedTable("E8:F9");
        RequireSuccess(_tables.RemoveColumn(_fixture.BatchToken, table, "Priority"));
        AssertGrid(sheet, "A1:B5",
            [["Group", "Record"], ["B", "b-low"], ["A", "a-low"], ["B", "b-high"], ["A", "a-high"]]);
        var native = ReadNativeTable(sheet, table);
        Assert.Equal(["Group", "Record"], native.Headers);
        Assert.Equal(4, native.Rows);
        AssertGrid(sheet, "E8:F9", [["Other", "Value"], ["untouched", 77]]);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Sort_OrdersWholeRecordsWithoutChangingNeighbor(bool ascending)
    {
        var (sheet, table) = SeedTable();
        RequireSuccess(_tables.Sort(_fixture.BatchToken, table, "Record", ascending));
        List<List<object?>> expected = ascending
            ? [["A", 9, "a-high"], ["A", 2, "a-low"], ["B", 8, "b-high"], ["B", 1, "b-low"]]
            : [["B", 1, "b-low"], ["B", 8, "b-high"], ["A", 2, "a-low"], ["A", 9, "a-high"]];
        AssertGrid(sheet, "A2:C5", expected);
        AssertGrid(sheet, "A1:C1", [["Group", "Priority", "Record"]]);
        AssertGrid(sheet, "E1:F2", [["Other", "Value"], ["untouched", 77]]);
    }

    [Fact]
    public void SortMulti_OneKeyOrdersWholeRecordsWithoutChangingHeadersOrNeighbor()
    {
        var (sheet, table) = SeedTable();
        RequireSuccess(_tables.SortMulti(_fixture.BatchToken, table,
            [new() { ColumnName = "Record", Ascending = false }]));
        AssertGrid(sheet, "A1:C5",
            [["Group", "Priority", "Record"], ["B", 1, "b-low"], ["B", 8, "b-high"], ["A", 2, "a-low"], ["A", 9, "a-high"]]);
        var native = ReadNativeTable(sheet, table);
        Assert.Equal(["Group", "Priority", "Record"], native.Headers);
        Assert.Equal(4, native.Rows);
        var neighbor = ReadNativeTable(sheet, table + "_Neighbor");
        Assert.Equal(["Other", "Value"], neighbor.Headers);
        Assert.Equal(1, neighbor.Rows);
        Assert.Equal("TableStyleLight1", neighbor.Style);
        AssertGrid(sheet, "E1:F2", [["Other", "Value"], ["untouched", 77]]);
    }

    [Fact]
    public void SortMulti_UsesPrimaryAndDescendingTieBreakerAndKeepsRecordsAligned()
    {
        var (sheet, table) = SeedTable();
        RequireSuccess(_tables.SortMulti(_fixture.BatchToken, table,
            [new() { ColumnName = "Group", Ascending = true }, new() { ColumnName = "Priority", Ascending = false }]));
        AssertGrid(sheet, "A1:C5",
            [["Group", "Priority", "Record"], ["A", 9, "a-high"], ["A", 2, "a-low"], ["B", 8, "b-high"], ["B", 1, "b-low"]]);
        AssertGrid(sheet, "E1:F2", [["Other", "Value"], ["untouched", 77]]);
    }

    [Fact]
    public void SortMulti_ThreeKeysBreakPrimaryAndSecondaryTiesWithoutChangingHeadersOrNeighbor()
    {
        var (sheet, table) = SeedTable(records:
            [["Group", "Priority", "Record"],
             ["B", 1, "b-low-z"], ["A", 2, "a-high-z"], ["B", 2, "b-high-z"], ["A", 1, "a-low-z"],
             ["A", 2, "a-high-a"], ["B", 1, "b-low-a"], ["A", 1, "a-low-a"], ["B", 2, "b-high-a"]]);
        RequireSuccess(_tables.SortMulti(_fixture.BatchToken, table,
            [new() { ColumnName = "Group", Ascending = true },
             new() { ColumnName = "Priority", Ascending = false },
             new() { ColumnName = "Record", Ascending = true }]));
        AssertGrid(sheet, "A1:C9",
            [["Group", "Priority", "Record"],
             ["A", 2, "a-high-a"], ["A", 2, "a-high-z"], ["A", 1, "a-low-a"], ["A", 1, "a-low-z"],
             ["B", 2, "b-high-a"], ["B", 2, "b-high-z"], ["B", 1, "b-low-a"], ["B", 1, "b-low-z"]]);
        var native = ReadNativeTable(sheet, table);
        Assert.Equal(["Group", "Priority", "Record"], native.Headers);
        Assert.Equal(8, native.Rows);
        var neighbor = ReadNativeTable(sheet, table + "_Neighbor");
        Assert.Equal(["Other", "Value"], neighbor.Headers);
        Assert.Equal(1, neighbor.Rows);
        Assert.Equal("TableStyleLight1", neighbor.Style);
        AssertGrid(sheet, "E1:F2", [["Other", "Value"], ["untouched", 77]]);
    }

    [Fact]
    public void GetColumnNumberFormat_ReadsHeaderAndEveryMixedDataCell()
    {
        var (sheet, table) = SeedTable();
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheet, "B2:B5", [[1.25], [0.375], [42.5], [-7.75]],
            overwritePolicy: OverwritePolicy.Allow));
        string[] formats = ["0.00", "0%", "0.000", "0.0"];
        RequireSuccess(_commands.SetNumberFormat(_fixture.BatchToken, sheet, "B1", "@"));
        for (int i = 0; i < formats.Length; i++)
            RequireSuccess(_commands.SetNumberFormat(_fixture.BatchToken, sheet, $"B{i + 2}", formats[i]));
        var result = RequireSuccess(_tables.GetColumnNumberFormat(_fixture.BatchToken, table, "Priority"));
        Assert.Equal(sheet, result.SheetName);
        Assert.Equal("$B$1:$B$5", result.RangeAddress);
        Assert.Equal(5, result.RowCount);
        Assert.Equal(1, result.ColumnCount);
        Assert.Equal(formats.Prepend("@"), result.Formats.Select(row => Assert.Single(row)));
        Assert.Equal(formats, ReadNativeTable(sheet, table).Formats);
        AssertGrid(sheet, "A1:C5",
            [["Group", "Priority", "Record"], ["B", 1.25, "b-low"], ["A", 0.375, "a-low"], ["B", 42.5, "b-high"], ["A", -7.75, "a-high"]]);
        AssertGrid(sheet, "E1:F2", [["Other", "Value"], ["untouched", 77]]);
    }

    [Fact]
    public void SetColumnNumberFormat_ReplacesMixedFormatsWithoutChangingHeaderOtherColumnsOrValues()
    {
        var (sheet, table) = SeedTable();
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheet, "B2:B5", [[1.25], [0.375], [42.5], [-7.75]],
            overwritePolicy: OverwritePolicy.Allow));
        RequireSuccess(_commands.SetNumberFormat(_fixture.BatchToken, sheet, "B2:B3", "0%"));
        RequireSuccess(_commands.SetNumberFormat(_fixture.BatchToken, sheet, "B4:B5", "0.000"));
        RequireSuccess(_commands.SetNumberFormat(_fixture.BatchToken, sheet, "B1", "@"));
        RequireSuccess(_commands.SetNumberFormat(_fixture.BatchToken, sheet, "C2:C5", "@"));
        RequireSuccess(_tables.SetColumnNumberFormat(_fixture.BatchToken, table, "Priority", "0.00"));
        Assert.Equal(["0.00", "0.00", "0.00", "0.00"], ReadNativeTable(sheet, table).Formats);
        Assert.Equal("@", PersistentServiceRangeVerification.ReadFormat(_fixture, sheet, "B1").NumberFormat);
        for (int row = 2; row <= 5; row++)
            Assert.Equal("@", PersistentServiceRangeVerification.ReadFormat(_fixture, sheet, $"C{row}").NumberFormat);
        AssertGrid(sheet, "A1:C5",
            [["Group", "Priority", "Record"], ["B", 1.25, "b-low"], ["A", 0.375, "a-low"], ["B", 42.5, "b-high"], ["A", -7.75, "a-high"]]);
        AssertGrid(sheet, "E1:F2", [["Other", "Value"], ["untouched", 77]]);
    }

    [Fact]
    public void SetStyle_ChangesOnlyTargetTableAndPreservesAllData()
    {
        var (sheet, table) = SeedTable();
        var before = ReadNativeTable(sheet, table);
        Assert.Equal("TableStyleMedium2", before.Style);
        RequireSuccess(_tables.SetStyle(_fixture.BatchToken, table, "TableStyleDark9"));
        Assert.Equal("TableStyleDark9", ReadNativeTable(sheet, table).Style);
        Assert.Equal("TableStyleLight1", ReadNativeTable(sheet, table + "_Neighbor").Style);
        AssertGrid(sheet, "A1:C5", OriginalRecords());
        AssertGrid(sheet, "E1:F2", [["Other", "Value"], ["untouched", 77]]);
    }

    private (string Sheet, string Table) SeedTable(string neighborAddress = "E1:F2", List<List<object?>>? records = null)
    {
        string sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        string table = $"Coverage_{Guid.NewGuid():N}";
        records ??= OriginalRecords();
        string address = $"A1:C{records.Count}";
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheet, address, records));
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheet, neighborAddress, [["Other", "Value"], ["untouched", 77]]));
        RequireSuccess(_tables.Create(_fixture.BatchToken, sheet, table, address, tableStyle: "TableStyleMedium2"));
        _fixture.RegisterTableForCleanup(table);
        RequireSuccess(_tables.Create(_fixture.BatchToken, sheet, table + "_Neighbor", neighborAddress, tableStyle: "TableStyleLight1"));
        _fixture.RegisterTableForCleanup(table + "_Neighbor");
        return (sheet, table);
    }

    private static List<List<object?>> OriginalRecords() =>
        [["Group", "Priority", "Record"], ["B", 1, "b-low"], ["A", 2, "a-low"], ["B", 8, "b-high"], ["A", 9, "a-high"]];

    private void AssertGrid(string sheet, string address, List<List<object?>> expected)
    {
        var actual = RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheet, address));
        Assert.Equal(expected.Count, actual.RowCount);
        Assert.Equal(expected[0].Count, actual.ColumnCount);
        Assert.Equal(expected.SelectMany(row => row).Select(value => Convert.ToString(value, CultureInfo.InvariantCulture)),
            actual.Values.SelectMany(row => row).Select(value => Convert.ToString(value, CultureInfo.InvariantCulture)));
    }

    private (string[] Headers, string[] Formats, string Style, int Rows, string Name, string Address) ReadNativeTable(string sheetName, string tableName) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.ListObjects? tables = null;
            Excel.ListObject? table = null;
            Excel.ListRows? rows = null;
            Excel.Range? headers = null;
            Excel.ListColumns? columns = null;
            Excel.ListColumn? column = null;
            Excel.Range? data = null;
            Excel.Range? cells = null;
            Excel.Range? cell = null;
            Excel.Range? range = null;
            Excel.TableStyle? style = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName) ??
                    throw new InvalidOperationException($"Excel did not return worksheet '{sheetName}'.");
                tables = sheet.ListObjects;
                table = tables[tableName];
                rows = table.ListRows;
                headers = table.HeaderRowRange;
                var values = (object[,])(object)(headers ??
                    throw new InvalidOperationException("Excel did not return the table header range.")).Value2;
                string[] names = Enumerable.Range(1, values.GetLength(1))
                    .Select(index => Convert.ToString(values[1, index], CultureInfo.InvariantCulture)!).ToArray();
                columns = table.ListColumns;
                string[] formats = [];
                if (names.Contains("Priority"))
                {
                    column = columns["Priority"];
                    data = column.DataBodyRange;
                    cells = data.Cells;
                    formats = new string[rows.Count];
                    for (int index = 1; index <= formats.Length; index++)
                    {
                        try
                        {
                            cell = (Excel.Range)cells[index, 1];
                            formats[index - 1] = Convert.ToString((object?)cell.NumberFormat, CultureInfo.InvariantCulture)!;
                        }
                        finally { ComUtilities.Release(ref cell); }
                    }
                }
                style = (Excel.TableStyle)(object)table.TableStyle;
                range = table.Range;
                return (names, formats, style.Name, rows.Count, table.Name, range.Address);
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref style);
                ComUtilities.Release(ref cells);
                ComUtilities.Release(ref data);
                ComUtilities.Release(ref column);
                ComUtilities.Release(ref columns);
                ComUtilities.Release(ref headers);
                ComUtilities.Release(ref rows);
                ComUtilities.Release(ref table);
                ComUtilities.Release(ref tables);
                ComUtilities.Release(ref sheet);
            }
        });
}
