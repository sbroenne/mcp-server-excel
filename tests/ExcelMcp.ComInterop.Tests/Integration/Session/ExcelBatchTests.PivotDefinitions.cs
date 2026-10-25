using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

public partial class ExcelBatchTests
{
    [Theory]
    [InlineData(true, 0)]
    [InlineData(true, 1)]
    [InlineData(true, 2)]
    [InlineData(false, 2)]
    public void Execute_TypedPivotDefinitions_PreserveValuesPlaceholderAndVisibleMembers(
        bool readPlaceholder, int valueFieldCount)
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);
        batch.Execute((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? source = null;
            Excel.Range? literalHeader = null;
            Excel.Range? destination = null;
            Excel.PivotCaches? caches = null;
            Excel.PivotCache? cache = null;
            Excel.PivotTable? pivot = null;
            Excel.PivotFields? fields = null;
            Excel.PivotField? region = null;
            Excel.PivotField? sales = null;
            Excel.PivotField? quantity = null;
            Excel.PivotField? salesTotal = null;
            Excel.PivotField? quantityTotal = null;
            Excel.PivotField? valuesField = null;
            Excel.PivotItems? items = null;
            Excel.PivotItem? south = null;
            Excel.PivotItems? visible = null;
            Excel.PivotItem? remaining = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[1];
                literalHeader = sheet.Range["D1"];
                literalHeader.NumberFormat = "@";
                source = sheet.Range["A1:E5"];
                source.Value2 = new object[,]
                {
                    { "Region", "Sales", "Quantity", "-2146826246", "Values" },
                    { "North", 17, 1, "One", "First" },
                    { "South", 29, 2, "Two", "Second" },
                    { "North", 31, 3, "Three", "Third" },
                    { "South", 43, 4, "Four", "Fourth" }
                };
                destination = sheet.Range["F1"];
                caches = ctx.Book.PivotCaches();
                cache = caches.Create(Excel.XlPivotTableSourceType.xlDatabase,
                    $"'{sheet.Name.Replace("'", "''", StringComparison.Ordinal)}'!A1:E5");
                pivot = cache.CreatePivotTable(destination, "TypedDefinitionPivot");
                fields = (Excel.PivotFields)pivot.PivotFields();
                region = fields.Item("Region");
                region.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                sales = fields.Item("Sales");
                quantity = fields.Item("Quantity");
                if (valueFieldCount >= 1)
                    salesTotal = pivot.AddDataField(sales, "Sales total", Excel.XlConsolidationFunction.xlSum);
                if (valueFieldCount >= 2)
                    quantityTotal = pivot.AddDataField(quantity, "Quantity total", Excel.XlConsolidationFunction.xlSum);
                Assert.True(pivot.RefreshTable());

                if (readPlaceholder)
                {
                    if (valueFieldCount >= 2)
                        valuesField = pivot.DataPivotField;
                    ComUtilities.Release(ref fields);
                    fields = (Excel.PivotFields)pivot.PivotFields();
                    int placeholderCount = 0;
                    int ordinarySentinelCount = 0;
                    for (int index = 1; index <= fields.Count; index++)
                    {
                        Excel.PivotField? candidate = null;
                        Excel.PivotTable? owner = null;
                        Excel.PivotFields? ownerValues = null;
                        Excel.PivotField? ownerPlaceholder = null;
                        try
                        {
                            candidate = fields.Item(index);
                            object sourceName = candidate.SourceName;
                            bool placeholder = false;
                            if (sourceName is string marker && marker == "-2146826246")
                            {
                                owner = (Excel.PivotTable)candidate.Parent;
                                Assert.Equal(pivot.Name, owner.Name);
                                ownerValues = (Excel.PivotFields)owner.DataFields;
                                Assert.Equal(valueFieldCount, ownerValues.Count);
                                if (ownerValues.Count >= 2)
                                {
                                    ownerPlaceholder = owner.DataPivotField;
                                    placeholder = candidate.Name == ownerPlaceholder.Name;
                                }
                            }
                            if (placeholder)
                            {
                                Assert.Equal("-2146826246", Assert.IsType<string>(sourceName));
                                placeholderCount++;
                            }
                            else if (sourceName is string text && text == "-2146826246")
                            {
                                if (valuesField is not null)
                                    Assert.NotEqual(valuesField.Name, candidate.Name);
                                ordinarySentinelCount++;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref ownerPlaceholder);
                            ComUtilities.Release(ref ownerValues);
                            ComUtilities.Release(ref owner);
                            ComUtilities.Release(ref candidate);
                        }
                    }
                    Assert.Equal(valueFieldCount >= 2 ? 1 : 0, placeholderCount);
                    Assert.Equal(1, ordinarySentinelCount);
                }
                else
                {
                    visible = (Excel.PivotItems)region.VisibleItems[Type.Missing];
                    Assert.Equal(2, visible.Count);
                    ComUtilities.Release(ref visible);
                    items = (Excel.PivotItems)region.PivotItems();
                    south = items.Item("South");
                    south.Visible = false;
                    visible = (Excel.PivotItems)region.VisibleItems[Type.Missing];
                    Assert.Equal(1, visible.Count);
                    remaining = (Excel.PivotItem)region.VisibleItems[1];
                    Assert.Equal("North", remaining.Name);
                    Assert.True(remaining.Visible);
                    Assert.False(south.Visible);
                }

                object[,] unchanged = Assert.IsType<object[,]>((object)source.Value2);
                Assert.Equal("Region", unchanged[1, 1]);
                Assert.Equal(43d, Convert.ToDouble(unchanged[5, 2], CultureInfo.InvariantCulture));
                Assert.Equal(4d, Convert.ToDouble(unchanged[5, 3], CultureInfo.InvariantCulture));
                Assert.Equal("-2146826246", unchanged[1, 4]);
                Assert.Equal("Four", unchanged[5, 4]);
                Assert.Equal("Values", unchanged[1, 5]);
                Assert.Equal("Fourth", unchanged[5, 5]);
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref remaining);
                ComUtilities.Release(ref visible);
                ComUtilities.Release(ref south);
                ComUtilities.Release(ref items);
                ComUtilities.Release(ref valuesField);
                ComUtilities.Release(ref quantityTotal);
                ComUtilities.Release(ref salesTotal);
                ComUtilities.Release(ref quantity);
                ComUtilities.Release(ref sales);
                ComUtilities.Release(ref region);
                ComUtilities.Release(ref fields);
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref caches);
                ComUtilities.Release(ref destination);
                ComUtilities.Release(ref source);
                ComUtilities.Release(ref literalHeader);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        AssertNoExternalInteropAssemblies();
    }
}
