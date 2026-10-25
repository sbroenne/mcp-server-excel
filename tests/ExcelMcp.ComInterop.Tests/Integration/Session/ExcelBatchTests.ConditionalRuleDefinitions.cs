using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

public partial class ExcelBatchTests
{
    [Fact]
    public void Execute_TypedConditionalRuleDefinitions_ReadAllSevenNativeInterfaces()
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);
        batch.Execute((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.FormatConditions? conditions = null;
            Excel.FormatCondition? basic = null;
            Excel.ColorScale? scale = null;
            Excel.Databar? bar = null;
            Excel.IconSetCondition? icons = null;
            Excel.Top10? top = null;
            Excel.AboveAverage? average = null;
            Excel.UniqueValues? unique = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[1];
                range = sheet.Range["A1:A3"];
                range.Value2 = new object[,] { { -5 }, { 17 }, { 29 } };
                conditions = range.FormatConditions;
                conditions.Delete();
                basic = conditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlGreater, 10);
                scale = conditions.AddColorScale(3);
                bar = conditions.AddDatabar();
                icons = conditions.AddIconSetCondition();
                top = conditions.AddTop10();
                average = conditions.AddAboveAverage();
                unique = conditions.AddUniqueValues();
                Assert.Equal(7, conditions.Count);
                Assert.Equal(1, basic.Type);
                Assert.Equal(3, scale.Type);
                Assert.Equal(4, bar.Type);
                Assert.Equal(6, icons.Type);
                Assert.Equal(5, top.Type);
                Assert.Equal(12, average.Type);
                Assert.Equal(8, unique.Type);
                var types = new List<int>();
                var priorities = new List<int>();
                for (int index = 1; index <= conditions.Count; index++)
                {
                    object? rule = null;
                    Excel.Range? appliesTo = null;
                    try
                    {
                        rule = conditions[index];
                        (int type, int priority, Excel.Range nativeRange) = rule switch
                        {
                            Excel.FormatCondition item => (item.Type, item.Priority, item.AppliesTo),
                            Excel.ColorScale item => (item.Type, item.Priority, item.AppliesTo),
                            Excel.Databar item => (item.Type, item.Priority, item.AppliesTo),
                            Excel.IconSetCondition item => (item.Type, item.Priority, item.AppliesTo),
                            Excel.Top10 item => (item.Type, item.Priority, item.AppliesTo),
                            Excel.AboveAverage item => (item.Type, item.Priority, item.AppliesTo),
                            Excel.UniqueValues item => (item.Type, item.Priority, item.AppliesTo),
                            _ => throw new InvalidOperationException("Excel returned an unexpected conditional-rule interface.")
                        };
                        appliesTo = nativeRange;
                        types.Add(type);
                        priorities.Add(priority);
                        Assert.Equal("$A$1:$A$3", appliesTo.Address);
                    }
                    finally
                    {
                        ComUtilities.Release(ref appliesTo);
                        ComUtilities.Release(ref rule);
                    }
                }
                Assert.Equal([1, 3, 4, 5, 6, 8, 12], types.Order());
                Assert.Equal([1, 2, 3, 4, 5, 6, 7], priorities.Order());
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref unique);
                ComUtilities.Release(ref average);
                ComUtilities.Release(ref top);
                ComUtilities.Release(ref icons);
                ComUtilities.Release(ref bar);
                ComUtilities.Release(ref scale);
                ComUtilities.Release(ref basic);
                ComUtilities.Release(ref conditions);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        AssertNoExternalInteropAssemblies();
    }

    [Theory]
    [InlineData(null)]
    [InlineData(false)]
    [InlineData(true)]
    public void Execute_TypedConditionalRuleDefinitions_PreserveUnsetFontValues(bool? bold)
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);
        batch.Execute((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.FormatConditions? conditions = null;
            Excel.FormatCondition? condition = null;
            Excel.Font? font = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[1];
                range = sheet.Range["A1:A3"];
                conditions = range.FormatConditions;
                conditions.Delete();
                condition = conditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlGreater, 5);
                font = condition.Font;
                if (bold.HasValue)
                {
                    font.Bold = bold.Value;
                    Assert.Equal(bold.Value, (object)font.Bold);
                    Assert.Equal(false, (object)font.Italic);
                }
                else
                {
                    Assert.IsType<DBNull>((object)font.Bold);
                    Assert.IsType<DBNull>((object)font.Italic);
                }
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref font);
                ComUtilities.Release(ref condition);
                ComUtilities.Release(ref conditions);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        AssertNoExternalInteropAssemblies();
    }
}
