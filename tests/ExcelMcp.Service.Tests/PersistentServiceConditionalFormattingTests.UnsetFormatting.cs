using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceConditionalFormattingTests
{
    [Theory]
    [InlineData(null)]
    [InlineData(false)]
    [InlineData(true)]
    public void ListRules_UnsetNativeFontVariants_AreNullNotFalse(bool? bold)
    {
        RequireSuccess(_conditionalFormattingCommands.AddRule(
            _fixture.BatchToken, _sheetName, "A1:A5", "cellValue", "greater", "5", null,
            interiorColor: "#FFFF00", fontBold: bold));
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.FormatConditions? conditions = null;
            Excel.FormatCondition? condition = null;
            Excel.Font? font = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, _sheetName);
                Assert.NotNull(sheet);
                range = sheet.Range["A1:A5"];
                conditions = range.FormatConditions;
                condition = (Excel.FormatCondition)conditions.Item(1);
                font = condition.Font;
                object? nativeBold = font.Bold;
                object? nativeItalic = font.Italic;
                if (bold.HasValue)
                {
                    Assert.Equal(bold.Value, Assert.IsType<bool>(nativeBold));
                    Assert.False(Assert.IsType<bool>(nativeItalic));
                }
                else
                {
                    Assert.IsType<DBNull>(nativeBold);
                    Assert.IsType<DBNull>(nativeItalic);
                }
            }
            finally
            {
                ComUtilities.Release(ref font);
                ComUtilities.Release(ref condition);
                ComUtilities.Release(ref conditions);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
        var rule = Assert.Single(RequireSuccess(
            _conditionalFormattingCommands.ListRules(_fixture.BatchToken, _sheetName, "A1:A5")).Rules);
        Assert.Equal(bold, rule.FontBold);
        Assert.Equal(bold.HasValue ? false : (bool?)null, rule.FontItalic);
        Assert.Null(rule.FontColor);
        Assert.Null(rule.BorderStyle);
        Assert.Equal("#FFFF00", rule.InteriorColor);
        var worksheetRule = Assert.Single(RequireSuccess(
            _conditionalFormattingCommands.ListWorksheetRules(_fixture.BatchToken, _sheetName)).Rules);
        Assert.Equal(rule.Fingerprint, worksheetRule.Fingerprint);
    }
}
