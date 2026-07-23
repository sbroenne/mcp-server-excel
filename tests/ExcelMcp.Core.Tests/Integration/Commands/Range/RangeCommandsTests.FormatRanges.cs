using System.Reflection;
using System.Runtime.ExceptionServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

public partial class RangeCommandsTests
{
    private const int YellowFillColor = 65535;
    private const int CenterAlignment = -4108;

    [Fact]
    public void FormatRanges_AppliesSharedFormattingToEachTargetRange_AndLeavesUntargetedCellsUnchanged()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        var method = GetFormatRangesMethod();

        var untouchedBefore = new[]
        {
            ReadCellFormattingState(batch, sheetName, "B1"),
            ReadCellFormattingState(batch, sheetName, "B2")
        };

        var result = InvokeFormatRanges(
            method,
            batch,
            sheetName,
            ["A1:A2", "C1:C2"],
            bold: true,
            fillColor: "#FFFF00",
            horizontalAlignment: "center");

        Assert.True(result.Success, $"FormatRanges failed: {result.ErrorMessage}");

        Assert.Equal(new CellFormattingState(true, YellowFillColor, CenterAlignment), ReadCellFormattingState(batch, sheetName, "A1"));
        Assert.Equal(new CellFormattingState(true, YellowFillColor, CenterAlignment), ReadCellFormattingState(batch, sheetName, "A2"));
        Assert.Equal(new CellFormattingState(true, YellowFillColor, CenterAlignment), ReadCellFormattingState(batch, sheetName, "C1"));
        Assert.Equal(new CellFormattingState(true, YellowFillColor, CenterAlignment), ReadCellFormattingState(batch, sheetName, "C2"));

        Assert.Equal(untouchedBefore[0], ReadCellFormattingState(batch, sheetName, "B1"));
        Assert.Equal(untouchedBefore[1], ReadCellFormattingState(batch, sheetName, "B2"));
    }

    [Fact]
    public void FormatRanges_WithNumberFormat_AppliesFormatCodeToAllTargetRanges()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        var method = GetFormatRangesMethod();

        var result = InvokeFormatRanges(
            method,
            batch,
            sheetName,
            ["A1:A2", "C1:C2"],
            numberFormat: "0.00%");

        Assert.True(result.Success, $"FormatRanges failed: {result.ErrorMessage}");

        Assert.Equal("0.00%", ReadCellNumberFormat(batch, sheetName, "A1"));
        Assert.Equal("0.00%", ReadCellNumberFormat(batch, sheetName, "A2"));
        Assert.Equal("0.00%", ReadCellNumberFormat(batch, sheetName, "C1"));
        Assert.Equal("0.00%", ReadCellNumberFormat(batch, sheetName, "C2"));
        // Untouched cells must keep their original format
        Assert.NotEqual("0.00%", ReadCellNumberFormat(batch, sheetName, "B1"));
    }

    [Fact]
    public void FormatRanges_InvalidTargetAddress_ErrorMessageIncludesIndex()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        var method = GetFormatRangesMethod();

        // "NotARange" is at index 1 (the second item)
        var ex = Assert.Throws<ArgumentException>(() => InvokeFormatRanges(
            method,
            batch,
            sheetName,
            ["A1:A2", "NotARange"],
            bold: true));

        // Error message must include the array index so the caller can identify which entry is bad
        Assert.Contains("1", ex.Message);
    }

    [Fact]
    public void FormatRanges_InvalidTargetAddress_FailsFast_AndDoesNotPartiallyApplyEarlierRanges()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        var method = GetFormatRangesMethod();

        var a1Before = ReadCellFormattingState(batch, sheetName, "A1");
        var a2Before = ReadCellFormattingState(batch, sheetName, "A2");

        var exception = Assert.Throws<ArgumentException>(() => InvokeFormatRanges(
            method,
            batch,
            sheetName,
            ["A1:A2", "NotARange"],
            bold: true,
            fillColor: "#FFFF00",
            horizontalAlignment: "center"));

        Assert.Contains("range", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(a1Before, ReadCellFormattingState(batch, sheetName, "A1"));
        Assert.Equal(a2Before, ReadCellFormattingState(batch, sheetName, "A2"));
    }

    private static MethodInfo GetFormatRangesMethod()
    {
        var parameterTypes = new[]
        {
            typeof(IExcelBatch),
            typeof(string),
            typeof(string[]),
            typeof(string),
            typeof(double?),
            typeof(bool?),
            typeof(bool?),
            typeof(bool?),
            typeof(string),
            typeof(string),
            typeof(string),
            typeof(string),
            typeof(string),
            typeof(string),
            typeof(string),
            typeof(bool?),
            typeof(int?),
            typeof(string)  // numberFormat
        };

        var expectedParameters = new (string Name, Type Type)[]
        {
            ("batch", typeof(IExcelBatch)),
            ("sheetName", typeof(string)),
            ("rangeAddresses", typeof(string[])),
            ("fontName", typeof(string)),
            ("fontSize", typeof(double?)),
            ("bold", typeof(bool?)),
            ("italic", typeof(bool?)),
            ("underline", typeof(bool?)),
            ("fontColor", typeof(string)),
            ("fillColor", typeof(string)),
            ("borderStyle", typeof(string)),
            ("borderColor", typeof(string)),
            ("borderWeight", typeof(string)),
            ("horizontalAlignment", typeof(string)),
            ("verticalAlignment", typeof(string)),
            ("wrapText", typeof(bool?)),
            ("orientation", typeof(int?)),
            ("numberFormat", typeof(string))  // Bug 3 closure: apply number format atomically
        };

        var interfaceMethod = typeof(IRangeFormatCommands).GetMethod("FormatRanges", parameterTypes);
        Assert.True(interfaceMethod is not null,
            "Expected IRangeFormatCommands.FormatRanges(batch, sheetName, rangeAddresses, ...format props...) to exist for Bug 4 v1.");

        Assert.Equal(expectedParameters.Length, interfaceMethod!.GetParameters().Length);
        for (var index = 0; index < expectedParameters.Length; index++)
        {
            Assert.Equal(expectedParameters[index].Name, interfaceMethod.GetParameters()[index].Name);
            Assert.Equal(expectedParameters[index].Type, interfaceMethod.GetParameters()[index].ParameterType);
        }

        var implementationMethod = typeof(RangeCommands).GetMethod("FormatRanges", parameterTypes);
        Assert.True(implementationMethod is not null,
            "Expected RangeCommands.FormatRanges(batch, sheetName, rangeAddresses, ...format props...) implementation to exist for Bug 4 v1.");
        Assert.Equal(typeof(OperationResult), implementationMethod!.ReturnType);

        return implementationMethod;
    }

    private OperationResult InvokeFormatRanges(
        MethodInfo method,
        IExcelBatch batch,
        string sheetName,
        string[] rangeAddresses,
        string? fontName = null,
        double? fontSize = null,
        bool? bold = null,
        bool? italic = null,
        bool? underline = null,
        string? fontColor = null,
        string? fillColor = null,
        string? borderStyle = null,
        string? borderColor = null,
        string? borderWeight = null,
        string? horizontalAlignment = null,
        string? verticalAlignment = null,
        bool? wrapText = null,
        int? orientation = null,
        string? numberFormat = null)
    {
        try
        {
            return Assert.IsType<OperationResult>(method.Invoke(_commands, [
                batch,
                sheetName,
                rangeAddresses,
                fontName,
                fontSize,
                bold,
                italic,
                underline,
                fontColor,
                fillColor,
                borderStyle,
                borderColor,
                borderWeight,
                horizontalAlignment,
                verticalAlignment,
                wrapText,
                orientation,
                numberFormat
            ]));
        }
        catch (TargetInvocationException ex) when (ex.InnerException is not null)
        {
            ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
            throw;
        }
    }

    private static CellFormattingState ReadCellFormattingState(IExcelBatch batch, string sheetName, string cellAddress)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? font = null;
            dynamic? interior = null;

            try
            {
                sheet = ctx.Book.Worksheets[sheetName];
                range = sheet.Range[cellAddress];
                font = range.Font;
                interior = range.Interior;

                return new CellFormattingState(
                    Bold: Convert.ToBoolean(font.Bold),
                    FillColor: Convert.ToInt32(interior.Color),
                    HorizontalAlignment: Convert.ToInt32(range.HorizontalAlignment));
            }
            finally
            {
                ComUtilities.Release(ref interior!);
                ComUtilities.Release(ref font!);
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    private static string ReadCellNumberFormat(IExcelBatch batch, string sheetName, string cellAddress)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;

            try
            {
                sheet = ctx.Book.Worksheets[sheetName];
                range = sheet.Range[cellAddress];
                return (string)(range.NumberFormat ?? "General");
            }
            finally
            {
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    private readonly record struct CellFormattingState(bool Bold, int FillColor, int HorizontalAlignment);
}