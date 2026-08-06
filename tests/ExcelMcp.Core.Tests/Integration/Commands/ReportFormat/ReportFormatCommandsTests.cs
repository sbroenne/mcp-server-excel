using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.ReportFormat;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands.ReportFormat;

[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "ReportFormat")]
public sealed class ReportFormatCommandsTests : IClassFixture<WindowTestsFixture>
{
    private readonly ReportFormatCommands _commands = new();
    private readonly WindowTestsFixture _fixture;

    public ReportFormatCommandsTests(WindowTestsFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void Apply_ReadbackAndReapply_AreDeterministicAndPreserveWorkbookContent()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        SeedReport(batch);

        var applied = _commands.Apply(
            batch,
            "Sheet1",
            "A1:D1",
            "A2:D2",
            "A3:D5",
            "A6:D6",
            ReportFormatPreset.Professional,
            "#1F4E78",
            autoFitColumns: true);

        Assert.True(applied.Success, applied.ErrorMessage);
        Assert.Equal("professional", applied.Preset);
        Assert.Equal(4, applied.Sections.Count);
        Assert.Equal(64, applied.Fingerprint.Length);
        var header = Assert.Single(applied.Sections, section => section.Name == "header");
        Assert.Equal("#1F4E78", header.FillColor);
        Assert.Equal("#FFFFFF", header.FontColor);
        Assert.True(header.Bold);
        var body = Assert.Single(applied.Sections, section => section.Name == "body");
        Assert.Equal("#FFFFFF", body.FillColor);
        Assert.False(body.Bold);

        var inspected = _commands.GetState(batch, "Sheet1", "A1:D1", "A2:D2", "A3:D5", "A6:D6");
        Assert.Equal(applied.Fingerprint, inspected.Fingerprint);

        var reapplied = _commands.Apply(
            batch,
            "Sheet1",
            "A1:D1",
            "A2:D2",
            "A3:D5",
            "A6:D6",
            ReportFormatPreset.Professional,
            "#1F4E78",
            autoFitColumns: true);
        Assert.Equal(applied.Fingerprint, reapplied.Fingerprint);

        var content = batch.Execute((ctx, ct) =>
        {
            object? sheet = null;
            object? formulaCell = null;
            object? numberCell = null;
            try
            {
                sheet = ctx.Book.Worksheets["Sheet1"];
                formulaCell = ((dynamic)sheet).Range["D3"];
                numberCell = ((dynamic)sheet).Range["B3"];
                return (
                    Formula: Convert.ToString(((dynamic)formulaCell).Formula, CultureInfo.InvariantCulture),
                    NumberFormat: Convert.ToString(((dynamic)numberCell).NumberFormat, CultureInfo.InvariantCulture),
                    Value: Convert.ToDouble(((dynamic)numberCell).Value2, CultureInfo.InvariantCulture));
            }
            finally
            {
                ComUtilities.Release(ref numberCell);
                ComUtilities.Release(ref formulaCell);
                ComUtilities.Release(ref sheet);
            }
        });

        Assert.Equal("=B3*C3", content.Formula);
        Assert.Equal("$#,##0.00", content.NumberFormat);
        Assert.Equal(10d, content.Value);
    }

    [Fact]
    public void Apply_WithOverlappingSections_FailsBeforeAnyFormattingMutation()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        SeedReport(batch);
        SetFill(batch, "A2:D2", 0x0000FF); // Excel BGR integer for red.

        Assert.Throws<ArgumentException>(() => _commands.Apply(
            batch,
            "Sheet1",
            null,
            "A2:D3",
            "A3:D5",
            null));

        var fill = GetFill(batch, "A2");
        Assert.Equal("#FF0000", FormattingHelpers.ColorToHex(fill));
    }

    private static void SeedReport(IExcelBatch batch)
    {
        batch.Execute((ctx, ct) =>
        {
            object? sheet = null;
            object? data = null;
            object? formulaRange = null;
            object? currencyRange = null;
            try
            {
                sheet = ctx.Book.Worksheets["Sheet1"];
                data = ((dynamic)sheet).Range["A1:D6"];
                ((dynamic)data).Value2 = new object[,]
                {
                    { "Quarterly Sales", null!, null!, null! },
                    { "Product", "Price", "Units", "Total" },
                    { "Alpha", 10d, 2d, null! },
                    { "Beta", 12d, 3d, null! },
                    { "Gamma", 8d, 4d, null! },
                    { "Total", null!, null!, null! },
                };
                formulaRange = ((dynamic)sheet).Range["D3:D5"];
                ((dynamic)formulaRange).Formula = new object[,]
                {
                    { "=B3*C3" },
                    { "=B4*C4" },
                    { "=B5*C5" },
                };
                currencyRange = ((dynamic)sheet).Range["B3:B5"];
                ((dynamic)currencyRange).NumberFormat = "$#,##0.00";
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref currencyRange);
                ComUtilities.Release(ref formulaRange);
                ComUtilities.Release(ref data);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static void SetFill(IExcelBatch batch, string address, int color)
    {
        batch.Execute((ctx, ct) =>
        {
            object? sheet = null;
            object? range = null;
            object? interior = null;
            try
            {
                sheet = ctx.Book.Worksheets["Sheet1"];
                range = ((dynamic)sheet).Range[address];
                interior = ((dynamic)range).Interior;
                ((dynamic)interior).Color = color;
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref interior);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static int GetFill(IExcelBatch batch, string address) => batch.Execute((ctx, ct) =>
    {
        object? sheet = null;
        object? range = null;
        object? interior = null;
        try
        {
            sheet = ctx.Book.Worksheets["Sheet1"];
            range = ((dynamic)sheet).Range[address];
            interior = ((dynamic)range).Interior;
            return Convert.ToInt32(((dynamic)interior).Color, CultureInfo.InvariantCulture);
        }
        finally
        {
            ComUtilities.Release(ref interior);
            ComUtilities.Release(ref range);
            ComUtilities.Release(ref sheet);
        }
    });
}
