using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

public partial class ExcelBatchTests
{
    [Fact]
    public void Execute_TypedChartDefinitions_PreserveGeometryMaterialsAndForecasts()
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);
        batch.Execute((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? source = null;
            Excel.Shapes? shapes = null;
            Excel.Shape? shape = null;
            Excel.Chart? chart = null;
            Excel.PivotLayout? pivotLayout = null;
            Excel.Range? anchor = null;
            Excel.ChartArea? area = null;
            Excel.Chart? parent = null;
            Excel.ChartFormat? format = null;
            Excel.FillFormat? fill = null;
            Excel.ColorFormat? fillColor = null;
            Excel.LineFormat? line = null;
            Excel.ColorFormat? lineColor = null;
            Excel.SeriesCollection? series = null;
            Excel.Series? first = null;
            Excel.Trendlines? trendlines = null;
            Excel.Trendline? trendline = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[1];
                source = sheet.Range["A1:C4"];
                source.Value2 = new object[,]
                {
                    { "Quarter", "Alpha", "Beta" },
                    { "Q1", 17, 29 },
                    { "Q2", 31, 43 },
                    { "Q3", 47, 59 }
                };
                shapes = sheet.Shapes;
                shape = shapes.AddChart((int)Excel.XlChartType.xlColumnClustered, 100, 50, 400, 300);
                chart = shape.Chart;
                chart.SetSourceData(source);
                Assert.Equal(100f, shape.Left);
                Assert.Equal(50f, shape.Top);
                Assert.Equal(400f, shape.Width);
                Assert.Equal(300f, shape.Height);
                anchor = shape.TopLeftCell;
                Assert.NotEmpty(anchor.Address);
                pivotLayout = chart.PivotLayout;
                Assert.Null(pivotLayout);
                area = chart.ChartArea;
                parent = (Excel.Chart)area.Parent;
                Assert.Equal(chart.Name, parent.Name);
                format = area.Format;
                fill = format.Fill;
                fill.Solid();
                fillColor = fill.ForeColor;
                fillColor.RGB = 0x0000FF;
                fill.Transparency = 0.25f;
                line = format.Line;
                line.Visible = Office.MsoTriState.msoTrue;
                lineColor = line.ForeColor;
                lineColor.RGB = 0xFF0000;
                line.Weight = 2.5f;
                Assert.Equal(0x0000FF, fillColor.RGB);
                Assert.Equal(0.25f, fill.Transparency);
                Assert.Equal(0xFF0000, lineColor.RGB);
                Assert.Equal(2.5f, line.Weight);
                series = (Excel.SeriesCollection)chart.SeriesCollection();
                Assert.Equal(2, series.Count);
                first = series.Item(1);
                object values = first.Values;
                object categories = first.XValues;
                Assert.Equal([17d, 31d, 47d], Assert.IsAssignableFrom<Array>(values).Cast<object>()
                    .Select(value => Convert.ToDouble(value, CultureInfo.InvariantCulture)));
                Assert.Equal(["Q1", "Q2", "Q3"], Assert.IsAssignableFrom<Array>(categories).Cast<object>());
                trendlines = (Excel.Trendlines)first.Trendlines();
                trendline = trendlines.Add(Excel.XlTrendlineType.xlLinear);
                WriteNativeForecastDefinitions(trendline);
                trendline.Forward2 = 2.5;
                trendline.Backward2 = 0.5;
                Assert.Equal(2.5, trendline.Forward2);
                Assert.Equal(0.5, trendline.Backward2);
                Assert.Equal(2, trendline.Forward);
                Assert.Equal(0, trendline.Backward);
                ReadOnlySpan<(double Value, int Rounded)> forecasts = [(0.5, 0), (1.5, 2), (2.5, 2)];
                foreach (var (value, rounded) in forecasts)
                {
                    trendline.Forward2 = value;
                    Assert.Equal(rounded, trendline.Forward);
                    trendline.Forward = Convert.ToInt32(value);
                    Assert.Equal((double)rounded, trendline.Forward2);
                }
                trendline.Backward = 0;
                Assert.Equal(0, trendline.Backward);
                chart.ChartType = Excel.XlChartType.xlXYScatter;
                first.XValues = (double[])[10, 20, 30];
                foreach (var (value, rounded) in forecasts)
                {
                    trendline.Backward2 = value;
                    Assert.Equal(value, trendline.Backward2);
                    Assert.Equal(rounded, trendline.Backward);
                    trendline.Backward = Convert.ToInt32(value);
                    Assert.Equal((double)rounded, trendline.Backward2);
                }
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref trendline);
                ComUtilities.Release(ref trendlines);
                ComUtilities.Release(ref first);
                ComUtilities.Release(ref series);
                ComUtilities.Release(ref lineColor);
                ComUtilities.Release(ref line);
                ComUtilities.Release(ref fillColor);
                ComUtilities.Release(ref fill);
                ComUtilities.Release(ref format);
                ComUtilities.Release(ref parent);
                ComUtilities.Release(ref area);
                ComUtilities.Release(ref anchor);
                ComUtilities.Release(ref pivotLayout);
                ComUtilities.Release(ref chart);
                ComUtilities.Release(ref shape);
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref source);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        AssertNoExternalInteropAssemblies();
    }

    private void WriteNativeForecastDefinitions(Excel.Trendline trendline)
    {
        ITypeInfo? information = null;
        try
        {
            ((IAutomationTypeInformation)trendline).GetTypeInfo(0, 0, out information);
            information.GetTypeAttr(out nint attributes);
            try
            {
                var type = Marshal.PtrToStructure<TYPEATTR>(attributes);
                for (int index = 0; index < type.cFuncs; index++)
                {
                    information.GetFuncDesc(index, out nint descriptor);
                    try
                    {
                        var function = Marshal.PtrToStructure<FUNCDESC>(descriptor);
                        var names = new string[1];
                        information.GetNames(function.memid, names, names.Length, out int count);
                        if (count == 1 && (names[0].Contains("Forward", StringComparison.Ordinal) ||
                            names[0].Contains("Backward", StringComparison.Ordinal)))
                        {
                            _output.WriteLine($"{names[0]} {function.invkind}: " +
                                $"{(VarEnum)function.elemdescFunc.tdesc.vt}; flags {function.wFuncFlags}.");
                        }
                    }
                    finally
                    {
                        information.ReleaseFuncDesc(descriptor);
                    }
                }
            }
            finally
            {
                information.ReleaseTypeAttr(attributes);
            }
        }
        finally
        {
            ComUtilities.Release(ref information);
        }
    }
}
