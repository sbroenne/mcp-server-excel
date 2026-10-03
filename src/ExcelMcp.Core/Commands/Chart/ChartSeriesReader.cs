using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

internal static class ChartSeriesReader
{
    internal static int Count(Excel.Chart chart)
    {
        Excel.SeriesCollection? collection = null;
        try
        {
            collection = (Excel.SeriesCollection)chart.SeriesCollection();
            return collection.Count;
        }
        finally
        {
            ComUtilities.Release(ref collection);
        }
    }

    internal static List<SeriesInfo> Read(Excel.Chart chart)
    {
        var result = new List<SeriesInfo>();
        Excel.SeriesCollection? collection = null;
        try
        {
            collection = (Excel.SeriesCollection)chart.SeriesCollection();
            var count = collection.Count;
            for (var index = 1; index <= count; index++)
            {
                Excel.Series? series = null;
                try
                {
                    series = collection.Item(index);
                    var values = series.Values;
                    var categories = series.XValues;
                    result.Add(new SeriesInfo
                    {
                        ChartType = (ChartType)Convert.ToInt32(series.ChartType, CultureInfo.InvariantCulture),
                        AxisGroup = (ChartAxisGroup)Convert.ToInt32(series.AxisGroup, CultureInfo.InvariantCulture),
                        Name = series.Name,
                        Values = ToList(values),
                        Categories = ToList(categories)
                    });
                }
                finally
                {
                    ComUtilities.Release(ref series);
                }
            }
            return result;
        }
        finally
        {
            ComUtilities.Release(ref collection);
        }
    }

    private static List<object?> ToList(object? value) =>
        value switch
        {
            null => [],
            Array array => array.Cast<object?>().ToList(),
            _ => [value]
        };
}
