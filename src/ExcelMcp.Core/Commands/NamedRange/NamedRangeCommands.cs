namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Named range/parameter management commands implementation
/// </summary>
public partial class NamedRangeCommands : INamedRangeCommands
{
    private static List<List<object?>> ConvertArrayToList(object[,] array2D)
    {
        var result = new List<List<object?>>();

        var rowLower = array2D.GetLowerBound(0);
        var rowUpper = array2D.GetUpperBound(0);
        var colLower = array2D.GetLowerBound(1);
        var colUpper = array2D.GetUpperBound(1);

        for (var row = rowLower; row <= rowUpper; row++)
        {
            var rowList = new List<object?>();
            for (var col = colLower; col <= colUpper; col++)
            {
                rowList.Add(ConvertValueForJson(array2D[row, col]));
            }
            result.Add(rowList);
        }

        return result;
    }

    private static object? ConvertValueForJson(object? value)
    {
        if (value == null || value == DBNull.Value)
        {
            return null;
        }

        return value switch
        {
            string or bool or int or long or decimal => value,
            DateTime dateTime => dateTime.ToString("O", System.Globalization.CultureInfo.InvariantCulture),
            double doubleValue when double.IsNaN(doubleValue) || double.IsInfinity(doubleValue)
                => doubleValue.ToString(System.Globalization.CultureInfo.InvariantCulture),
            float floatValue when float.IsNaN(floatValue) || float.IsInfinity(floatValue)
                => floatValue.ToString(System.Globalization.CultureInfo.InvariantCulture),
            double or float => value,
            _ => value.ToString()
        };
    }
}



