namespace Sbroenne.ExcelMcp.Core.Utilities;

internal static class ExcelValueNormalizer
{
    internal static ExcelValueGrid Normalize(object? value)
    {
        if (value is not Array array)
        {
            return new ExcelValueGrid([[value]], 1, 1);
        }

        if (array.Rank != 2)
        {
            throw new InvalidOperationException($"Excel returned an unsupported {array.Rank}-dimensional value array.");
        }

        int rowLower = array.GetLowerBound(0);
        int rowUpper = array.GetUpperBound(0);
        int columnLower = array.GetLowerBound(1);
        int columnUpper = array.GetUpperBound(1);
        var rows = new List<List<object?>>(array.GetLength(0));

        for (int row = rowLower; row <= rowUpper; row++)
        {
            var values = new List<object?>(array.GetLength(1));
            for (int column = columnLower; column <= columnUpper; column++)
            {
                values.Add(array.GetValue(row, column));
            }

            rows.Add(values);
        }

        return new ExcelValueGrid(rows, array.GetLength(0), array.GetLength(1));
    }
}

internal sealed record ExcelValueGrid(
    List<List<object?>> Values,
    int RowCount,
    int ColumnCount);
