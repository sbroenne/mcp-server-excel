namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Named range/parameter management commands implementation
/// </summary>
public partial class NamedRangeCommands : INamedRangeCommands
{
    private static List<List<object?>> ConvertArrayToList(object[,] array2D)
    {
        var result = new List<List<object?>>();

        // Excel arrays are 1-based, get the bounds
        int rows = array2D.GetLength(0);
        int cols = array2D.GetLength(1);

        for (int row = 1; row <= rows; row++)
        {
            var rowList = new List<object?>();
            for (int col = 1; col <= cols; col++)
            {
                rowList.Add(array2D[row, col]);
            }
            result.Add(rowList);
        }

        return result;
    }
}
