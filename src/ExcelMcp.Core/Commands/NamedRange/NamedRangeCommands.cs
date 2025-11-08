using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;


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

    /// <inheritdoc />
    public async Task<OperationResult> CreateBulkAsync(IExcelBatch batch, IEnumerable<NamedRangeDefinition> parameters)
    {
        var parameterList = parameters?.ToList();

        if (parameterList == null || parameterList.Count == 0)
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = "No parameters provided",
                FilePath = batch.WorkbookPath
            };
        }

        var createdCount = 0;
        var errors = new List<string>();

        foreach (var param in parameterList)
        {
            // Validate parameter
            if (string.IsNullOrWhiteSpace(param.Name))
            {
                errors.Add($"Parameter with empty name skipped");
                continue;
            }

            if (param.Name.Length > 255)
            {
                errors.Add($"Parameter '{param.Name}' exceeds Excel's 255-character limit ({param.Name.Length} characters) - skipped");
                continue;
            }

            if (string.IsNullOrWhiteSpace(param.Reference))
            {
                errors.Add($"Parameter '{param.Name}' has empty reference - skipped");
                continue;
            }

            // Create named range
            var createResult = await CreateAsync(batch, param.Name, param.Reference);
            if (!createResult.Success)
            {
                errors.Add($"Failed to create '{param.Name}': {createResult.ErrorMessage}");
                continue;
            }

            createdCount++;

            // Set value if provided
            if (param.Value != null)
            {
                var valueStr = Convert.ToString(param.Value, System.Globalization.CultureInfo.InvariantCulture) ?? "";
                var setResult = await SetAsync(batch, param.Name, valueStr);
                if (!setResult.Success)
                {
                    errors.Add($"Created '{param.Name}' but failed to set value: {setResult.ErrorMessage}");
                }
            }
        }

        if (createdCount == 0)
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"Failed to create any parameters. Errors: {string.Join("; ", errors)}",
                FilePath = batch.WorkbookPath
            };
        }

        var result = new OperationResult
        {
            Success = true,
            FilePath = batch.WorkbookPath
        };

        if (errors.Count > 0)
        {
            result.ErrorMessage = string.Join("; ", errors);
        }

        return result;
    }
}
