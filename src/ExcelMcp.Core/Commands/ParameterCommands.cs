using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

#pragma warning disable CS1998 // Async method lacks 'await' operators - intentional for COM synchronous operations

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Named range/parameter management commands implementation
/// </summary>
public class ParameterCommands : IParameterCommands
{
    /// <inheritdoc />
    public async Task<ParameterListResult> ListAsync(IExcelBatch batch)
    {
        var result = new ParameterListResult { FilePath = batch.WorkbookPath };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? namesCollection = null;
            try
            {
                namesCollection = ctx.Book.Names;
                int count = namesCollection.Count;

                for (int i = 1; i <= count; i++)
                {
                    dynamic? nameObj = null;
                    dynamic? refersToRange = null;
                    try
                    {
                        nameObj = namesCollection.Item(i);
                        string name = nameObj.Name;
                        string refersTo = nameObj.RefersTo ?? "";

                        // Try to get value
                        object? value = null;
                        string valueType = "null";
                        try
                        {
                            refersToRange = nameObj.RefersToRange;
                            var rawValue = refersToRange?.Value2;

                            // Convert 2D array to List<List<object?>> for JSON serialization
                            if (rawValue is object[,] array2D)
                            {
                                value = ConvertArrayToList(array2D);
                                valueType = "Array";
                            }
                            else
                            {
                                value = rawValue;
                                valueType = rawValue?.GetType().Name ?? "null";
                            }
                        }
                        catch { }

                        result.Parameters.Add(new ParameterInfo
                        {
                            Name = name,
                            RefersTo = refersTo,
                            Value = value,
                            ValueType = valueType
                        });
                    }
                    catch { }
                    finally
                    {
                        ComUtilities.Release(ref refersToRange);
                        ComUtilities.Release(ref nameObj);
                    }
                }

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref namesCollection);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> SetAsync(IExcelBatch batch, string paramName, string value)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "set-parameter" };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? nameObj = null;
            dynamic? refersToRange = null;
            try
            {
                nameObj = ComUtilities.FindName(ctx.Book, paramName);
                if (nameObj == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Parameter '{paramName}' not found";
                    return result;
                }

                refersToRange = nameObj.RefersToRange;

                // Try to parse as number, otherwise set as text
                if (double.TryParse(value, out double numValue))
                {
                    refersToRange.Value2 = numValue;
                }
                else if (bool.TryParse(value, out bool boolValue))
                {
                    refersToRange.Value2 = boolValue;
                }
                else
                {
                    refersToRange.Value2 = value;
                }

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref refersToRange);
                ComUtilities.Release(ref nameObj);
            }
        });
    }

    /// <inheritdoc />
    public async Task<ParameterValueResult> GetAsync(IExcelBatch batch, string paramName)
    {
        var result = new ParameterValueResult { FilePath = batch.WorkbookPath, ParameterName = paramName };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? nameObj = null;
            dynamic? refersToRange = null;
            try
            {
                nameObj = ComUtilities.FindName(ctx.Book, paramName);
                if (nameObj == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Parameter '{paramName}' not found";
                    return result;
                }

                result.RefersTo = nameObj.RefersTo ?? "";
                refersToRange = nameObj.RefersToRange;
                result.Value = refersToRange?.Value2;
                result.ValueType = result.Value?.GetType().Name ?? "null";
                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref refersToRange);
                ComUtilities.Release(ref nameObj);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> CreateAsync(IExcelBatch batch, string paramName, string reference)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "create-parameter" };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? existing = null;
            dynamic? namesCollection = null;
            try
            {
                // Check if parameter already exists
                existing = ComUtilities.FindName(ctx.Book, paramName);
                if (existing != null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Parameter '{paramName}' already exists";
                    return result;
                }

                // Create new named range
                namesCollection = ctx.Book.Names;
                // Remove any existing = prefix to avoid double ==
                string formattedReference = reference.TrimStart('=');
                // Add exactly one = prefix (required by Excel COM API)
                formattedReference = $"={formattedReference}";
                namesCollection.Add(paramName, formattedReference);

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref namesCollection);
                ComUtilities.Release(ref existing);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> UpdateAsync(IExcelBatch batch, string paramName, string reference)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "update-parameter" };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? nameObj = null;
            try
            {
                nameObj = ComUtilities.FindName(ctx.Book, paramName);
                if (nameObj == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Parameter '{paramName}' not found";
                    result.SuggestedNextActions =
                    [
                        "Use 'param-list' to see available parameters",
                        "Use 'param-create' to create a new named range"
                    ];
                    return result;
                }

                // Remove any existing = prefix to avoid double ==
                string formattedReference = reference.TrimStart('=');
                // Add exactly one = prefix (required by Excel COM API)
                formattedReference = $"={formattedReference}";

                // Update the reference
                nameObj.RefersTo = formattedReference;

                result.Success = true;
                result.SuggestedNextActions =
                [
                    $"Parameter '{paramName}' reference updated to '{reference}'",
                    "Use 'param-get' to verify new value",
                    "Use 'param-set' to change the value"
                ];
                result.WorkflowHint = "Parameter reference updated. Next, verify or modify the value.";

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error updating parameter: {ex.Message}";
                result.SuggestedNextActions =
                [
                    "Check that reference is valid (e.g., 'Sheet1!A1' or '=Sheet1!A1')",
                    "Ensure referenced sheet and cells exist"
                ];
                return result;
            }
            finally
            {
                ComUtilities.Release(ref nameObj);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> DeleteAsync(IExcelBatch batch, string paramName)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "delete-parameter" };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? nameObj = null;
            try
            {
                nameObj = ComUtilities.FindName(ctx.Book, paramName);
                if (nameObj == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Parameter '{paramName}' not found";
                    return result;
                }

                nameObj.Delete();
                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref nameObj);
            }
        });
    }

    /// <summary>
    /// Converts a 2D array from Excel to a serializable List of Lists
    /// </summary>
    /// <param name="array2D">The 2D array from Excel (1-based indexing)</param>
    /// <returns>List of Lists representation</returns>
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
    public async Task<OperationResult> CreateBulkAsync(IExcelBatch batch, List<ParameterDefinition> parameters)
    {
        if (parameters == null || parameters.Count == 0)
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

        foreach (var param in parameters)
        {
            // Validate parameter
            if (string.IsNullOrWhiteSpace(param.Name))
            {
                errors.Add($"Parameter with empty name skipped");
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
                var setResult = await SetAsync(batch, param.Name, param.Value.ToString() ?? "");
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
            result.WorkflowHint = $"Created {createdCount} of {parameters.Count} parameter(s). Partial failures: {string.Join("; ", errors)}";
            result.SuggestedNextActions = 
            [
                $"Successfully created {createdCount} out of {parameters.Count} parameters",
                "Review errors to understand which parameters failed",
                "Use 'list' to verify created parameters"
            ];
        }
        else
        {
            result.WorkflowHint = $"Successfully created {createdCount} parameter(s) in bulk operation";
            result.SuggestedNextActions =
            [
                $"All {createdCount} parameters created successfully",
                "Parameters can now be used in formulas and Power Query code",
                "Use 'list' to view all parameters"
            ];
        }

        return result;
    }
}
