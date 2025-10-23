using Sbroenne.ExcelMcp.Core.Models;
using static Sbroenne.ExcelMcp.Core.ExcelHelper;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Named range/parameter management commands implementation
/// </summary>
public class ParameterCommands : IParameterCommands
{
    /// <inheritdoc />
    public ParameterListResult List(string filePath)
    {
        if (!File.Exists(filePath))
        {
            return new ParameterListResult
            {
                Success = false,
                ErrorMessage = $"File not found: {filePath}",
                FilePath = filePath
            };
        }

        var result = new ParameterListResult { FilePath = filePath };

        WithExcel(filePath, false, (excel, workbook) =>
        {
            dynamic? namesCollection = null;
            try
            {
                namesCollection = workbook.Names;
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
                        ReleaseComObject(ref refersToRange);
                        ReleaseComObject(ref nameObj);
                    }
                }

                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return 1;
            }
            finally
            {
                ReleaseComObject(ref namesCollection);
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult Set(string filePath, string paramName, string value)
    {
        if (!File.Exists(filePath))
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"File not found: {filePath}",
                FilePath = filePath,
                Action = "set-parameter"
            };
        }

        var result = new OperationResult { FilePath = filePath, Action = "set-parameter" };

        WithExcel(filePath, true, (excel, workbook) =>
        {
            dynamic? nameObj = null;
            dynamic? refersToRange = null;
            try
            {
                nameObj = FindName(workbook, paramName);
                if (nameObj == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Parameter '{paramName}' not found";
                    return 1;
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

                workbook.Save();
                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return 1;
            }
            finally
            {
                ReleaseComObject(ref refersToRange);
                ReleaseComObject(ref nameObj);
            }
        });

        return result;
    }

    /// <inheritdoc />
    public ParameterValueResult Get(string filePath, string paramName)
    {
        if (!File.Exists(filePath))
        {
            return new ParameterValueResult
            {
                Success = false,
                ErrorMessage = $"File not found: {filePath}",
                FilePath = filePath,
                ParameterName = paramName
            };
        }

        var result = new ParameterValueResult { FilePath = filePath, ParameterName = paramName };

        WithExcel(filePath, false, (excel, workbook) =>
        {
            dynamic? nameObj = null;
            dynamic? refersToRange = null;
            try
            {
                nameObj = FindName(workbook, paramName);
                if (nameObj == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Parameter '{paramName}' not found";
                    return 1;
                }

                result.RefersTo = nameObj.RefersTo ?? "";
                refersToRange = nameObj.RefersToRange;
                result.Value = refersToRange?.Value2;
                result.ValueType = result.Value?.GetType().Name ?? "null";
                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return 1;
            }
            finally
            {
                ReleaseComObject(ref refersToRange);
                ReleaseComObject(ref nameObj);
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult Create(string filePath, string paramName, string reference)
    {
        if (!File.Exists(filePath))
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"File not found: {filePath}",
                FilePath = filePath,
                Action = "create-parameter"
            };
        }

        var result = new OperationResult { FilePath = filePath, Action = "create-parameter" };

        WithExcel(filePath, true, (excel, workbook) =>
        {
            dynamic? existing = null;
            dynamic? namesCollection = null;
            try
            {
                // Check if parameter already exists
                existing = FindNamedRange(workbook, paramName);
                if (existing != null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Parameter '{paramName}' already exists";
                    return 1;
                }

                // Create new named range
                namesCollection = workbook.Names;
                // Ensure reference is properly formatted for Excel COM
                string formattedReference = reference.StartsWith("=") ? reference : $"={reference}";
                namesCollection.Add(paramName, formattedReference);

                workbook.Save();
                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return 1;
            }
            finally
            {
                ReleaseComObject(ref namesCollection);
                ReleaseComObject(ref existing);
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult Update(string filePath, string paramName, string reference)
    {
        if (!File.Exists(filePath))
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"File not found: {filePath}",
                FilePath = filePath,
                Action = "update-parameter"
            };
        }

        var result = new OperationResult { FilePath = filePath, Action = "update-parameter" };

        WithExcel(filePath, true, (excel, workbook) =>
        {
            dynamic? nameObj = null;
            try
            {
                nameObj = FindName(workbook, paramName);
                if (nameObj == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Parameter '{paramName}' not found";
                    result.SuggestedNextActions = new List<string>
                    {
                        "Use 'param-list' to see available parameters",
                        "Use 'param-create' to create a new named range"
                    };
                    return 1;
                }

                // Ensure reference is properly formatted with = prefix
                string formattedReference = reference.StartsWith("=") ? reference : $"={reference}";
                
                // Update the reference
                nameObj.RefersTo = formattedReference;
                
                result.Success = true;
                result.SuggestedNextActions = new List<string>
                {
                    $"Parameter '{paramName}' reference updated to '{reference}'",
                    "Use 'param-get' to verify new value",
                    "Use 'param-set' to change the value"
                };
                result.WorkflowHint = "Parameter reference updated. Next, verify or modify the value.";
                
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error updating parameter: {ex.Message}";
                result.SuggestedNextActions = new List<string>
                {
                    "Check that reference is valid (e.g., 'Sheet1!A1' or '=Sheet1!A1')",
                    "Ensure referenced sheet and cells exist"
                };
                return 1;
            }
            finally
            {
                ReleaseComObject(ref nameObj);
            }
        });

        return result;
    }

    /// <inheritdoc />
    public OperationResult Delete(string filePath, string paramName)
    {
        if (!File.Exists(filePath))
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"File not found: {filePath}",
                FilePath = filePath,
                Action = "delete-parameter"
            };
        }

        var result = new OperationResult { FilePath = filePath, Action = "delete-parameter" };

        WithExcel(filePath, true, (excel, workbook) =>
        {
            dynamic? nameObj = null;
            try
            {
                nameObj = FindName(workbook, paramName);
                if (nameObj == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Parameter '{paramName}' not found";
                    return 1;
                }

                nameObj.Delete();
                workbook.Save();
                result.Success = true;
                return 0;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return 1;
            }
            finally
            {
                ReleaseComObject(ref nameObj);
            }
        });

        return result;
    }

    private static dynamic? FindNamedRange(dynamic workbook, string name)
    {
        try
        {
            dynamic namesCollection = workbook.Names;
            int count = namesCollection.Count;

            for (int i = 1; i <= count; i++)
            {
                dynamic nameObj = namesCollection.Item(i);
                if (nameObj.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                {
                    return nameObj;
                }
            }
        }
        catch { }

        return null;
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
}
