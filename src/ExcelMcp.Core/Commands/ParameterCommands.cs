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
            try
            {
                dynamic namesCollection = workbook.Names;
                int count = namesCollection.Count;
                
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic nameObj = namesCollection.Item(i);
                        string name = nameObj.Name;
                        string refersTo = nameObj.RefersTo ?? "";
                        
                        // Try to get value
                        object? value = null;
                        try
                        {
                            value = nameObj.RefersToRange?.Value2;
                        }
                        catch { }
                        
                        result.Parameters.Add(new ParameterInfo
                        {
                            Name = name,
                            RefersTo = refersTo,
                            Value = value,
                            ValueType = value?.GetType().Name ?? "null"
                        });
                    }
                    catch { }
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
            try
            {
                dynamic? nameObj = FindNamedRange(workbook, paramName);
                if (nameObj == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Parameter '{paramName}' not found";
                    return 1;
                }

                dynamic refersToRange = nameObj.RefersToRange;
                
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
            try
            {
                dynamic? nameObj = FindNamedRange(workbook, paramName);
                if (nameObj == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Parameter '{paramName}' not found";
                    return 1;
                }

                result.RefersTo = nameObj.RefersTo ?? "";
                result.Value = nameObj.RefersToRange?.Value2;
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
            try
            {
                // Check if parameter already exists
                dynamic? existing = FindNamedRange(workbook, paramName);
                if (existing != null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Parameter '{paramName}' already exists";
                    return 1;
                }

                // Create new named range
                dynamic namesCollection = workbook.Names;
                namesCollection.Add(paramName, reference);

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
            try
            {
                dynamic? nameObj = FindNamedRange(workbook, paramName);
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
}
