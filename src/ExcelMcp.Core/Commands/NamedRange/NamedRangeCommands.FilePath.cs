using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Named range operations - FilePath-based API implementations
/// </summary>
public partial class NamedRangeCommands
{
    /// <inheritdoc />
    public async Task<NamedRangeListResult> ListAsync(string filePath)
    {
        var result = new NamedRangeListResult { FilePath = filePath };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? namesCollection = null;
                try
                {
                    namesCollection = handle.Workbook.Names;
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

                            result.NamedRanges.Add(new NamedRangeInfo
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
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = ex.Message;
                }
                finally
                {
                    ComUtilities.Release(ref namesCollection);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
    }

    /// <inheritdoc />
    public async Task<OperationResult> SetAsync(string filePath, string paramName, string value)
    {
        var result = new OperationResult { FilePath = filePath, Action = "set-parameter" };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? nameObj = null;
                dynamic? refersToRange = null;
                try
                {
                    nameObj = ComUtilities.FindName(handle.Workbook, paramName);
                    if (nameObj == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Named range '{paramName}' not found";
                        return;
                    }

                    refersToRange = nameObj.RefersToRange;
                    if (refersToRange == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Named range '{paramName}' does not refer to a valid range";
                        return;
                    }

                    refersToRange.Value2 = value;
                    result.Success = true;
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = ex.Message;
                }
                finally
                {
                    ComUtilities.Release(ref refersToRange);
                    ComUtilities.Release(ref nameObj);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
    }

    /// <inheritdoc />
    public async Task<NamedRangeValueResult> GetAsync(string filePath, string paramName)
    {
        var result = new NamedRangeValueResult { FilePath = filePath, NamedRangeName = paramName };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? nameObj = null;
                dynamic? refersToRange = null;
                try
                {
                    nameObj = ComUtilities.FindName(handle.Workbook, paramName);
                    if (nameObj == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Named range '{paramName}' not found";
                        return;
                    }

                    string refersTo = nameObj.RefersTo ?? "";
                    result.RefersTo = refersTo;

                    try
                    {
                        refersToRange = nameObj.RefersToRange;
                        var rawValue = refersToRange?.Value2;

                        if (rawValue is object[,] array2D)
                        {
                            result.Value = ConvertArrayToList(array2D);
                            result.ValueType = "Array";
                        }
                        else
                        {
                            result.Value = rawValue;
                            result.ValueType = rawValue?.GetType().Name ?? "null";
                        }
                    }
                    catch (Exception ex)
                    {
                        result.ErrorMessage = $"Warning: Could not read value: {ex.Message}";
                    }

                    result.Success = true;
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = ex.Message;
                }
                finally
                {
                    ComUtilities.Release(ref refersToRange);
                    ComUtilities.Release(ref nameObj);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
    }

    /// <inheritdoc />
    public async Task<OperationResult> UpdateAsync(string filePath, string paramName, string reference)
    {
        var result = new OperationResult { FilePath = filePath, Action = "update-parameter" };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? nameObj = null;
                try
                {
                    nameObj = ComUtilities.FindName(handle.Workbook, paramName);
                    if (nameObj == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Named range '{paramName}' not found";
                        return;
                    }

                    string normalizedRef = reference.StartsWith('=') ? reference : $"={reference}";
                    nameObj.RefersTo = normalizedRef;
                    result.Success = true;
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = ex.Message;
                }
                finally
                {
                    ComUtilities.Release(ref nameObj);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
    }

    /// <inheritdoc />
    public async Task<OperationResult> CreateAsync(string filePath, string paramName, string reference)
    {
        var result = new OperationResult { FilePath = filePath, Action = "create-parameter" };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? namesCollection = null;
                dynamic? newName = null;
                try
                {
                    if (paramName.Length > 255)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Parameter name exceeds Excel's 255-character limit ({paramName.Length} characters)";
                        return;
                    }

                    namesCollection = handle.Workbook.Names;
                    string normalizedRef = reference.StartsWith('=') ? reference : $"={reference}";

                    newName = namesCollection.Add(Name: paramName, RefersTo: normalizedRef);
                    result.Success = true;
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = ex.Message;
                }
                finally
                {
                    ComUtilities.Release(ref newName);
                    ComUtilities.Release(ref namesCollection);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
    }

    /// <inheritdoc />
    public async Task<OperationResult> DeleteAsync(string filePath, string paramName)
    {
        var result = new OperationResult { FilePath = filePath, Action = "delete-parameter" };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? nameObj = null;
                try
                {
                    nameObj = ComUtilities.FindName(handle.Workbook, paramName);
                    if (nameObj == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Named range '{paramName}' not found";
                        return;
                    }

                    nameObj.Delete();
                    result.Success = true;
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = ex.Message;
                }
                finally
                {
                    ComUtilities.Release(ref nameObj);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
    }

    /// <inheritdoc />
    public async Task<OperationResult> CreateBulkAsync(string filePath, IEnumerable<NamedRangeDefinition> parameters)
    {
        var parameterList = parameters?.ToList();

        if (parameterList == null || parameterList.Count == 0)
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = "No parameters provided",
                FilePath = filePath
            };
        }

        var createdCount = 0;
        var errors = new List<string>();

        foreach (var param in parameterList)
        {
            // Validate parameter
            if (string.IsNullOrWhiteSpace(param.Name))
            {
                errors.Add("Parameter with empty name skipped");
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
            var createResult = await CreateAsync(filePath, param.Name, param.Reference);
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
                var setResult = await SetAsync(filePath, param.Name, valueStr);
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
                FilePath = filePath
            };
        }

        var result = new OperationResult
        {
            Success = true,
            FilePath = filePath
        };

        if (errors.Count > 0)
        {
            result.ErrorMessage = string.Join("; ", errors);
        }

        return result;
    }
}
