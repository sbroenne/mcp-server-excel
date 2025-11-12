using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Named range/parameter management commands - FilePath-based API implementations
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
                dynamic? names = null;
                try
                {
                    names = handle.Workbook.Names;
                    for (int i = 1; i <= names.Count; i++)
                    {
                        dynamic? name = null;
                        try
                        {
                            name = names.Item(i);
                            string nameName = name.Name;
                            string refersTo = name.RefersTo;

                            result.NamedRanges.Add(new NamedRangeInfo
                            {
                                Name = nameName,
                                RefersTo = refersTo
                            });
                        }
                        finally
                        {
                            ComUtilities.Release(ref name);
                        }
                    }
                    result.Success = true;
                }
                finally
                {
                    ComUtilities.Release(ref names);
                }
            });
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to list named ranges: {ex.Message}";
        }

        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> SetAsync(string filePath, string paramName, string value)
    {
        var result = new OperationResult { FilePath = filePath };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? names = null;
                dynamic? targetName = null;
                dynamic? refRange = null;

                try
                {
                    names = handle.Workbook.Names;
                    targetName = names.Item(paramName);
                    refRange = targetName.RefersToRange;

                    if (refRange.Cells.Count == 1)
                    {
                        refRange.Value2 = value;
                    }
                    else
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Named range '{paramName}' refers to multiple cells. Use excel_range for multi-cell operations.";
                        return;
                    }

                    result.Success = true;
                }
                finally
                {
                    ComUtilities.Release(ref refRange);
                    ComUtilities.Release(ref targetName);
                    ComUtilities.Release(ref names);
                }
            });

            // Auto-save after write operation
            if (result.Success)
            {
                await FileHandleManager.Instance.SaveAsync(filePath);
            }
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to set named range value: {ex.Message}";
        }

        return result;
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
                dynamic? names = null;
                dynamic? targetName = null;
                dynamic? refRange = null;

                try
                {
                    names = handle.Workbook.Names;
                    targetName = names.Item(paramName);
                    refRange = targetName.RefersToRange;

                    result.RefersTo = targetName.RefersTo ?? string.Empty;
                    result.Value = refRange?.Value2;
                    result.ValueType = result.Value?.GetType().Name ?? "null";
                    result.Success = true;
                }
                finally
                {
                    ComUtilities.Release(ref refRange);
                    ComUtilities.Release(ref targetName);
                    ComUtilities.Release(ref names);
                }
            });
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to get named range value: {ex.Message}";
        }

        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> UpdateAsync(string filePath, string paramName, string reference)
    {
        var result = new OperationResult { FilePath = filePath };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? names = null;
                dynamic? targetName = null;

                try
                {
                    names = handle.Workbook.Names;
                    targetName = names.Item(paramName);

                    string normalizedRef = reference.StartsWith('=') ? reference : $"={reference}";
                    targetName.RefersTo = normalizedRef;

                    result.Success = true;
                }
                finally
                {
                    ComUtilities.Release(ref targetName);
                    ComUtilities.Release(ref names);
                }
            });

            // Auto-save after write operation
            if (result.Success)
            {
                await FileHandleManager.Instance.SaveAsync(filePath);
            }
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to update named range reference: {ex.Message}";
        }

        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> CreateAsync(string filePath, string paramName, string reference)
    {
        var result = new OperationResult { FilePath = filePath };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? names = null;
                dynamic? newName = null;

                try
                {
                    names = handle.Workbook.Names;

                    string normalizedRef = reference.StartsWith('=') ? reference : $"={reference}";
                    newName = names.Add(Name: paramName, RefersTo: normalizedRef);

                    result.Success = true;
                }
                finally
                {
                    ComUtilities.Release(ref newName);
                    ComUtilities.Release(ref names);
                }
            });

            // Auto-save after write operation
            if (result.Success)
            {
                await FileHandleManager.Instance.SaveAsync(filePath);
            }
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to create named range: {ex.Message}";
        }

        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> DeleteAsync(string filePath, string paramName)
    {
        var result = new OperationResult { FilePath = filePath };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? names = null;
                dynamic? targetName = null;

                try
                {
                    names = handle.Workbook.Names;
                    targetName = names.Item(paramName);
                    targetName.Delete();

                    result.Success = true;
                }
                finally
                {
                    ComUtilities.Release(ref targetName);
                    ComUtilities.Release(ref names);
                }
            });

            // Auto-save after write operation
            if (result.Success)
            {
                await FileHandleManager.Instance.SaveAsync(filePath);
            }
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to delete named range: {ex.Message}";
        }

        return result;
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

        var result = new OperationResult { FilePath = filePath };
        var createdCount = 0;
        var errors = new List<string>();

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? names = null;

                try
                {
                    names = handle.Workbook.Names;

                    foreach (var param in parameterList)
                    {
                        dynamic? newName = null;
                        dynamic? refRange = null;

                        try
                        {
                            string normalizedRef = param.Reference.StartsWith('=') ? param.Reference : $"={param.Reference}";
                            newName = names.Add(Name: param.Name, RefersTo: normalizedRef);
                            createdCount++;

                            if (param.Value != null)
                            {
                                refRange = newName.RefersToRange;
                                if (refRange.Cells.Count == 1)
                                {
                                    refRange.Value2 = param.Value;
                                }
                                else
                                {
                                    errors.Add($"Warning: '{param.Name}' created but value not set (multi-cell range)");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            errors.Add($"Failed to create '{param.Name}': {ex.Message}");
                        }
                        finally
                        {
                            ComUtilities.Release(ref refRange);
                            ComUtilities.Release(ref newName);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref names);
                }
            });

            if (createdCount > 0)
            {
                result.Success = true;
                result.ErrorMessage = errors.Count > 0
                    ? $"Created {createdCount}/{parameterList.Count} parameters. Errors: {string.Join("; ", errors)}"
                    : $"Created {createdCount} named range(s) successfully";

                // Auto-save after write operation
                await FileHandleManager.Instance.SaveAsync(filePath);
            }
            else
            {
                result.Success = false;
                result.ErrorMessage = $"Failed to create any parameters. Errors: {string.Join("; ", errors)}";
            }
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to create bulk named ranges: {ex.Message}";
        }

        return result;
    }
}
