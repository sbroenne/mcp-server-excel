using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Named range lifecycle operations (List, Set, Get, Create, Update, Delete, CreateBulk)
/// </summary>
public partial class NamedRangeCommands
{
    /// <inheritdoc />
    public NamedRangeListResult List(IExcelBatch batch)
    {
        var result = new NamedRangeListResult { FilePath = batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
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
                return result;
            }
            finally
            {
                ComUtilities.Release(ref namesCollection);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Write(IExcelBatch batch, string paramName, string value)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "set-parameter" };

        return batch.Execute((ctx, ct) =>
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
            finally
            {
                ComUtilities.Release(ref refersToRange);
                ComUtilities.Release(ref nameObj);
            }
        });
    }

    /// <inheritdoc />
    public NamedRangeValueResult Read(IExcelBatch batch, string paramName)
    {
        var result = new NamedRangeValueResult { FilePath = batch.WorkbookPath, NamedRangeName = paramName };

        return batch.Execute((ctx, ct) =>
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
            finally
            {
                ComUtilities.Release(ref refersToRange);
                ComUtilities.Release(ref nameObj);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Create(IExcelBatch batch, string paramName, string reference)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "create-parameter" };

        // Validate parameter name length (Excel limit: 255 characters)
        if (string.IsNullOrWhiteSpace(paramName))
        {
            result.Success = false;
            result.ErrorMessage = "Parameter name cannot be empty or whitespace";
            return result;
        }

        if (paramName.Length > 255)
        {
            result.Success = false;
            result.ErrorMessage = $"Parameter name exceeds Excel's 255-character limit (current length: {paramName.Length})";
            return result;
        }

        return batch.Execute((ctx, ct) =>
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
            finally
            {
                ComUtilities.Release(ref namesCollection);
                ComUtilities.Release(ref existing);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Update(IExcelBatch batch, string paramName, string reference)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "update-parameter" };

        // Validate parameter name length (Excel limit: 255 characters)
        if (string.IsNullOrWhiteSpace(paramName))
        {
            result.Success = false;
            result.ErrorMessage = "Parameter name cannot be empty or whitespace";
            return result;
        }

        if (paramName.Length > 255)
        {
            result.Success = false;
            result.ErrorMessage = $"Parameter name exceeds Excel's 255-character limit (current length: {paramName.Length})";
            return result;
        }

        return batch.Execute((ctx, ct) =>
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

                // Remove any existing = prefix to avoid double ==
                string formattedReference = reference.TrimStart('=');
                // Add exactly one = prefix (required by Excel COM API)
                formattedReference = $"={formattedReference}";

                // Update the reference
                nameObj.RefersTo = formattedReference;

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref nameObj);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Delete(IExcelBatch batch, string paramName)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "delete-parameter" };

        return batch.Execute((ctx, ct) =>
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
            finally
            {
                ComUtilities.Release(ref nameObj);
            }
        });
    }
}

