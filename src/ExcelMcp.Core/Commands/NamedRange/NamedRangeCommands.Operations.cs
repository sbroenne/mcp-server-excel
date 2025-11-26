using System.Runtime.InteropServices;
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
    public List<NamedRangeInfo> List(IExcelBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            var namedRanges = new List<NamedRangeInfo>();
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
                        catch (COMException)
                        {
                            // Named range may not have a valid RefersToRange (e.g., formula-based or external reference)
                            // Continue with null value - this is expected for some named ranges
                        }

                        namedRanges.Add(new NamedRangeInfo
                        {
                            Name = name,
                            RefersTo = refersTo,
                            Value = value,
                            ValueType = valueType
                        });
                    }
                    catch (COMException)
                    {
                        // Skip corrupted or inaccessible named ranges - continue listing remaining
                        continue;
                    }
                    finally
                    {
                        ComUtilities.Release(ref refersToRange);
                        ComUtilities.Release(ref nameObj);
                    }
                }

                return namedRanges;
            }
            finally
            {
                ComUtilities.Release(ref namesCollection);
            }
        });
    }

    /// <inheritdoc />
    public void Write(IExcelBatch batch, string paramName, string value)
    {
        batch.Execute((ctx, ct) =>
        {
            dynamic? nameObj = null;
            dynamic? refersToRange = null;
            try
            {
                nameObj = ComUtilities.FindName(ctx.Book, paramName);
                if (nameObj == null)
                {
                    throw new InvalidOperationException($"Parameter '{paramName}' not found");
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

                return 0; // Dummy return for batch.Execute
            }
            finally
            {
                ComUtilities.Release(ref refersToRange);
                ComUtilities.Release(ref nameObj);
            }
        });
    }

    /// <inheritdoc />
    public NamedRangeValue Read(IExcelBatch batch, string paramName)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? nameObj = null;
            dynamic? refersToRange = null;
            try
            {
                nameObj = ComUtilities.FindName(ctx.Book, paramName);
                if (nameObj == null)
                {
                    throw new InvalidOperationException($"Parameter '{paramName}' not found");
                }

                string refersTo = nameObj.RefersTo ?? "";
                refersToRange = nameObj.RefersToRange;
                object? value = refersToRange?.Value2;
                string valueType = value?.GetType().Name ?? "null";

                return new NamedRangeValue
                {
                    Name = paramName,
                    RefersTo = refersTo,
                    Value = value,
                    ValueType = valueType
                };
            }
            finally
            {
                ComUtilities.Release(ref refersToRange);
                ComUtilities.Release(ref nameObj);
            }
        });
    }

    /// <inheritdoc />
    public void Create(IExcelBatch batch, string paramName, string reference)
    {
        // Validate parameter name length (Excel limit: 255 characters)
        if (string.IsNullOrWhiteSpace(paramName))
        {
            throw new ArgumentException("Parameter name cannot be empty or whitespace", nameof(paramName));
        }

        if (paramName.Length > 255)
        {
            throw new ArgumentException($"Parameter name exceeds Excel's 255-character limit (current length: {paramName.Length})", nameof(paramName));
        }

        batch.Execute((ctx, ct) =>
        {
            dynamic? existing = null;
            dynamic? namesCollection = null;
            try
            {
                // Check if parameter already exists
                existing = ComUtilities.FindName(ctx.Book, paramName);
                if (existing != null)
                {
                    throw new InvalidOperationException($"Parameter '{paramName}' already exists");
                }

                // Create new named range
                namesCollection = ctx.Book.Names;
                // Remove any existing = prefix to avoid double ==
                string formattedReference = reference.TrimStart('=');
                // Add exactly one = prefix (required by Excel COM API)
                formattedReference = $"={formattedReference}";
                namesCollection.Add(paramName, formattedReference);

                return 0; // Dummy return for batch.Execute
            }
            finally
            {
                ComUtilities.Release(ref namesCollection);
                ComUtilities.Release(ref existing);
            }
        });
    }

    /// <inheritdoc />
    public void Update(IExcelBatch batch, string paramName, string reference)
    {
        // Validate parameter name length (Excel limit: 255 characters)
        if (string.IsNullOrWhiteSpace(paramName))
        {
            throw new ArgumentException("Parameter name cannot be empty or whitespace", nameof(paramName));
        }

        if (paramName.Length > 255)
        {
            throw new ArgumentException($"Parameter name exceeds Excel's 255-character limit (current length: {paramName.Length})", nameof(paramName));
        }

        batch.Execute((ctx, ct) =>
        {
            dynamic? nameObj = null;
            try
            {
                nameObj = ComUtilities.FindName(ctx.Book, paramName);
                if (nameObj == null)
                {
                    throw new InvalidOperationException($"Parameter '{paramName}' not found");
                }

                // Remove any existing = prefix to avoid double ==
                string formattedReference = reference.TrimStart('=');
                // Add exactly one = prefix (required by Excel COM API)
                formattedReference = $"={formattedReference}";

                // Update the reference
                nameObj.RefersTo = formattedReference;

                return 0; // Dummy return for batch.Execute
            }
            finally
            {
                ComUtilities.Release(ref nameObj);
            }
        });
    }

    /// <inheritdoc />
    public void Delete(IExcelBatch batch, string paramName)
    {
        batch.Execute((ctx, ct) =>
        {
            dynamic? nameObj = null;
            try
            {
                nameObj = ComUtilities.FindName(ctx.Book, paramName);
                if (nameObj == null)
                {
                    throw new InvalidOperationException($"Parameter '{paramName}' not found");
                }

                nameObj.Delete();
                return 0; // Dummy return for batch.Execute
            }
            finally
            {
                ComUtilities.Release(ref nameObj);
            }
        });
    }
}

