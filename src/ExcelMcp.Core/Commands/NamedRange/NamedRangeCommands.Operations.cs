using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Named range lifecycle operations (List, Read, Write, Create, Update, Delete)
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
    public OperationResult Write(IExcelBatch batch, string name, string value)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Name? nameObj = null;
            dynamic? refersToRange = null;
            int originalCalculation = -1;// xlCalculationAutomatic = -4105, xlCalculationManual = -4135
            bool calculationChanged = false;

            try
            {
                nameObj = ComUtilities.FindName(ctx.Book, name);
                if (nameObj == null)
                {
                    throw new InvalidOperationException($"Named range '{name}' not found.");
                }

                refersToRange = nameObj.RefersToRange;

                // CRITICAL: Temporarily disable automatic calculation to prevent Excel from
                // hanging when changed parameter values trigger dependent formulas that reference Data Model/DAX.
                // Without this, setting values can block the COM interface during recalculation.
                originalCalculation = (int)ctx.App.Calculation;
                if (originalCalculation != -4135) // xlCalculationManual
                {
                    ctx.App.Calculation = (Excel.XlCalculation)(-4135); // xlCalculationManual
                    calculationChanged = true;
                }

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

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath }; // Dummy return for batch.Execute
            }
            finally
            {
                // Restore original calculation mode
                if (calculationChanged && originalCalculation != -1)
                {
                    try
                    {
                        ctx.App.Calculation = (Excel.XlCalculation)originalCalculation;
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        // Ignore errors restoring calculation mode - not critical
                    }
                }
                ComUtilities.Release(ref refersToRange);
                ComUtilities.Release(ref nameObj);
            }
        });
    }

    /// <inheritdoc />
    public NamedRangeValue Read(IExcelBatch batch, string name)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Name? nameObj = null;
            dynamic? refersToRange = null;
            try
            {
                nameObj = ComUtilities.FindName(ctx.Book, name);
                if (nameObj == null)
                {
                    throw new InvalidOperationException($"Named range '{name}' not found.");
                }

                string refersTo = nameObj.RefersTo?.ToString() ?? "";
                refersToRange = nameObj.RefersToRange;
                object? value = refersToRange?.Value2;
                string valueType = value?.GetType().Name ?? "null";

                return new NamedRangeValue
                {
                    Name = name,
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
    public OperationResult Create(IExcelBatch batch, string name, string reference)
    {
        // Validate parameter name length (Excel limit: 255 characters)
        if (string.IsNullOrWhiteSpace(name))
        {
            throw new ArgumentException("Named range name cannot be empty or whitespace", nameof(name));
        }

        if (name.Length > 255)
        {
            throw new ArgumentException($"Named range name exceeds Excel's 255-character limit (current length: {name.Length})", nameof(name));
        }

        return batch.Execute((ctx, ct) =>
        {
            Excel.Name? existing = null;
            dynamic? namesCollection = null;
            try
            {
                // Check if parameter already exists
                existing = ComUtilities.FindName(ctx.Book, name);
                if (existing != null)
                {
                    throw new InvalidOperationException($"Named range '{name}' already exists");
                }

                // Create new named range
                namesCollection = ctx.Book.Names;
                // Remove any existing = prefix to avoid double ==
                string formattedReference = reference.TrimStart('=');
                // Add exactly one = prefix (required by Excel COM API)
                formattedReference = $"={formattedReference}";
                namesCollection.Add(name, formattedReference);

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath }; // Dummy return for batch.Execute
            }
            finally
            {
                ComUtilities.Release(ref namesCollection);
                ComUtilities.Release(ref existing);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Update(IExcelBatch batch, string name, string reference)
    {
        // Validate parameter name length (Excel limit: 255 characters)
        if (string.IsNullOrWhiteSpace(name))
        {
            throw new ArgumentException("Named range name cannot be empty or whitespace", nameof(name));
        }

        if (name.Length > 255)
        {
            throw new ArgumentException($"Named range name exceeds Excel's 255-character limit (current length: {name.Length})", nameof(name));
        }

        return batch.Execute((ctx, ct) =>
        {
            Excel.Name? nameObj = null;
            try
            {
                nameObj = ComUtilities.FindName(ctx.Book, name);
                if (nameObj == null)
                {
                    throw new InvalidOperationException($"Named range '{name}' not found.");
                }

                // Remove any existing = prefix to avoid double ==
                string formattedReference = reference.TrimStart('=');
                // Add exactly one = prefix (required by Excel COM API)
                formattedReference = $"={formattedReference}";

                // Update the reference
                nameObj.RefersTo = formattedReference;

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath }; // Dummy return for batch.Execute
            }
            finally
            {
                ComUtilities.Release(ref nameObj);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Delete(IExcelBatch batch, string name)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Name? nameObj = null;
            try
            {
                nameObj = ComUtilities.FindName(ctx.Book, name);
                if (nameObj == null)
                {
                    throw new InvalidOperationException($"Named range '{name}' not found.");
                }

                nameObj.Delete();
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath }; // Dummy return for batch.Execute
            }
            finally
            {
                ComUtilities.Release(ref nameObj);
            }
        });
    }
}



