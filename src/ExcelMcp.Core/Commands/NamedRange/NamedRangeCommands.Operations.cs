using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;
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
    private const long MaxListValuePreviewCellCount = 10_000;

    /// <inheritdoc />
    public NamedRangeListResult List(IExcelBatch batch)
    {
        var result = new NamedRangeListResult
        {
            FilePath = batch.WorkbookPath
        };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? namesCollection = null;
            try
            {
                namesCollection = ctx.Book.Names;
                int count = Convert.ToInt32(namesCollection.Count);

                for (int i = 1; i <= count; i++)
                {
                    ct.ThrowIfCancellationRequested();

                    dynamic? nameObj = null;
                    dynamic? refersToRange = null;
                    try
                    {
                        nameObj = namesCollection.Item(i);
                        string name = nameObj.Name?.ToString() ?? string.Empty;

                        if (ShouldSkipNameFromList(nameObj, name))
                        {
                            continue;
                        }

                        string refersTo = nameObj.RefersTo?.ToString() ?? string.Empty;

                        var info = new NamedRangeInfo
                        {
                            Name = name,
                            RefersTo = refersTo,
                            ValueType = "null"
                        };

                        try
                        {
                            refersToRange = nameObj.RefersToRange;
                            PopulateListValuePreview(refersToRange, info);
                        }
                        catch (Exception ex) when (IsRecoverableNamedRangeException(ex))
                        {
                            // Named range may not have a valid RefersToRange (e.g., formula-based or external reference)
                            // Continue with metadata only - this is expected for some named ranges.
                            info.ValueType = "Unavailable";
                            info.ValueOmittedReason = ex.Message;
                        }

                        result.NamedRanges.Add(info);
                    }
                    catch (Exception ex) when (IsRecoverableNamedRangeException(ex))
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

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref namesCollection);
            }
        });
    }

    private static bool ShouldSkipNameFromList(dynamic nameObj, string name)
    {
        if (IsHiddenName(nameObj))
        {
            return true;
        }

        return IsBuiltInName(name);
    }

    private static bool IsHiddenName(dynamic nameObj)
    {
        try
        {
            return !Convert.ToBoolean(nameObj.Visible);
        }
        catch (Exception ex) when (IsRecoverableNamedRangeException(ex))
        {
            return true;
        }
    }

    private static bool IsBuiltInName(string name)
    {
        var localName = GetLocalName(name);
        return localName.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase)
            || localName.Equals("_FilterDatabase", StringComparison.OrdinalIgnoreCase);
    }

    private static string GetLocalName(string name)
    {
        var bangIndex = name.LastIndexOf('!');
        return bangIndex >= 0 ? name[(bangIndex + 1)..].Trim('\'') : name;
    }

    private static void PopulateListValuePreview(dynamic refersToRange, NamedRangeInfo info)
    {
        var areaCount = GetAreaCount(refersToRange);
        if (areaCount > 1)
        {
            info.ValueType = "MultiAreaRange";
            info.ValueOmittedReason = "Named range resolves to multiple areas; list omits multi-area value previews.";
            return;
        }

        var cellCount = GetCellCount(refersToRange);
        info.CellCount = cellCount;

        if (cellCount > MaxListValuePreviewCellCount)
        {
            info.ValueType = "RangeTooLarge";
            info.ValueOmittedReason =
                $"Named range contains {cellCount} cells, which exceeds the list preview limit of {MaxListValuePreviewCellCount}.";
            return;
        }

        var rawValue = refersToRange?.Value2;

        if (rawValue is object[,] array2D)
        {
            info.Value = ConvertArrayToList(array2D);
            info.ValueType = "Array";
        }
        else
        {
            info.Value = ConvertValueForJson(rawValue);
            info.ValueType = rawValue?.GetType().Name ?? "null";
        }
    }

    private static int GetAreaCount(dynamic range)
    {
        dynamic? areas = null;
        try
        {
            areas = range.Areas;
            return Convert.ToInt32(areas.Count);
        }
        finally
        {
            ComUtilities.Release(ref areas);
        }
    }

    private static long GetCellCount(dynamic range)
    {
        return Convert.ToInt64(range.CountLarge);
    }

    private static bool IsRecoverableNamedRangeException(Exception ex) =>
        ex is COMException
            or InvalidCastException
            or RuntimeBinderException
            or SafeArrayRankMismatchException
            or SafeArrayTypeMismatchException;

    /// <inheritdoc />
    public OperationResult Write(IExcelBatch batch, string name, string value)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Name? nameObj = null;
            dynamic? refersToRange = null;
            int originalCalculation = -1;
            bool calculationChanged = false;

            try
            {
                nameObj = ComUtilities.FindName(ctx.Book, name);
                if (nameObj == null)
                {
                    throw new InvalidOperationException($"Named range '{name}' not found.");
                }

                refersToRange = nameObj.RefersToRange;

                // Calculation suppressed here (not in ExcelWriteGuard) because Data Model ops need it enabled
                originalCalculation = (int)ctx.App.Calculation;
                if (originalCalculation != -4135) // xlCalculationManual
                {
                    ctx.App.Calculation = (Excel.XlCalculation)(-4135);
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
                if (calculationChanged && originalCalculation != -1)
                {
                    try
                    {
                        ctx.App.Calculation = (Excel.XlCalculation)originalCalculation;
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        // Ignore errors restoring calculation mode
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
                object? rawValue = refersToRange?.Value2;
                object? value;
                string valueType;
                if (rawValue is object[,] array2D)
                {
                    value = ConvertArrayToList(array2D);
                    valueType = "Array";
                }
                else
                {
                    value = ConvertValueForJson(rawValue);
                    valueType = rawValue?.GetType().Name ?? "null";
                }

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



