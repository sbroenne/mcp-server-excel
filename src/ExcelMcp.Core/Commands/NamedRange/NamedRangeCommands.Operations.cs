using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Named range lifecycle operations (List, Read, Write, Create, Update, Delete)
/// </summary>
public partial class NamedRangeCommands
{
    private const long MaxListValuePreviewCellCount = 10_000;

    private static string ToR1C1Formula(Excel.Application application, string formula)
    {
        Excel.Range? activeCell = null;
        try
        {
            activeCell = (Excel.Range)application.ActiveCell;
            object? converted = application.ConvertFormula(
                formula,
                Excel.XlReferenceStyle.xlA1,
                Excel.XlReferenceStyle.xlR1C1,
                Type.Missing,
                activeCell);
            if (converted is string convertedFormula)
                return convertedFormula;

            // Excel cannot convert structured table references, which already use the same syntax in R1C1.
            if (formula.Contains('[') && !formula.Contains("]!", StringComparison.Ordinal))
                return formula;

            throw new InvalidOperationException("Excel could not convert the named formula to R1C1 notation.");
        }
        finally
        {
            ComUtilities.Release(ref activeCell);
        }
    }

    /// <inheritdoc />
    public NamedRangeListResult List(IExcelBatch batch)
    {
        var result = new NamedRangeListResult
        {
            FilePath = batch.WorkbookPath
        };

        return batch.Execute(ListNames);

        [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
            Justification = "RefersToRange can reject formula/external names; reports that specific getter failure per-name as Unavailable with native HRESULT. Other metadata and preview failures propagate.")]
        NamedRangeListResult ListNames(ExcelContext ctx, CancellationToken ct)
        {
            Excel.Names? namesCollection = null;
            try
            {
                namesCollection = ctx.Book.Names;
                int count = namesCollection.Count;

                for (int i = 1; i <= count; i++)
                {
                    ct.ThrowIfCancellationRequested();

                    Excel.Name? nameObj = null;
                    Excel.Range? refersToRange = null;
                    try
                    {
                        nameObj = namesCollection.Item(i);
                        string name = nameObj.Name?.ToString() ?? string.Empty;

                        if (ShouldSkipNameFromList(nameObj, name))
                        {
                            continue;
                        }

                        string refersTo = ((object?)nameObj.RefersTo)?.ToString() ?? string.Empty;

                        var info = new NamedRangeInfo
                        {
                            Name = name,
                            RefersTo = refersTo,
                            ValueType = "null"
                        };

                        try
                        {
                            refersToRange = nameObj.RefersToRange;
                        }
                        catch (COMException ex) when (ex.HResult == unchecked((int)0x800A03EC) || ComErrors.IsMemberUnavailable(ex))
                        {
                            // Named range may not have a valid RefersToRange (e.g., formula-based or external reference)
                            // Continue with metadata only - this is expected for some named ranges.
                            info.ValueType = "Unavailable";
                            info.ValueOmittedReason = $"{ex.Message} (HRESULT 0x{ex.HResult:X8}).";
                        }
                        if (refersToRange is not null)
                            PopulateListValuePreview(refersToRange, info);

                        result.NamedRanges.Add(info);
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
        }
    }

    private static bool ShouldSkipNameFromList(Excel.Name nameObj, string name)
    {
        if (IsHiddenName(nameObj))
        {
            return true;
        }

        return IsBuiltInName(name);
    }

    private static bool IsHiddenName(Excel.Name nameObj) => !nameObj.Visible;

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

    private static Excel.Name? FindNameByKey(Excel.Workbook workbook, string name)
    {
        Excel.Names? names = null;
        Excel.Name? nameObj = null;
        try
        {
            names = workbook.Names;
            int count = names.Count;
            for (int index = 1; index <= count; index++)
            {
                nameObj = names.Item(index);
                if (string.Equals(nameObj.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    var result = nameObj;
                    nameObj = null;
                    return result;
                }
                ComUtilities.Release(ref nameObj);
            }
            return null;
        }
        finally
        {
            ComUtilities.Release(ref nameObj);
            ComUtilities.Release(ref names);
        }
    }

    private static void PopulateListValuePreview(Excel.Range refersToRange, NamedRangeInfo info)
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

        object? rawValue = refersToRange.Value2;

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

    private static int GetAreaCount(Excel.Range range)
    {
        Excel.Areas? areas = null;
        try
        {
            areas = range.Areas;
            return areas.Count;
        }
        finally
        {
            ComUtilities.Release(ref areas);
        }
    }

    private static long GetCellCount(Excel.Range range)
    {
        return Convert.ToInt64((object)range.CountLarge, CultureInfo.InvariantCulture);
    }

    internal static object ParseWriteValue(string value)
    {
        if (double.TryParse(
            value,
            NumberStyles.Float,
            CultureInfo.InvariantCulture,
            out double numericValue))
        {
            return numericValue;
        }

        return bool.TryParse(value, out bool booleanValue)
            ? booleanValue
            : value;
    }

    /// <inheritdoc />
    public OperationResult Write(IExcelBatch batch, string name, string value)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Name? nameObj = null;
            Excel.Range? refersToRange = null;
            int originalCalculation = -1;
            bool calculationChanged = false;
            Exception? operationFailure = null;

            try
            {
                nameObj = FindNameByKey(ctx.Book, name);
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

                refersToRange.Value2 = ParseWriteValue(value);

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath }; // Dummy return for batch.Execute
            }
            catch (Exception ex)
            {
                operationFailure = ex;
                throw;
            }
            finally
            {
                try
                {
                    if (calculationChanged && originalCalculation != -1)
                    {
                        OperationCleanup.Run(() => ctx.App.Calculation = (Excel.XlCalculation)originalCalculation,
                            operationFailure);
                    }
                }
                finally
                {
                    ComUtilities.Release(ref refersToRange);
                    ComUtilities.Release(ref nameObj);
                }
            }
        });
    }

    /// <inheritdoc />
    public NamedRangeValue Read(IExcelBatch batch, string name)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Name? nameObj = null;
            Excel.Range? refersToRange = null;
            try
            {
                nameObj = FindNameByKey(ctx.Book, name);
                if (nameObj == null)
                {
                    throw new InvalidOperationException($"Named range '{name}' not found.");
                }

                string refersTo = ((object?)nameObj.RefersTo)?.ToString() ?? "";
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
            Excel.Name? created = null;
            Excel.Names? namesCollection = null;
            try
            {
                // Check if parameter already exists
                existing = FindNameByKey(ctx.Book, name);
                if (existing != null)
                {
                    throw new InvalidOperationException($"Named range '{name}' already exists");
                }

                namesCollection = ctx.Book.Names;
                string formula = $"={reference.TrimStart('=')}";
                string r1c1Formula = ToR1C1Formula(ctx.App, formula);
                created = namesCollection.Add(Name: name, RefersToR1C1: r1c1Formula);
                if (created is null)
                    throw new InvalidOperationException(
                        $"Excel did not confirm creating named range '{name}'. Inspect workbook names before retrying.");

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath }; // Dummy return for batch.Execute
            }
            finally
            {
                ComUtilities.Release(ref created);
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
                nameObj = FindNameByKey(ctx.Book, name);
                if (nameObj == null)
                {
                    throw new InvalidOperationException($"Named range '{name}' not found.");
                }

                string formula = $"={reference.TrimStart('=')}";
                nameObj.RefersToR1C1 = ToR1C1Formula(ctx.App, formula);

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
                nameObj = FindNameByKey(ctx.Book, name);
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
