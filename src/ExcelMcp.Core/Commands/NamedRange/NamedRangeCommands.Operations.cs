using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Xml.Linq;
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
    private const long ExcelMaxRows = 1_048_576;
    private const long ExcelMaxColumns = 16_384;

    private sealed record WorkbookDefinedName(string Name, string RefersTo);

    /// <inheritdoc />
    public NamedRangeListResult List(IExcelBatch batch)
    {
        var result = new NamedRangeListResult
        {
            FilePath = batch.WorkbookPath
        };

        var packageDefinedNames = TryReadVisibleDefinedNamesFromPackage(batch.WorkbookPath, out var hasHiddenOrInternalNames);
        if (hasHiddenOrInternalNames && packageDefinedNames != null)
        {
            return ListFromWorkbookPackage(batch, result, packageDefinedNames);
        }

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

    private static NamedRangeListResult ListFromWorkbookPackage(
        IExcelBatch batch,
        NamedRangeListResult result,
        List<WorkbookDefinedName> packageDefinedNames)
    {
        return batch.Execute((ctx, ct) =>
        {
            foreach (var definedName in packageDefinedNames)
            {
                ct.ThrowIfCancellationRequested();

                var info = new NamedRangeInfo
                {
                    Name = definedName.Name,
                    RefersTo = definedName.RefersTo,
                    ValueType = "Unavailable",
                    ValueOmittedReason = "Value preview unavailable from workbook defined-name metadata."
                };

                PopulateListValuePreviewFromReference(ctx.Book, definedName.RefersTo, info);
                result.NamedRanges.Add(info);
            }

            result.Success = true;
            return result;
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

    private static List<WorkbookDefinedName>? TryReadVisibleDefinedNamesFromPackage(
        string workbookPath,
        out bool hasHiddenOrInternalNames)
    {
        hasHiddenOrInternalNames = false;

        if (string.IsNullOrWhiteSpace(workbookPath) || !File.Exists(workbookPath))
        {
            return null;
        }

        try
        {
            using var stream = new FileStream(
                workbookPath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
            var workbookEntry = archive.GetEntry("xl/workbook.xml");
            if (workbookEntry == null)
            {
                return null;
            }

            using var workbookStream = workbookEntry.Open();
            var document = XDocument.Load(workbookStream);
            var definedNames = document
                .Descendants()
                .Where(element => element.Name.LocalName == "definedName");

            var visibleNames = new List<WorkbookDefinedName>();
            foreach (var definedName in definedNames)
            {
                var name = definedName.Attribute("name")?.Value ?? string.Empty;
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                if (IsTruthyOpenXmlBoolean(definedName.Attribute("hidden")?.Value) || IsBuiltInName(name))
                {
                    hasHiddenOrInternalNames = true;
                    continue;
                }

                visibleNames.Add(new WorkbookDefinedName(name, definedName.Value));
            }

            return visibleNames;
        }
        catch (InvalidDataException)
        {
            return null;
        }
        catch (IOException)
        {
            return null;
        }
        catch (UnauthorizedAccessException)
        {
            return null;
        }
        catch (System.Xml.XmlException)
        {
            return null;
        }
    }

    private static bool IsTruthyOpenXmlBoolean(string? value)
    {
        return value != null
            && (value.Equals("1", StringComparison.OrdinalIgnoreCase)
                || value.Equals("true", StringComparison.OrdinalIgnoreCase));
    }

    private static void PopulateListValuePreviewFromReference(dynamic workbook, string refersTo, NamedRangeInfo info)
    {
        if (!TrySplitRangeReference(refersTo, out var sheetName, out var rangeAddress))
        {
            return;
        }

        if (rangeAddress.Contains(',', StringComparison.Ordinal))
        {
            info.ValueType = "MultiAreaRange";
            info.ValueOmittedReason = "Named range resolves to multiple areas; list omits multi-area value previews.";
            return;
        }

        if (TryGetCellCountFromAddress(rangeAddress, out var cellCount))
        {
            info.CellCount = cellCount;
            if (cellCount > MaxListValuePreviewCellCount)
            {
                info.ValueType = "RangeTooLarge";
                info.ValueOmittedReason =
                    $"Named range contains {cellCount} cells, which exceeds the list preview limit of {MaxListValuePreviewCellCount}.";
                return;
            }
        }

        dynamic? sheet = null;
        dynamic? range = null;
        try
        {
            sheet = ComUtilities.FindSheet(workbook, sheetName);
            if (sheet == null)
            {
                return;
            }

            range = sheet.Range[rangeAddress];
            PopulateListValuePreview(range, info);
        }
        catch (Exception ex) when (IsRecoverableNamedRangeException(ex))
        {
            info.ValueType = "Unavailable";
            info.ValueOmittedReason = ex.Message;
        }
        finally
        {
            ComUtilities.Release(ref range);
            ComUtilities.Release(ref sheet);
        }
    }

    private static bool TrySplitRangeReference(string refersTo, out string sheetName, out string rangeAddress)
    {
        sheetName = string.Empty;
        rangeAddress = string.Empty;

        var reference = refersTo.Trim();
        if (reference.StartsWith('='))
        {
            reference = reference[1..];
        }

        if (reference.Contains('[', StringComparison.Ordinal))
        {
            return false;
        }

        var bangIndex = reference.LastIndexOf('!');
        if (bangIndex <= 0 || bangIndex == reference.Length - 1)
        {
            return false;
        }

        sheetName = reference[..bangIndex].Trim();
        if (sheetName.Length >= 2 && sheetName[0] == '\'' && sheetName[^1] == '\'')
        {
            sheetName = sheetName[1..^1].Replace("''", "'", StringComparison.Ordinal);
        }

        rangeAddress = reference[(bangIndex + 1)..].Trim();
        return sheetName.Length > 0 && rangeAddress.Length > 0;
    }

    private static bool TryGetCellCountFromAddress(string rangeAddress, out long cellCount)
    {
        cellCount = 0;
        var address = rangeAddress.Replace("$", string.Empty, StringComparison.Ordinal).Trim();

        var cellMatch = Regex.Match(
            address,
            @"^(?<col1>[A-Za-z]{1,3})(?<row1>\d+)(:(?<col2>[A-Za-z]{1,3})(?<row2>\d+))?$",
            RegexOptions.CultureInvariant);
        if (cellMatch.Success)
        {
            var col1 = ColumnNameToNumber(cellMatch.Groups["col1"].Value);
            var row1 = long.Parse(cellMatch.Groups["row1"].Value, System.Globalization.CultureInfo.InvariantCulture);
            var col2 = cellMatch.Groups["col2"].Success ? ColumnNameToNumber(cellMatch.Groups["col2"].Value) : col1;
            var row2 = cellMatch.Groups["row2"].Success
                ? long.Parse(cellMatch.Groups["row2"].Value, System.Globalization.CultureInfo.InvariantCulture)
                : row1;

            cellCount = (Math.Abs(row2 - row1) + 1) * (Math.Abs(col2 - col1) + 1);
            return true;
        }

        var columnMatch = Regex.Match(
            address,
            @"^(?<col1>[A-Za-z]{1,3}):(?<col2>[A-Za-z]{1,3})$",
            RegexOptions.CultureInvariant);
        if (columnMatch.Success)
        {
            var col1 = ColumnNameToNumber(columnMatch.Groups["col1"].Value);
            var col2 = ColumnNameToNumber(columnMatch.Groups["col2"].Value);
            cellCount = (Math.Abs(col2 - col1) + 1) * ExcelMaxRows;
            return true;
        }

        var rowMatch = Regex.Match(
            address,
            @"^(?<row1>\d+):(?<row2>\d+)$",
            RegexOptions.CultureInvariant);
        if (rowMatch.Success)
        {
            var row1 = long.Parse(rowMatch.Groups["row1"].Value, System.Globalization.CultureInfo.InvariantCulture);
            var row2 = long.Parse(rowMatch.Groups["row2"].Value, System.Globalization.CultureInfo.InvariantCulture);
            cellCount = (Math.Abs(row2 - row1) + 1) * ExcelMaxColumns;
            return true;
        }

        return false;
    }

    private static long ColumnNameToNumber(string columnName)
    {
        long number = 0;
        foreach (var character in columnName.ToUpperInvariant())
        {
            number = (number * 26) + character - 'A' + 1;
        }

        return number;
    }

    private static Excel.Name? FindNameByKey(dynamic workbook, string name)
    {
        dynamic? names = null;
        Excel.Name? nameObj = null;
        try
        {
            names = workbook.Names;
            nameObj = (Excel.Name)names.Item(name);
            var result = nameObj;
            nameObj = null;
            return result;
        }
        catch (Exception ex) when (IsRecoverableNamedRangeException(ex))
        {
            return null;
        }
        finally
        {
            ComUtilities.Release(ref nameObj);
            ComUtilities.Release(ref names);
        }
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
                nameObj = FindNameByKey(ctx.Book, name);
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
                existing = FindNameByKey(ctx.Book, name);
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
                nameObj = FindNameByKey(ctx.Book, name);
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



