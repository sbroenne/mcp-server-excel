using System.Text.Json.Nodes;
using System.Diagnostics;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal static class MacNativeRange
{
    internal static MacRangeGeometry Describe(string filePath, string sheetName, string address, bool forWrite, TimeSpan timeout) =>
        WithRange(filePath, sheetName, address, timeout, (sheet, range, started) =>
        {
            TimeSpan Remaining() => MacAppleEvents.Remaining(timeout, started);
            var geometry = new MacRangeGeometry(
                true, Address(range, false, Remaining()),
                Read(range, MacExcelDictionary.FirstRowIndex, Remaining())!.GetValue<int>(),
                Read(range, MacExcelDictionary.FirstColumnIndex, Remaining())!.GetValue<int>(),
                Count(range, MacExcelDictionary.RowClass, Remaining()),
                Count(range, MacExcelDictionary.ColumnClass, Remaining()), false, []);
            if (!forWrite)
                return geometry;
            var mergeState = Read(range, MacExcelDictionary.MergeCells, Remaining())?.GetValue<bool>();
            // Excel reports false for a mixed native range, so only a single unmerged
            // cell can safely bypass the per-cell scan on this backend.
            if (geometry.Rows == 1 && geometry.Columns == 1 && mergeState == false)
                return geometry;
            using var first = Resolve(sheet, $"${RangeCommandValidation.ColumnLetter(geometry.Column)}${geometry.Row}");
            if (Read(first, MacExcelDictionary.MergeCells, Remaining())?.GetValue<bool>() == true)
            {
                using var merge = MacAppleEvents.Property(first, MacExcelDictionary.MergeArea);
                var mergeAddress = Address(merge, false, Remaining());
                if (geometry.Rows == 1 && geometry.Columns == 1)
                    return geometry with
                    {
                        IsMergedTopLeft = Read(merge, MacExcelDictionary.FirstRowIndex, Remaining())!.GetValue<int>() == geometry.Row
                            && Read(merge, MacExcelDictionary.FirstColumnIndex, Remaining())!.GetValue<int>() == geometry.Column,
                        MergedRanges = [mergeAddress]
                    };
                if (string.Equals(geometry.Address, mergeAddress, StringComparison.OrdinalIgnoreCase))
                    return geometry with { MergedRanges = [mergeAddress] };
            }
            var cells = (long)geometry.Rows * geometry.Columns;
            RangeMergeDiscovery.RequireSafeScanCount(cells, geometry.Address);
            var merges = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var row = 0; row < geometry.Rows; row++)
                for (var column = 0; column < geometry.Columns; column++)
                {
                    using var cell = Resolve(sheet,
                        $"${RangeCommandValidation.ColumnLetter(geometry.Column + column)}${geometry.Row + row}");
                    if (Read(cell, MacExcelDictionary.MergeCells, Remaining())?.GetValue<bool>() != true)
                        continue;
                    using var merge = MacAppleEvents.Property(cell, MacExcelDictionary.MergeArea);
                    var mergeAddress = Address(merge, false, Remaining());
                    if (seen.Add(mergeAddress)) merges.Add(mergeAddress);
                }
            return geometry with { MergedRanges = merges };
        });

    internal static MacRangeData ReadData(string filePath, string sheetName, string address, FormulaReferenceStyle style, TimeSpan timeout) =>
        WithRange(filePath, sheetName, address, timeout, (sheet, range, started) =>
        {
            TimeSpan Remaining() => MacAppleEvents.Remaining(timeout, started);
            var rows = Count(range, MacExcelDictionary.RowClass, Remaining());
            var columns = Count(range, MacExcelDictionary.ColumnClass, Remaining());
            var formulas = Matrix(Read(range, style == FormulaReferenceStyle.R1C1
                ? MacExcelDictionary.Formula2R1C1 : MacExcelDictionary.Formula2, Remaining()), rows, columns);
            var external = Address(range, true, Remaining());
            var blanks = Matrix(Evaluate($"ISBLANK({external})", Remaining()), rows, columns, allowFlatRow: true);
            var errors = Matrix(Evaluate($"IF(ISERROR({external}),ERROR.TYPE({external}),0)", Remaining()), rows, columns, allowFlatRow: true);
            var hasErrors = errors.Any(row => row!.AsArray().Any(cell => cell!.GetValue<double>() != 0));
            JsonArray values;
            if (!hasErrors)
                values = Matrix(Read(range, MacExcelDictionary.Value2, Remaining()), rows, columns);
            else
            {
                // Excel's native bulk Value2 getter can terminate Excel when its matrix contains an error.
                // Inspect first, then read only contiguous non-error runs; never retry the unsafe getter.
                values = new JsonArray();
                var firstRow = Read(range, MacExcelDictionary.FirstRowIndex, Remaining())!.GetValue<int>();
                var firstColumn = Read(range, MacExcelDictionary.FirstColumnIndex, Remaining())!.GetValue<int>();
                for (var row = 0; row < rows; row++)
                {
                    var cells = new JsonArray();
                    for (var column = 0; column < columns; column++) cells.Add((JsonNode?)null);
                    values.Add(cells);
                    for (var column = 0; column < columns;)
                    {
                        var errorType = errors[row]![column]!.GetValue<double>();
                        if (errorType != 0)
                        {
                            cells[column++] = NativeErrorCode(errorType);
                            continue;
                        }
                        var start = column;
                        while (column < columns && errors[row]![column]!.GetValue<double>() == 0) column++;
                        var runAddress = $"${RangeCommandValidation.ColumnLetter(firstColumn + start)}${firstRow + row}:" +
                            $"${RangeCommandValidation.ColumnLetter(firstColumn + column - 1)}${firstRow + row}";
                        using var run = Resolve(sheet, runAddress);
                        var runValues = Matrix(Read(run, MacExcelDictionary.Value2, Remaining()), 1, column - start);
                        for (var index = start; index < column; index++)
                            cells[index] = runValues[0]![index - start]?.DeepClone();
                    }
                }
            }
            for (var row = 0; row < rows; row++)
                for (var column = 0; column < columns; column++)
                {
                    if (blanks[row]![column]!.GetValue<bool>())
                        values[row]![column] = null;
                    var errorType = errors[row]![column]!.GetValue<double>();
                    if (errorType != 0)
                        values[row]![column] = NativeErrorCode(errorType);
                }
            var firstRowIndex = Read(range, MacExcelDictionary.FirstRowIndex, Remaining())!.GetValue<int>();
            var firstColumnIndex = Read(range, MacExcelDictionary.FirstColumnIndex, Remaining())!.GetValue<int>();
            for (var row = 0; row < rows; row++)
                for (var column = 0; column < columns; column++)
                {
                    if (formulas[row]![column] is not JsonValue formulaNode
                        || !formulaNode.TryGetValue<string>(out var formula)
                        || !formula.StartsWith('=')
                        || values[row]![column] is not JsonValue valueNode
                        || !valueNode.TryGetValue<string>(out var text)
                        || !string.Equals(text, formula, StringComparison.Ordinal))
                        continue;
                    using var cell = Resolve(sheet,
                        $"${RangeCommandValidation.ColumnLetter(firstColumnIndex + column)}${firstRowIndex + row}");
                    if (Read(cell, MacExcelDictionary.HasFormula, Remaining())?.GetValue<bool>() is not { } hasFormula)
                        throw new InvalidDataException("Excel did not return formula state for an ambiguous text value.");
                    if (!hasFormula)
                        formulas[row]![column] = JsonValue.Create(string.Empty);
                }
            return new MacRangeData(true, rows, columns, formulas, values);
        });

    internal static OperationResult SetFormulas(string filePath, string sheetName, string address,
        List<List<string>> formulas, FormulaReferenceStyle style, TimeSpan timeout) =>
        WithRange(filePath, sheetName, address, timeout, (_, range, started) =>
        {
            using var matrix = MacAppleEvents.List();
            foreach (var formulasRow in formulas)
            {
                using var row = MacAppleEvents.List();
                foreach (var formula in formulasRow)
                {
                    using var value = MacAppleEvents.Text(formula);
                    MacAppleEvents.Append(row, value);
                }
                MacAppleEvents.Append(matrix, row);
            }
            // One native matrix write preserves calculation mode without mutating shared application settings.
            Write(range, style == FormulaReferenceStyle.R1C1 ? MacExcelDictionary.Formula2R1C1 : MacExcelDictionary.Formula2,
                matrix, MacAppleEvents.Remaining(timeout, started));
            return new OperationResult { Success = true, FilePath = filePath, Action = "set-formulas" };
        });

    private static T WithRange<T>(string filePath, string sheetName, string address, TimeSpan timeout,
        Func<MacAppleEvents.Descriptor, MacAppleEvents.Descriptor, long, T> action)
    {
        var started = Stopwatch.GetTimestamp();
        using var workbook = MacNativeWorkbook.Resolve(filePath, timeout);
        using var name = MacAppleEvents.Text(sheetName);
        using var sheet = MacAppleEvents.Object(MacExcelDictionary.WorksheetClass, workbook, MacAppleEvents.Code("name"), name);
        using var range = Resolve(sheet, address);
        return action(sheet, range, started);
    }

    internal static JsonArray Matrix(JsonNode? value, int rows, int columns, bool allowFlatRow = false)
    {
        if (rows <= 0 || columns <= 0)
            throw new InvalidDataException("Excel returned invalid range dimensions.");
        if (value is not JsonArray)
        {
            if (rows != 1 || columns != 1)
                throw new InvalidDataException("Excel returned an unexpected range shape.");
            return new JsonArray(new JsonArray(value?.DeepClone()));
        }
        var matrix = value.AsArray();
        // Evaluate flattens a horizontal vector; range property getters retain both dimensions.
        if (allowFlatRow && rows == 1 && matrix.Count == columns
            && matrix.All(cell => cell is null or JsonValue))
            return new JsonArray(matrix.DeepClone());
        if (matrix.Count != rows || matrix.Any(row => row is not JsonArray cells || cells.Count != columns))
            throw new InvalidDataException("Excel returned an unexpected range shape.");
        return (JsonArray)matrix.DeepClone();
    }

    internal static object? Scalar(JsonNode? node)
    {
        if (node is null) return null;
        if (node is not JsonValue value) throw new InvalidDataException("Excel returned a nonscalar cell.");
        if (value.TryGetValue<string>(out var text)) return text;
        if (value.TryGetValue<bool>(out var boolean)) return boolean;
        if (value.TryGetValue<int>(out var integer)) return integer;
        if (value.TryGetValue<double>(out var number)) return number;
        throw new InvalidDataException("Excel returned an unsupported cell value.");
    }

    private static int NativeErrorCode(double errorType) => errorType switch
    {
        1 => -2146826288,
        2 => -2146826281,
        3 => -2146826273,
        4 => -2146826265,
        5 => -2146826259,
        6 => -2146826252,
        7 => -2146826246,
        8 => -2146826245,
        _ => throw new PlatformNotSupportedException(
            $"Native macOS formula reads do not yet support Excel ERROR.TYPE {errorType}. No result was approximated.")
    };

    internal static MacAppleEvents.Descriptor Resolve(MacAppleEvents.Descriptor sheet, string address)
    {
        using var name = MacAppleEvents.Text(address);
        return MacAppleEvents.Object(MacExcelDictionary.RangeClass, sheet, MacAppleEvents.Code("name"), name);
    }

    internal static JsonNode? Read(MacAppleEvents.Descriptor range, uint property, TimeSpan timeout)
    {
        using var target = MacAppleEvents.Property(range, property);
        using var command = MacAppleEvents.Event(MacAppleEvents.Code("core"), MacAppleEvents.Code("getd"));
        MacAppleEvents.Put(command, MacAppleEvents.Code("----"), target);
        return MacAppleEvents.Send(command, timeout);
    }

    internal static void Write(MacAppleEvents.Descriptor range, uint property, MacAppleEvents.Descriptor value, TimeSpan timeout)
    {
        using var target = MacAppleEvents.Property(range, property);
        using var command = MacAppleEvents.Event(MacAppleEvents.Code("core"), MacAppleEvents.Code("setd"));
        MacAppleEvents.Put(command, MacAppleEvents.Code("----"), target);
        MacAppleEvents.Put(command, MacAppleEvents.Code("data"), value);
        MacAppleEvents.SendCommand(command, timeout);
    }

    internal static string Address(MacAppleEvents.Descriptor range, bool external, TimeSpan timeout)
    {
        using var flag = MacAppleEvents.Create(MacAppleEvents.Code("bool"), [external ? (byte)1 : (byte)0]);
        using var command = MacAppleEvents.Event(MacExcelDictionary.GetAddressClass, MacExcelDictionary.GetAddressId);
        MacAppleEvents.Put(command, MacAppleEvents.Code("----"), range);
        MacAppleEvents.Put(command, MacExcelDictionary.AddressExternalParameter, flag);
        return MacAppleEvents.Send(command, timeout)?.GetValue<string>()
            ?? throw new InvalidDataException("Excel did not return a native range address.");
    }

    internal static int Count(MacAppleEvents.Descriptor range, uint itemClass, TimeSpan timeout)
    {
        using var type = MacAppleEvents.Create(MacAppleEvents.Code("type"), BitConverter.GetBytes(itemClass));
        using var command = MacAppleEvents.Event(MacExcelDictionary.CountClass, MacExcelDictionary.CountId);
        MacAppleEvents.Put(command, MacAppleEvents.Code("----"), range);
        MacAppleEvents.Put(command, MacExcelDictionary.CountClassParameter, type);
        return MacAppleEvents.Send(command, timeout)?.GetValue<int>()
            ?? throw new InvalidDataException("Excel did not return a native range dimension.");
    }

    internal static JsonNode? Evaluate(string expression, TimeSpan timeout)
    {
        using var name = MacAppleEvents.Text(expression);
        using var command = MacAppleEvents.Event(MacExcelDictionary.EvaluateClass, MacExcelDictionary.EvaluateId);
        MacAppleEvents.Put(command, MacExcelDictionary.EvaluateNameParameter, name);
        return MacAppleEvents.Send(command, timeout);
    }
}

internal sealed record MacRangeGeometry(bool Success, string Address, int Row, int Column,
    int Rows, int Columns, bool IsMergedTopLeft, List<string> MergedRanges);

internal sealed record MacRangeData(bool Success, int Rows, int Columns, JsonArray Formulas, JsonArray Values);
