using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Resolves chart data sources, including sources made of separate cell blocks.
/// </summary>
internal static class ChartSourceRange
{
    /// <summary>
    /// A chart source split into the sheet it reads from and its sheet-local blocks.
    /// </summary>
    internal sealed record ParsedSource(string SheetName, IReadOnlyList<string> Blocks);

    /// <summary>
    /// Resolves a chart source address to one Excel range on one worksheet.
    /// Blocks without a sheet name use the first block's sheet, or <paramref name="defaultSheetName"/>
    /// when the first block has no sheet name. Throws before any workbook change when the
    /// address is invalid, mixes sheets, or names a missing sheet.
    /// </summary>
    internal static dynamic Resolve(dynamic book, string defaultSheetName, string sourceAddress)
    {
        var parsed = Parse(sourceAddress, defaultSheetName);
        try
        {
            var sheetName = ActualSheetName(book, parsed.SheetName);
            return RangeHelpers.ResolveRange(book, sheetName, string.Join(",", parsed.Blocks))!;
        }
        catch (OperationFailureException ex)
            when (ex.ErrorCategory == OperationFailureCategory.InvalidInput && parsed.Blocks.Count == 1)
        {
            return ResolveDefinedName(book, sourceAddress, ex);
        }
        catch (OperationFailureException ex)
            when (ex.ErrorCategory == OperationFailureCategory.InvalidInput)
        {
            throw InvalidAddress(sourceAddress, ex.Message, ex);
        }
    }

    /// <summary>
    /// Applies an already-resolved source to a chart, explaining Excel's generic rejection.
    /// </summary>
    internal static void Apply(dynamic chart, dynamic sourceRange, string sourceAddress)
    {
        try
        {
            chart.SetSourceData(sourceRange);
        }
        catch (COMException ex) when (ex.HResult == unchecked((int)0x800A03EC))
        {
            throw new InvalidOperationException(
                $"Excel could not use '{sourceAddress}' as the chart data source. " +
                "Check that the cells contain chartable data and that every block has the same number of rows " +
                "(for column-based series) or columns (for row-based series).",
                ex);
        }
    }

    /// <summary>
    /// Splits a source address into sheet-local blocks on a single sheet without touching Excel.
    /// </summary>
    internal static ParsedSource Parse(string sourceAddress, string defaultSheetName)
    {
        if (string.IsNullOrWhiteSpace(sourceAddress))
        {
            throw InvalidAddress(sourceAddress, "The address is empty.");
        }

        var rawBlocks = SplitBlocks(sourceAddress);
        string? sourceSheet = null;
        var blocks = new List<string>(rawBlocks.Count);
        for (int index = 0; index < rawBlocks.Count; index++)
        {
            var (blockSheet, local) = SplitSheetPrefix(sourceAddress, rawBlocks[index]);
            if (index == 0)
            {
                sourceSheet = blockSheet ?? defaultSheetName;
            }

            var resolvedSheet = blockSheet ?? sourceSheet!;
            if (!string.Equals(resolvedSheet, sourceSheet, StringComparison.OrdinalIgnoreCase))
            {
                throw new OperationFailureException(
                    OperationFailureCategory.InvalidInput,
                    $"Chart source range '{sourceAddress}' uses blocks from sheet '{sourceSheet}' and sheet '{resolvedSheet}'. " +
                    "All blocks in a chart source must be on the same sheet.");
            }

            blocks.Add(local);
        }

        return new ParsedSource(sourceSheet!, blocks);
    }

    private static List<string> SplitBlocks(string sourceAddress)
    {
        var blocks = new List<string>();
        int start = 0;
        int bracketDepth = 0;
        bool inQuotedSheet = false;
        for (int index = 0; index < sourceAddress.Length; index++)
        {
            char character = sourceAddress[index];
            if (inQuotedSheet)
            {
                if (character == '\'')
                {
                    if (index + 1 < sourceAddress.Length && sourceAddress[index + 1] == '\'')
                    {
                        index++;
                    }
                    else
                    {
                        inQuotedSheet = false;
                    }
                }

                continue;
            }

            if (bracketDepth > 0)
            {
                if (character == '\'' && index + 1 < sourceAddress.Length)
                {
                    index++;
                }
                else if (character == '[')
                {
                    bracketDepth++;
                }
                else if (character == ']')
                {
                    bracketDepth--;
                }

                continue;
            }

            switch (character)
            {
                case '\'':
                    inQuotedSheet = true;
                    break;
                case '[':
                    bracketDepth++;
                    break;
                case ']':
                    throw InvalidAddress(sourceAddress, "It has an unmatched ']'.");
                case ',':
                    blocks.Add(TakeBlock(sourceAddress, start, index));
                    start = index + 1;
                    break;
            }
        }

        if (inQuotedSheet)
        {
            throw InvalidAddress(sourceAddress, "A quoted sheet name is missing its closing apostrophe.");
        }

        if (bracketDepth != 0)
        {
            throw InvalidAddress(sourceAddress, "It has an unmatched '['.");
        }

        blocks.Add(TakeBlock(sourceAddress, start, sourceAddress.Length));
        return blocks;
    }

    private static string TakeBlock(string sourceAddress, int start, int end)
    {
        var block = sourceAddress[start..end].Trim();
        if (block.Length == 0)
        {
            throw InvalidAddress(sourceAddress, "It contains an empty block between commas.");
        }

        return block;
    }

    private static (string? SheetName, string LocalAddress) SplitSheetPrefix(string sourceAddress, string block)
    {
        if (block[0] == '\'')
        {
            var sheetName = new System.Text.StringBuilder();
            int index = 1;
            for (; index < block.Length; index++)
            {
                if (block[index] == '\'')
                {
                    if (index + 1 < block.Length && block[index + 1] == '\'')
                    {
                        sheetName.Append('\'');
                        index++;
                        continue;
                    }

                    break;
                }

                sheetName.Append(block[index]);
            }

            if (index + 1 >= block.Length || block[index + 1] != '!' || sheetName.Length == 0)
            {
                throw InvalidAddress(
                    sourceAddress,
                    $"Block '{block}' must use the form 'Sheet Name'!A1:B10.");
            }

            return (sheetName.ToString(), RequireLocal(sourceAddress, block, block[(index + 2)..]));
        }

        int bang = block.IndexOf('!', StringComparison.Ordinal);
        int bracket = block.IndexOf('[', StringComparison.Ordinal);
        if (bang < 0 || (bracket >= 0 && bracket < bang))
        {
            return (null, block);
        }

        var unquotedSheet = block[..bang].Trim();
        if (unquotedSheet.Length == 0)
        {
            throw InvalidAddress(sourceAddress, $"Block '{block}' has an empty sheet name.");
        }

        return (unquotedSheet, RequireLocal(sourceAddress, block, block[(bang + 1)..]));
    }

    private static string RequireLocal(string sourceAddress, string block, string local)
    {
        local = local.Trim();
        if (local.Length == 0 || local.Contains('!', StringComparison.Ordinal))
        {
            throw InvalidAddress(sourceAddress, $"Block '{block}' has no valid cell address after the sheet name.");
        }

        return local;
    }

    /// <summary>
    /// Returns the sheet's name as stored in the workbook, because Excel matches sheet names
    /// regardless of capital letters. A missing sheet keeps the given name so the range lookup reports it.
    /// </summary>
    private static string ActualSheetName(dynamic book, string sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
        {
            return sheetName;
        }

        Excel.Sheets? worksheets = null;
        Excel.Worksheet? sheet = null;
        try
        {
            worksheets = (Excel.Sheets)book.Worksheets;
            for (int index = 1; index <= worksheets.Count; index++)
            {
                sheet = worksheets[index];
                if (string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return sheet.Name;
                }

                ComUtilities.Release(ref sheet);
            }

            return sheetName;
        }
        finally
        {
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref worksheets);
        }
    }

    private static dynamic ResolveDefinedName(dynamic book, string sourceAddress, OperationFailureException addressError)
    {
        try
        {
            return RangeHelpers.ResolveRange(book, string.Empty, sourceAddress.Trim())!;
        }
        catch (OperationFailureException ex) when (ex.ErrorCategory == OperationFailureCategory.NotFound)
        {
            throw InvalidAddress(sourceAddress, addressError.Message, addressError);
        }
    }

    private static OperationFailureException InvalidAddress(
        string sourceAddress, string detail, Exception? inner = null) =>
        new(
            OperationFailureCategory.InvalidInput,
            $"Chart source range '{sourceAddress}' is not valid. {detail} " +
            "Use cell blocks such as 'A1:D10' or 'A1:A10,C1:D10', optionally with a sheet name such as " +
            "'Sales Data'!A1:D10. Separate blocks with commas; all blocks must be on the same sheet.",
            inner);
}
