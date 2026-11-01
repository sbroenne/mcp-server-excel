using System.ComponentModel;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

public partial class NamedRangeCommands
{
    private static Excel.Name? FindExistingNameForCreation(Excel.Workbook workbook, string requestedName)
    {
        var (requestedScope, requestedLeaf) = GetCreationIdentity(requestedName);
        var activeScope = requestedScope is null ? GetActiveNameScope(workbook) : null;
        Excel.Names? names = null;
        Excel.Name? candidate = null;
        try
        {
            names = workbook.Names;
            var count = names.Count;
            for (var index = 1; index <= count; index++)
            {
                candidate = names.Item(index);
                var (scope, leaf) = GetCreationIdentity(candidate.Name);
                // Names.Add can update an existing name, including an active-sheet local name.
                if (SameExcelName(leaf, requestedLeaf)
                    && (SameExcelName(scope, requestedScope)
                        || requestedScope is null
                            && scope is not null
                            && SameExcelName(scope, activeScope)))
                {
                    var result = candidate;
                    candidate = null;
                    return result;
                }
                ComUtilities.Release(ref candidate);
            }
            return null;
        }
        finally
        {
            ComUtilities.Release(ref candidate);
            ComUtilities.Release(ref names);
        }
    }

    private static bool SameExcelName(string? left, string? right)
    {
        if (left is null || right is null)
            return left is null && right is null;

        // Desktop Excel uses Windows linguistic equivalence; .NET's ICU and ordinal
        // comparisons miss aliases such as Straße/STRASSE and composed accents.
        var comparison = CompareStringEx(string.Empty, 1, left, left.Length,
            right, right.Length, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
        if (comparison == 0)
            throw new Win32Exception(Marshal.GetLastWin32Error());
        return comparison == 2;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
    private static extern int CompareStringEx(string localeName, uint flags,
        string left, int leftLength, string right, int rightLength,
        IntPtr versionInformation, IntPtr reserved, IntPtr sortHandle);

    private static (string? Scope, string Leaf) GetCreationIdentity(string name)
    {
        // Other Unicode whitespace can be a valid, distinct part of an Excel name.
        var normalized = name.Trim(' ');
        var separator = normalized.LastIndexOf('!');
        if (separator < 0)
            return (null, normalized);

        var scope = normalized[..separator];
        if (scope.Length >= 2 && scope[0] == '\'' && scope[^1] == '\'')
            scope = scope[1..^1].Replace("''", "'", StringComparison.Ordinal);
        return (scope, normalized[(separator + 1)..]);
    }

    private static string? GetActiveNameScope(Excel.Workbook workbook)
    {
        object? activeSheet = null;
        try
        {
            activeSheet = (object?)workbook.ActiveSheet;
            return activeSheet switch
            {
                Excel.Worksheet sheet => sheet.Name,
                Excel.Chart chart => chart.Name,
                _ => null
            };
        }
        finally
        {
            ComUtilities.Release(ref activeSheet);
        }
    }
}
