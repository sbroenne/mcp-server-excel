using System.Runtime.InteropServices;
using System.Text;

namespace Sbroenne.ExcelMcp.Core.DataModel;

/// <summary>
/// Keeps Excel from reading a DAX argument comma as a decimal mark.
/// </summary>
/// <remarks>
/// Excel's measure formula property (<c>ModelMeasures.Add</c> and <c>ModelMeasure.Formula</c>)
/// converts DAX from the Windows user number format before the Data Model sees it. When the
/// Windows decimal mark is a comma, a comma that touches a number becomes a decimal point, so
/// <c>IF(1=1, 1.5, 0)</c> reaches the engine as <c>IF(1=1. 1.5. 0)</c>. The thread culture, the
/// LCID, and Excel's own separator override do not change this. A space between the number and
/// the comma is stored and evaluated correctly with either Windows list separator (issue #978).
/// </remarks>
internal static class DaxNumericCommaSpacer
{
    internal const string AdjustmentMessage =
        "Spaces were added next to commas that touch a number in the DAX formula, because Windows uses a " +
        "comma as the decimal mark on this computer and Excel would otherwise read those commas as decimal " +
        "points. The formula's meaning is unchanged.";

    private const uint LocaleSDecimal = 0x0000000E;

    /// <summary>
    /// True when the Windows user decimal mark is a comma, so measure writes need the spacing.
    /// </summary>
    internal static bool IsNeededOnThisComputer() => IsNeededFor(ReadWindowsDecimalSeparator());

    internal static bool IsNeededFor(string? decimalSeparator) =>
        string.Equals(decimalSeparator, ",", StringComparison.Ordinal);

    /// <summary>
    /// Adds a space between a number and a comma that touches it, outside strings, quoted
    /// table names, bracketed column names, and comments. Everything else is left unchanged.
    /// </summary>
    internal static string AddSpaces(string formula)
    {
        if (string.IsNullOrEmpty(formula) || !formula.Contains(','))
        {
            return formula;
        }

        var result = new StringBuilder(formula.Length + 8);
        bool previousIsNumber = false;
        int i = 0;

        while (i < formula.Length)
        {
            char c = formula[i];
            int start = i;

            if (c == '"' || c == '\'')
            {
                i = SkipQuoted(formula, i, c);
            }
            else if (c == '[')
            {
                i = SkipQuoted(formula, i, ']');
            }
            else if ((c == '/' && Peek(formula, i + 1) == '/') || (c == '-' && Peek(formula, i + 1) == '-'))
            {
                i = SkipLineComment(formula, i);
            }
            else if (c == '/' && Peek(formula, i + 1) == '*')
            {
                int end = formula.IndexOf("*/", i + 2, StringComparison.Ordinal);
                i = end < 0 ? formula.Length : end + 2;
            }
            else if (IsNumberStart(formula, i))
            {
                while (i < formula.Length && (char.IsAsciiDigit(formula[i]) || formula[i] == '.'))
                {
                    i++;
                }

                if (Peek(formula, i) is 'E' or 'e')
                {
                    int exponentStart = i + 1;
                    if (Peek(formula, exponentStart) is '+' or '-')
                    {
                        exponentStart++;
                    }

                    int exponentEnd = exponentStart;
                    while (char.IsAsciiDigit(Peek(formula, exponentEnd)))
                    {
                        exponentEnd++;
                    }

                    if (exponentEnd > exponentStart)
                    {
                        i = exponentEnd;
                    }
                }

                result.Append(formula, start, i - start);
                previousIsNumber = true;
                continue;
            }
            else if (IsIdentifierChar(c))
            {
                while (i < formula.Length && IsIdentifierChar(formula[i]))
                {
                    i++;
                }
            }
            else if (c == ',')
            {
                if (previousIsNumber)
                {
                    result.Append(' ');
                }

                result.Append(c);
                if (IsNumberStart(formula, i + 1))
                {
                    result.Append(' ');
                }

                i++;
                previousIsNumber = false;
                continue;
            }
            else
            {
                i++;
            }

            result.Append(formula, start, i - start);
            previousIsNumber = false;
        }

        return result.ToString();
    }

    private static char Peek(string formula, int index) =>
        index < formula.Length ? formula[index] : '\0';

    private static bool IsIdentifierChar(char c) => char.IsLetterOrDigit(c) || c == '_';

    private static bool IsNumberStart(string formula, int index)
    {
        if (index >= formula.Length)
        {
            return false;
        }

        char c = formula[index];
        bool startsNumber = char.IsAsciiDigit(c) || (c == '.' && char.IsAsciiDigit(Peek(formula, index + 1)));
        return startsNumber && (index == 0 || !IsIdentifierChar(formula[index - 1]));
    }

    /// <summary>
    /// Returns the index after the closing character. A doubled closing character is an escape.
    /// </summary>
    private static int SkipQuoted(string formula, int openIndex, char close)
    {
        int i = openIndex + 1;
        while (i < formula.Length)
        {
            if (formula[i] == close)
            {
                if (Peek(formula, i + 1) == close)
                {
                    i += 2;
                    continue;
                }

                return i + 1;
            }

            i++;
        }

        return formula.Length;
    }

    private static int SkipLineComment(string formula, int start)
    {
        int end = formula.IndexOfAny(['\r', '\n'], start);
        return end < 0 ? formula.Length : end;
    }

    private static string? ReadWindowsDecimalSeparator()
    {
        if (!OperatingSystem.IsWindows())
        {
            return null;
        }

        var buffer = new char[8];
        int length = GetLocaleInfoEx(null, LocaleSDecimal, buffer, buffer.Length);
        return length > 1 ? new string(buffer, 0, length - 1) : null;
    }

    // Null locale name = LOCALE_NAME_USER_DEFAULT, including the user's customized separators,
    // which is what Excel reads. The .NET thread culture may differ from it.
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern int GetLocaleInfoEx(string? lpLocaleName, uint lcType, [Out] char[] lpLCData, int cchData);
}
