using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.ComInterop.Formatting;

/// <summary>
/// Translates invariant number and date/time codes for Range.NumberFormatLocal and
/// chart tick-label formats. Typed Range.NumberFormat reads do not need translation.
/// </summary>
/// <remarks>
/// <para><b>Why This Is Needed:</b></para>
/// <para>
/// Excel interprets format code characters based on the system locale:
/// </para>
/// <list type="bullet">
/// <item>Date codes: On German systems, 'd' (day), 'm' (month), 'y' (year) must be 'T', 'M', 'J'</item>
/// <item>Number separators: On German systems, '.' (decimal) and ',' (thousands) are swapped</item>
/// </list>
/// <para>
/// This translator reads the locale-specific codes from Excel's <c>Application.International</c> property
/// and translates US format codes to locale format codes.
/// </para>
/// <para><b>Usage:</b></para>
/// <code>
/// var translator = new NumberFormatTranslator(excelApp);
/// string dateFormat = translator.TranslateToLocale("m/d/yyyy");   // Returns "M/T/JJJJ" on German
/// string currencyFormat = translator.TranslateToLocale("$#,##0.00"); // Returns "$#.##0,00" on German
/// </code>
/// </remarks>
public sealed class NumberFormatTranslator
{
    // XlApplicationInternational enum values for date/time
    private const int XlDayCode = 21;
    private const int XlMonthCode = 20;
    private const int XlYearCode = 19;
    private const int XlHourCode = 22;
    private const int XlMinuteCode = 23;
    private const int XlSecondCode = 24;
    private const int XlDateSeparator = 17;
    private const int XlTimeSeparator = 18;

    // XlApplicationInternational enum values for number separators
    private const int XlDecimalSeparator = 3;
    private const int XlThousandsSeparator = 4;

    /// <summary>Locale-specific day code (e.g., 'd' for English, 'T' for German)</summary>
    public string DayCode { get; }

    /// <summary>Locale-specific month code (e.g., 'm' for English, 'M' for German)</summary>
    public string MonthCode { get; }

    /// <summary>Locale-specific year code (e.g., 'y' for English, 'J' for German)</summary>
    public string YearCode { get; }

    /// <summary>Locale-specific hour code (typically 'h' across locales)</summary>
    public string HourCode { get; }

    /// <summary>Locale-specific minute code (typically 'm' across locales - same as month!)</summary>
    public string MinuteCode { get; }

    /// <summary>Locale-specific second code (typically 's' across locales)</summary>
    public string SecondCode { get; }

    /// <summary>Locale-specific date separator (e.g., '/' or '.')</summary>
    public string DateSeparator { get; }

    /// <summary>Locale-specific time separator (typically ':')</summary>
    public string TimeSeparator { get; }

    /// <summary>Locale-specific decimal separator (e.g., '.' for English, ',' for German)</summary>
    public string DecimalSeparator { get; }

    /// <summary>Locale-specific thousands separator (e.g., ',' for English, '.' for German)</summary>
    public string ThousandsSeparator { get; }

    /// <summary>True if locale uses same codes as US English (d/m/y)</summary>
    public bool IsEnglishDateLocale { get; }

    /// <summary>True if locale uses same number separators as US English (. for decimal, , for thousands)</summary>
    public bool IsEnglishNumberLocale { get; }

    internal string GeneralFormatName { get; }

    /// <summary>
    /// Creates a new NumberFormatTranslator by reading locale codes from the Excel Application.
    /// </summary>
    /// <param name="excelApp">The Excel.Application COM object</param>
    public NumberFormatTranslator(Excel.Application excelApp)
    {
        // Read locale-specific codes from Excel's International property
        DayCode = GetInternationalValue(excelApp, XlDayCode) ?? "d";
        MonthCode = GetInternationalValue(excelApp, XlMonthCode) ?? "m";
        YearCode = GetInternationalValue(excelApp, XlYearCode) ?? "y";
        HourCode = GetInternationalValue(excelApp, XlHourCode) ?? "h";
        MinuteCode = GetInternationalValue(excelApp, XlMinuteCode) ?? "m";
        SecondCode = GetInternationalValue(excelApp, XlSecondCode) ?? "s";
        DateSeparator = GetInternationalValue(excelApp, XlDateSeparator) ?? "/";
        TimeSeparator = GetInternationalValue(excelApp, XlTimeSeparator) ?? ":";

        // Read number separators
        DecimalSeparator = GetInternationalValue(excelApp, XlDecimalSeparator) ?? ".";
        ThousandsSeparator = GetInternationalValue(excelApp, XlThousandsSeparator) ?? ",";
        GeneralFormatName = GetInternationalValue(excelApp, (int)Excel.XlApplicationInternational.xlGeneralFormatName)
            ?? throw new InvalidOperationException("Excel did not provide its General number-format name.");

        // Check if this is already English locale for dates (no translation needed)
        IsEnglishDateLocale = DayCode.Equals("d", StringComparison.OrdinalIgnoreCase) &&
                               MonthCode.Equals("m", StringComparison.OrdinalIgnoreCase) &&
                               YearCode.Equals("y", StringComparison.OrdinalIgnoreCase);

        // Check if this is already English locale for numbers (no translation needed)
        IsEnglishNumberLocale = DecimalSeparator == "." && ThousandsSeparator == ",";
    }

    internal NumberFormatTranslator(string decimalSeparator, string thousandsSeparator, string generalFormatName = "General")
    {
        DayCode = "d";
        MonthCode = MinuteCode = "m";
        YearCode = "y";
        HourCode = "h";
        SecondCode = "s";
        DateSeparator = "/";
        TimeSeparator = ":";
        DecimalSeparator = decimalSeparator;
        ThousandsSeparator = thousandsSeparator;
        GeneralFormatName = generalFormatName;
        IsEnglishDateLocale = true;
        IsEnglishNumberLocale = decimalSeparator == "." && thousandsSeparator == ",";
    }

    /// <summary>
    /// Translates a US (English) format string to the locale-specific format Excel expects.
    /// Handles both date/time codes and number separators.
    /// </summary>
    /// <param name="usFormat">US format string (e.g., "m/d/yyyy", "$#,##0.00")</param>
    /// <returns>Locale-specific format string (e.g., "M/T/JJJJ", "$#.##0,00" on German Excel)</returns>
    /// <remarks>
    /// <para>Translation rules:</para>
    /// <list type="bullet">
    /// <item>'d' or 'dd' (day) → locale day code (e.g., 'T' or 'TT' on German)</item>
    /// <item>'ddd' or 'dddd' (weekday names) → kept as-is (Excel handles these)</item>
    /// <item>'m' or 'mm' (month, when NOT after time separator) → locale month code</item>
    /// <item>'mmm' or 'mmmm' (month names) → kept as-is (Excel handles these)</item>
    /// <item>'y' or 'yy' or 'yyyy' (year) → locale year code</item>
    /// <item>'h', 'm' (after :), 's' (time) → locale time codes</item>
    /// <item>'.' (decimal separator in number formats) → locale decimal separator</item>
    /// <item>',' (thousands separator in number formats) → locale thousands separator</item>
    /// <item>Literal text in quotes or brackets is preserved</item>
    /// </list>
    /// </remarks>
    public string TranslateToLocale(string usFormat)
    {
        if (string.IsNullOrEmpty(usFormat))
            return usFormat;

        // If already English locale for both dates and numbers, no translation needed
        if (IsEnglishDateLocale && IsEnglishNumberLocale)
            return usFormat;

        // Don't translate if it already contains locale-specific codes
        // (user might have already used German codes)
        if (ContainsLocaleSpecificCodes(usFormat))
            return usFormat;

        // Parse and translate the format string
        return TranslateFormatString(usFormat);
    }

    /// <summary>Converts a localized format returned by chart axes to invariant format codes.</summary>
    public string TranslateFromLocale(string localFormat)
    {
        if (string.Equals(localFormat, GeneralFormatName, StringComparison.OrdinalIgnoreCase))
            return "General";

        return string.IsNullOrEmpty(localFormat) ||
            (IsEnglishDateLocale && IsEnglishNumberLocale &&
             GeneralFormatName.Equals("General", StringComparison.OrdinalIgnoreCase))
            ? localFormat
            : TranslateFormatString(localFormat, toInvariant: true);
    }

    /// <summary>
    /// Checks if the format string already contains locale-specific date codes.
    /// </summary>
    private bool ContainsLocaleSpecificCodes(string format)
    {
        // Check for German-style codes (case-insensitive)
        // T = Tag (day), J = Jahr (year) are unique to German
        // We check for these to avoid double-translation
        if (!DayCode.Equals("d", StringComparison.OrdinalIgnoreCase) &&
            format.Contains(DayCode, StringComparison.OrdinalIgnoreCase))
            return true;

        if (!YearCode.Equals("y", StringComparison.OrdinalIgnoreCase) &&
            format.Contains(YearCode, StringComparison.OrdinalIgnoreCase))
            return true;

        return false;
    }

    /// <summary>
    /// Translates format string character by character, handling context (date vs time vs number).
    /// </summary>
    private string TranslateFormatString(string format, bool toInvariant = false)
    {
        var result = new StringBuilder(format.Length);
        int i = 0;
        char sourceDecimal = toInvariant ? DecimalSeparator[0] : '.';
        char sourceThousands = toInvariant ? ThousandsSeparator[0] : ',';

        // Track if we're in a time context (after seeing 'h' or ':')
        bool inTimeContext = false;

        while (i < format.Length)
        {
            char c = format[i];

            // Comparison constants use local decimals; colour and locale metadata stay unchanged.
            if (c == '[')
            {
                int bracketEnd = format.IndexOf(']', i);
                if (bracketEnd > i)
                {
                    if (format[i + 1] is '<' or '>' or '=')
                    {
                        result.Append('[');
                        for (int index = i + 1; index < bracketEnd; index++)
                        {
                            if (format[index] == sourceDecimal)
                                result.Append(toInvariant ? "." : DecimalSeparator);
                            else
                                result.Append(format[index]);
                        }
                        result.Append(']');
                    }
                    else
                    {
                        result.Append(format.AsSpan(i, bracketEnd - i + 1));
                    }
                    i = bracketEnd + 1;
                    continue;
                }
            }

            // Skip content in quotes (literal text)
            if (c == '"')
            {
                int quoteEnd = format.IndexOf('"', i + 1);
                if (quoteEnd > i)
                {
                    result.Append(format.AsSpan(i, quoteEnd - i + 1));
                    i = quoteEnd + 1;
                    continue;
                }
            }

            // Skip escaped characters (backslash)
            if (c == '\\' && i + 1 < format.Length)
            {
                result.Append(format.AsSpan(i, 2));
                i += 2;
                continue;
            }

            if (toInvariant && GeneralFormatName.Length > 0 &&
                format.AsSpan(i).StartsWith(GeneralFormatName, StringComparison.OrdinalIgnoreCase))
            {
                result.Append("General");
                i += GeneralFormatName.Length;
                continue;
            }

            // Handle decimal separator '.' in number format context
            // A '.' is a decimal separator if it's followed by a digit placeholder (0 or #)
            if (c == sourceDecimal && !IsEnglishNumberLocale)
            {
                if (i + 1 < format.Length && IsDigitPlaceholder(format[i + 1]))
                {
                    // This is a decimal separator in a number format - translate it
                    result.Append(toInvariant ? "." : DecimalSeparator);
                    i++;
                    continue;
                }
            }

            // Commas after numeric placeholders also scale by thousands, including repeated commas.
            if (c == sourceThousands && !IsEnglishNumberLocale)
            {
                int previous = i - 1;
                while (previous >= 0 && format[previous] == sourceThousands)
                {
                    previous--;
                }
                if (previous >= 0 && IsDigitPlaceholder(format[previous]))
                {
                    result.Append(toInvariant ? "," : ThousandsSeparator);
                    i++;
                    continue;
                }
            }

            // Time separator - switch to time context
            if (c == ':')
            {
                inTimeContext = true;
                result.Append(c);
                i++;
                continue;
            }

            // Hour code - switch to time context
            if (char.ToLowerInvariant(c) == char.ToLowerInvariant(toInvariant ? HourCode[0] : 'h'))
            {
                inTimeContext = true;
                int count = CountRepeatingChar(format, i, c);
                if (!IsEnglishDateLocale)
                {
                    result.Append(toInvariant ? 'h' : HourCode[0], count);
                }
                else
                {
                    result.Append(c, count);
                }
                i += count;
                continue;
            }

            // Second code
            if (char.ToLowerInvariant(c) == char.ToLowerInvariant(toInvariant ? SecondCode[0] : 's'))
            {
                int count = CountRepeatingChar(format, i, c);
                if (!IsEnglishDateLocale)
                {
                    result.Append(toInvariant ? 's' : SecondCode[0], count);
                }
                else
                {
                    result.Append(c, count);
                }
                i += count;
                continue;
            }

            // Day code - 'd' or 'D'
            if (char.ToLowerInvariant(c) == char.ToLowerInvariant(toInvariant ? DayCode[0] : 'd') && !IsEnglishDateLocale)
            {
                int count = CountRepeatingChar(format, i, c);

                // ddd and dddd are weekday names - keep as-is
                if (count >= 3 && !toInvariant)
                {
                    result.Append(c, count);
                }
                else
                {
                    // d or dd = day number
                    result.Append(toInvariant ? 'd' : DayCode[0], count);
                }
                i += count;
                continue;
            }

            // Month/Minute code - 'm' or 'M'
            // This is the tricky one - 'm' means month in date context, minutes in time context
            if (char.ToLowerInvariant(c) == char.ToLowerInvariant(toInvariant ? (inTimeContext ? MinuteCode[0] : MonthCode[0]) : 'm') && !IsEnglishDateLocale)
            {
                int count = CountRepeatingChar(format, i, c);

                if (inTimeContext)
                {
                    // In time context, m = minutes
                    result.Append(toInvariant ? 'm' : MinuteCode[0], count);
                }
                else
                {
                    // In date context, m = month
                    // mmm and mmmm are month names - keep as-is (Excel handles translation)
                    if (count >= 3 && !toInvariant)
                    {
                        result.Append(c, count);
                    }
                    else
                    {
                        // m or mm = month number
                        result.Append(toInvariant ? 'm' : MonthCode[0], count);
                    }
                }
                i += count;
                continue;
            }

            // Year code - 'y' or 'Y'
            if (char.ToLowerInvariant(c) == char.ToLowerInvariant(toInvariant ? YearCode[0] : 'y') && !IsEnglishDateLocale)
            {
                int count = CountRepeatingChar(format, i, c);
                result.Append(toInvariant ? 'y' : YearCode[0], count);
                i += count;
                continue;
            }

            // Other characters pass through unchanged
            result.Append(c);
            i++;

            // Reset time context on section separator
            if (c == ';')
            {
                inTimeContext = false;
            }
        }

        return result.ToString();
    }

    /// <summary>
    /// Checks if a character is a digit placeholder in Excel number formats.
    /// </summary>
    private static bool IsDigitPlaceholder(char c) => c == '0' || c == '#' || c == '?';

    /// <summary>
    /// Counts how many times a character repeats starting at position.
    /// </summary>
    private static int CountRepeatingChar(string format, int startIndex, char c)
    {
        int count = 0;
        char lowerC = char.ToLowerInvariant(c);

        while (startIndex + count < format.Length &&
               char.ToLowerInvariant(format[startIndex + count]) == lowerC)
        {
            count++;
        }

        return count;
    }

    /// <summary>
    /// Gets a value from Excel's International property.
    /// </summary>
    private static string? GetInternationalValue(Excel.Application excelApp, int index)
    {
        try
        {
            // Access the International property with the index
            // Excel COM: excelApp.International(index) returns the locale-specific value
            object? value = excelApp.International[(Excel.XlApplicationInternational)index];
            return value?.ToString();
        }
        catch (Exception ex) when (ex is System.Runtime.InteropServices.COMException)
        {
            // International property access failed for this index
            return null;
        }
    }

    /// <summary>
    /// Returns a summary of the locale codes for debugging/logging.
    /// </summary>
    public override string ToString()
    {
        return $"NumberFormatTranslator: Day='{DayCode}' Month='{MonthCode}' Year='{YearCode}' " +
               $"Hour='{HourCode}' Minute='{MinuteCode}' Second='{SecondCode}' " +
               $"DateSep='{DateSeparator}' TimeSep='{TimeSeparator}' " +
               $"DecimalSep='{DecimalSeparator}' ThousandsSep='{ThousandsSeparator}' " +
               $"IsEnglishDate={IsEnglishDateLocale} IsEnglishNumber={IsEnglishNumberLocale}";
    }
}
