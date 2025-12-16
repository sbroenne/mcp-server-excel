using System.Text;

namespace Sbroenne.ExcelMcp.ComInterop.Formatting;

/// <summary>
/// Translates number and date/time format codes between US (English) format and the locale-specific format
/// that Excel expects based on the current system locale.
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

    /// <summary>
    /// Creates a new NumberFormatTranslator by reading locale codes from the Excel Application.
    /// </summary>
    /// <param name="excelApp">The Excel.Application COM object (dynamic)</param>
    public NumberFormatTranslator(dynamic excelApp)
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

        // Check if this is already English locale for dates (no translation needed)
        IsEnglishDateLocale = DayCode.Equals("d", StringComparison.OrdinalIgnoreCase) &&
                               MonthCode.Equals("m", StringComparison.OrdinalIgnoreCase) &&
                               YearCode.Equals("y", StringComparison.OrdinalIgnoreCase);

        // Check if this is already English locale for numbers (no translation needed)
        IsEnglishNumberLocale = DecimalSeparator == "." && ThousandsSeparator == ",";
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
    private string TranslateFormatString(string format)
    {
        var result = new StringBuilder(format.Length);
        int i = 0;

        // Track if we're in a time context (after seeing 'h' or ':')
        bool inTimeContext = false;

        while (i < format.Length)
        {
            char c = format[i];

            // Skip content in square brackets (locale prefixes, colors, conditions)
            if (c == '[')
            {
                int bracketEnd = format.IndexOf(']', i);
                if (bracketEnd > i)
                {
                    result.Append(format.AsSpan(i, bracketEnd - i + 1));
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

            // Handle decimal separator '.' in number format context
            // A '.' is a decimal separator if it's followed by a digit placeholder (0 or #)
            if (c == '.' && !IsEnglishNumberLocale)
            {
                if (i + 1 < format.Length && IsDigitPlaceholder(format[i + 1]))
                {
                    // This is a decimal separator in a number format - translate it
                    result.Append(DecimalSeparator);
                    i++;
                    continue;
                }
            }

            // Handle thousands separator ',' in number format context
            // A ',' is a thousands separator if it's between digit placeholders
            if (c == ',' && !IsEnglishNumberLocale)
            {
                // Check if this is a thousands separator (surrounded by digit placeholders)
                bool prevIsDigit = i > 0 && (IsDigitPlaceholder(format[i - 1]) || format[i - 1] == '.');
                bool nextIsDigit = i + 1 < format.Length && (IsDigitPlaceholder(format[i + 1]) || format[i + 1] == '#' || format[i + 1] == '0');

                if (prevIsDigit && nextIsDigit)
                {
                    // This is a thousands separator in a number format - translate it
                    result.Append(ThousandsSeparator);
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
            if (c == 'h' || c == 'H')
            {
                inTimeContext = true;
                int count = CountRepeatingChar(format, i, c);
                if (!IsEnglishDateLocale)
                {
                    result.Append(HourCode[0], count);
                }
                else
                {
                    result.Append(c, count);
                }
                i += count;
                continue;
            }

            // Second code
            if (c == 's' || c == 'S')
            {
                int count = CountRepeatingChar(format, i, c);
                if (!IsEnglishDateLocale)
                {
                    result.Append(SecondCode[0], count);
                }
                else
                {
                    result.Append(c, count);
                }
                i += count;
                continue;
            }

            // Day code - 'd' or 'D'
            if ((c == 'd' || c == 'D') && !IsEnglishDateLocale)
            {
                int count = CountRepeatingChar(format, i, c);

                // ddd and dddd are weekday names - keep as-is
                if (count >= 3)
                {
                    result.Append(c, count);
                }
                else
                {
                    // d or dd = day number
                    result.Append(DayCode[0], count);
                }
                i += count;
                continue;
            }

            // Month/Minute code - 'm' or 'M'
            // This is the tricky one - 'm' means month in date context, minutes in time context
            if ((c == 'm' || c == 'M') && !IsEnglishDateLocale)
            {
                int count = CountRepeatingChar(format, i, c);

                if (inTimeContext)
                {
                    // In time context, m = minutes
                    result.Append(MinuteCode[0], count);
                }
                else
                {
                    // In date context, m = month
                    // mmm and mmmm are month names - keep as-is (Excel handles translation)
                    if (count >= 3)
                    {
                        result.Append(c, count);
                    }
                    else
                    {
                        // m or mm = month number
                        result.Append(MonthCode[0], count);
                    }
                }
                i += count;
                continue;
            }

            // Year code - 'y' or 'Y'
            if ((c == 'y' || c == 'Y') && !IsEnglishDateLocale)
            {
                int count = CountRepeatingChar(format, i, c);
                result.Append(YearCode[0], count);
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
    private static string? GetInternationalValue(dynamic excelApp, int index)
    {
        try
        {
            // Access the International property with the index
            // Excel COM: excelApp.International(index) returns the locale-specific value
            object? value = excelApp.International[index];
            return value?.ToString();
        }
        catch (Exception ex) when (ex is System.Runtime.InteropServices.COMException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
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
