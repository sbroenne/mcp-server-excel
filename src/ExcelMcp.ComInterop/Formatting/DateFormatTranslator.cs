using System.Text;

namespace Sbroenne.ExcelMcp.ComInterop.Formatting;

/// <summary>
/// Translates date/time format codes between US (English) format and the locale-specific format
/// that Excel expects based on the current system locale.
/// </summary>
/// <remarks>
/// <para><b>Why This Is Needed:</b></para>
/// <para>
/// Excel interprets date format code letters based on the system locale. On German systems,
/// 'd' (day), 'm' (month), 'y' (year) are NOT recognized - Excel expects 'T' (Tag), 'M' (Monat), 'J' (Jahr).
/// </para>
/// <para>
/// This translator reads the locale-specific codes from Excel's <c>Application.International</c> property
/// and translates US format codes (like "m/d/yyyy") to locale format codes (like "M/T/JJJJ" on German).
/// </para>
/// <para><b>Usage:</b></para>
/// <code>
/// var translator = new DateFormatTranslator(excelApp);
/// string localeFormat = translator.TranslateToLocale("m/d/yyyy"); // Returns "M/T/JJJJ" on German Excel
/// </code>
/// </remarks>
public sealed class DateFormatTranslator
{
    // XlApplicationInternational enum values
    private const int XlDayCode = 21;
    private const int XlMonthCode = 20;
    private const int XlYearCode = 19;
    private const int XlHourCode = 22;
    private const int XlMinuteCode = 23;
    private const int XlSecondCode = 24;
    private const int XlDateSeparator = 17;
    private const int XlTimeSeparator = 18;

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

    /// <summary>True if locale uses same codes as US English (d/m/y)</summary>
    public bool IsEnglishLocale { get; }

    /// <summary>
    /// Creates a new DateFormatTranslator by reading locale codes from the Excel Application.
    /// </summary>
    /// <param name="excelApp">The Excel.Application COM object (dynamic)</param>
    public DateFormatTranslator(dynamic excelApp)
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

        // Check if this is already English locale (no translation needed)
        IsEnglishLocale = DayCode.Equals("d", StringComparison.OrdinalIgnoreCase) &&
                          MonthCode.Equals("m", StringComparison.OrdinalIgnoreCase) &&
                          YearCode.Equals("y", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Translates a US (English) date format string to the locale-specific format Excel expects.
    /// </summary>
    /// <param name="usFormat">US format string (e.g., "m/d/yyyy", "mm/dd/yyyy", "yyyy-mm-dd")</param>
    /// <returns>Locale-specific format string (e.g., "M/T/JJJJ" on German Excel)</returns>
    /// <remarks>
    /// <para>Translation rules:</para>
    /// <list type="bullet">
    /// <item>'d' or 'dd' (day) → locale day code (e.g., 'T' or 'TT' on German)</item>
    /// <item>'ddd' or 'dddd' (weekday names) → kept as-is (Excel handles these)</item>
    /// <item>'m' or 'mm' (month, when NOT after time separator) → locale month code</item>
    /// <item>'mmm' or 'mmmm' (month names) → kept as-is (Excel handles these)</item>
    /// <item>'y' or 'yy' or 'yyyy' (year) → locale year code</item>
    /// <item>'h', 'm' (after :), 's' (time) → locale time codes</item>
    /// <item>Literal text in quotes or brackets is preserved</item>
    /// </list>
    /// </remarks>
    public string TranslateToLocale(string usFormat)
    {
        if (string.IsNullOrEmpty(usFormat))
            return usFormat;

        // If already English locale, no translation needed
        if (IsEnglishLocale)
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
    /// Translates format string character by character, handling context (date vs time).
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
                result.Append(HourCode[0], count);
                i += count;
                continue;
            }

            // Second code
            if (c == 's' || c == 'S')
            {
                int count = CountRepeatingChar(format, i, c);
                result.Append(SecondCode[0], count);
                i += count;
                continue;
            }

            // Day code - 'd' or 'D'
            if (c == 'd' || c == 'D')
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
            if (c == 'm' || c == 'M')
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
            if (c == 'y' || c == 'Y')
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
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Returns a summary of the locale codes for debugging/logging.
    /// </summary>
    public override string ToString()
    {
        return $"DateFormatTranslator: Day='{DayCode}' Month='{MonthCode}' Year='{YearCode}' " +
               $"Hour='{HourCode}' Minute='{MinuteCode}' Second='{SecondCode}' " +
               $"DateSep='{DateSeparator}' TimeSep='{TimeSeparator}' IsEnglish={IsEnglishLocale}";
    }
}
