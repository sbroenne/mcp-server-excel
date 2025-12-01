using System.Globalization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Excel number format categories matching Format Cells dialog
/// </summary>
public enum NumberFormatCategory
{
    /// <summary>No specific format - Excel chooses based on value</summary>
    General,

    /// <summary>Numeric format with optional decimals and thousands separator</summary>
    Number,

    /// <summary>Currency format with symbol, decimals, and negative style</summary>
    Currency,

    /// <summary>Accounting format with aligned currency symbols</summary>
    Accounting,

    /// <summary>Date format</summary>
    Date,

    /// <summary>Time format</summary>
    Time,

    /// <summary>Percentage format</summary>
    Percentage,

    /// <summary>Fraction format</summary>
    Fraction,

    /// <summary>Scientific notation</summary>
    Scientific,

    /// <summary>Text format - values treated as text</summary>
    Text,

    /// <summary>Special formats (ZIP code, phone number, SSN)</summary>
    Special
}

/// <summary>
/// How to display negative numbers
/// </summary>
public enum NegativeNumberFormat
{
    /// <summary>-1234.56 (minus sign prefix)</summary>
    MinusSign,

    /// <summary>1234.56 in red color (no sign)</summary>
    Red,

    /// <summary>(1234.56) in parentheses</summary>
    Parentheses,

    /// <summary>(1234.56) in red with parentheses</summary>
    RedParentheses
}

/// <summary>
/// Predefined date format styles
/// </summary>
public enum DateFormatStyle
{
    /// <summary>System short date (locale-dependent)</summary>
    ShortDate,

    /// <summary>System long date (locale-dependent)</summary>
    LongDate,

    /// <summary>ISO format: 2025-12-31</summary>
    ISO,

    /// <summary>Month and Year: December 2025</summary>
    MonthYear,

    /// <summary>Day and Month: 31 December</summary>
    DayMonth,

    /// <summary>Year only: 2025</summary>
    Year,

    /// <summary>Month only: December</summary>
    Month,

    /// <summary>Day only: 31</summary>
    Day
}

/// <summary>
/// Predefined time format styles
/// </summary>
public enum TimeFormatStyle
{
    /// <summary>Short time: 13:30 or 1:30 PM (locale-dependent)</summary>
    ShortTime,

    /// <summary>Long time with seconds: 13:30:45</summary>
    LongTime,

    /// <summary>Duration in hours: [h]:mm:ss</summary>
    Duration,

    /// <summary>Hours and minutes: 13:30</summary>
    HoursMinutes,

    /// <summary>Hours, minutes, seconds: 13:30:45</summary>
    HoursMinutesSeconds
}

/// <summary>
/// Predefined fraction formats
/// </summary>
public enum FractionStyle
{
    /// <summary>Up to one digit: 1/4</summary>
    OneDigit,

    /// <summary>Up to two digits: 21/25</summary>
    TwoDigits,

    /// <summary>Up to three digits: 312/943</summary>
    ThreeDigits,

    /// <summary>Halves: 1/2</summary>
    Halves,

    /// <summary>Quarters: 1/4, 2/4, 3/4</summary>
    Quarters,

    /// <summary>Eighths: 1/8, 3/8, etc.</summary>
    Eighths,

    /// <summary>Sixteenths: 1/16, 3/16, etc.</summary>
    Sixteenths,

    /// <summary>Tenths: 1/10, 3/10, etc.</summary>
    Tenths,

    /// <summary>Hundredths: 1/100, 23/100, etc.</summary>
    Hundredths
}

/// <summary>
/// Special format types
/// </summary>
public enum SpecialFormatType
{
    /// <summary>ZIP Code: 00000</summary>
    ZipCode,

    /// <summary>ZIP Code + 4: 00000-0000</summary>
    ZipCodePlus4,

    /// <summary>Phone Number: (000) 000-0000</summary>
    PhoneNumber,

    /// <summary>Social Security Number: 000-00-0000</summary>
    SocialSecurityNumber
}

/// <summary>
/// Options for structured number formatting (alternative to raw formatCode)
/// </summary>
public class NumberFormatOptions
{
    /// <summary>Format category (required)</summary>
    public NumberFormatCategory Category { get; set; } = NumberFormatCategory.General;

    /// <summary>Number of decimal places (for Number, Currency, Accounting, Percentage, Scientific)</summary>
    public int? DecimalPlaces { get; set; }

    /// <summary>Use thousands separator (for Number, Currency, Accounting)</summary>
    public bool? UseThousandsSeparator { get; set; }

    /// <summary>Currency symbol (for Currency, Accounting). Examples: $, €, £, ¥</summary>
    public string? CurrencySymbol { get; set; }

    /// <summary>How to display negative numbers (for Number, Currency)</summary>
    public NegativeNumberFormat? NegativeFormat { get; set; }

    /// <summary>Date format style (for Date category)</summary>
    public DateFormatStyle? DateFormat { get; set; }

    /// <summary>Time format style (for Time category)</summary>
    public TimeFormatStyle? TimeFormat { get; set; }

    /// <summary>Include date with time (for Time category)</summary>
    public bool? IncludeDate { get; set; }

    /// <summary>Fraction format style (for Fraction category)</summary>
    public FractionStyle? FractionFormat { get; set; }

    /// <summary>Special format type (for Special category)</summary>
    public SpecialFormatType? SpecialFormat { get; set; }
}

/// <summary>
/// Builds locale-aware Excel format codes from structured options
/// </summary>
public static class LocaleAwareFormatBuilder
{
    /// <summary>
    /// Build a locale-aware format code from structured options.
    /// Uses the current system culture for locale-specific formatting.
    /// </summary>
    public static string BuildFormatCode(NumberFormatOptions options)
    {
        return BuildFormatCode(options, CultureInfo.CurrentCulture);
    }

    /// <summary>
    /// Build a format code for a specific culture/locale
    /// </summary>
    public static string BuildFormatCode(NumberFormatOptions options, CultureInfo culture)
    {
        return options.Category switch
        {
            NumberFormatCategory.General => "General",
            NumberFormatCategory.Number => BuildNumberFormat(options, culture),
            NumberFormatCategory.Currency => BuildCurrencyFormat(options, culture),
            NumberFormatCategory.Accounting => BuildAccountingFormat(options, culture),
            NumberFormatCategory.Percentage => BuildPercentageFormat(options, culture),
            NumberFormatCategory.Date => BuildDateFormat(options, culture),
            NumberFormatCategory.Time => BuildTimeFormat(options, culture),
            NumberFormatCategory.Fraction => BuildFractionFormat(options),
            NumberFormatCategory.Scientific => BuildScientificFormat(options, culture),
            NumberFormatCategory.Text => "@",
            NumberFormatCategory.Special => BuildSpecialFormat(options),
            _ => "General"
        };
    }

    private static string BuildNumberFormat(NumberFormatOptions options, CultureInfo culture)
    {
        // NumberFormatLocal interprets format codes using the locale's conventions
        // So we MUST use the locale's actual separators in the format code
        var nf = culture.NumberFormat;
        var decimals = options.DecimalPlaces ?? 2;
        var useThousands = options.UseThousandsSeparator ?? true;

        // Build decimal part using locale's decimal separator
        var decimalPart = decimals > 0
            ? nf.NumberDecimalSeparator + new string('0', decimals)
            : "";

        // Build integer part with or without thousands separator (using locale's separator)
        var integerPart = useThousands
            ? "#" + nf.NumberGroupSeparator + "##0"
            : "0";

        var positiveFormat = integerPart + decimalPart;

        // Handle negative format
        return options.NegativeFormat switch
        {
            NegativeNumberFormat.Red => positiveFormat + ";[Red]" + positiveFormat,
            NegativeNumberFormat.Parentheses => positiveFormat + ";(" + positiveFormat + ")",
            NegativeNumberFormat.RedParentheses => positiveFormat + ";[Red](" + positiveFormat + ")",
            _ => positiveFormat // Default: minus sign (Excel's default behavior)
        };
    }

    private static string BuildCurrencyFormat(NumberFormatOptions options, CultureInfo culture)
    {
        // NumberFormatLocal interprets format codes using the locale's conventions
        var nf = culture.NumberFormat;
        var decimals = options.DecimalPlaces ?? 2;
        var symbol = options.CurrencySymbol ?? nf.CurrencySymbol;
        var useThousands = options.UseThousandsSeparator ?? true;

        // Build decimal part using locale's separator
        var decimalPart = decimals > 0
            ? nf.CurrencyDecimalSeparator + new string('0', decimals)
            : "";

        // Build integer part using locale's separator
        var integerPart = useThousands
            ? "#" + nf.CurrencyGroupSeparator + "##0"
            : "0";

        // Standard currency format: symbol prefix (most common)
        var positiveFormat = symbol + integerPart + decimalPart;

        // Handle negative format
        return options.NegativeFormat switch
        {
            NegativeNumberFormat.Red => positiveFormat + ";[Red]" + positiveFormat,
            NegativeNumberFormat.Parentheses => positiveFormat + ";(" + symbol + integerPart + decimalPart + ")",
            NegativeNumberFormat.RedParentheses => positiveFormat + ";[Red](" + symbol + integerPart + decimalPart + ")",
            _ => positiveFormat + ";-" + symbol + integerPart + decimalPart
        };
    }

    private static string BuildAccountingFormat(NumberFormatOptions options, CultureInfo culture)
    {
        // NumberFormatLocal interprets format codes using the locale's conventions
        var nf = culture.NumberFormat;
        var decimals = options.DecimalPlaces ?? 2;
        var symbol = options.CurrencySymbol ?? nf.CurrencySymbol;

        var decimalPart = decimals > 0
            ? nf.CurrencyDecimalSeparator + new string('0', decimals)
            : "";

        var integerPart = "#" + nf.CurrencyGroupSeparator + "##0";

        // Accounting format: aligned symbols, negative in parentheses, zeros as dashes
        return $"_({symbol}* {integerPart}{decimalPart}_);_({symbol}* ({integerPart}{decimalPart});_({symbol}* \"-\"??_);_(@_)";
    }

    private static string BuildPercentageFormat(NumberFormatOptions options, CultureInfo culture)
    {
        // NumberFormatLocal interprets format codes using the locale's conventions
        var nf = culture.NumberFormat;
        var decimals = options.DecimalPlaces ?? 2;

        var decimalPart = decimals > 0
            ? nf.PercentDecimalSeparator + new string('0', decimals)
            : "";

        return "0" + decimalPart + "%";
    }

    private static string BuildDateFormat(NumberFormatOptions options, CultureInfo culture)
    {
        // Use Excel's built-in format IDs where possible - these are locale-independent
        // Excel interprets them according to the user's regional settings
        // Reference: https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat

        // For NumberFormatLocal, we use standard Excel format codes
        // Excel will interpret these according to the system locale automatically
        _ = culture; // Culture parameter kept for API consistency but not needed - Excel handles locale

        return options.DateFormat switch
        {
            // Standard Excel date format codes - Excel interprets per locale
            DateFormatStyle.ShortDate => "d",           // Built-in short date (locale-aware)
            DateFormatStyle.LongDate => "dddd, mmmm d, yyyy",  // Full date with day name
            DateFormatStyle.ISO => "yyyy-mm-dd",        // ISO 8601 - universal
            DateFormatStyle.MonthYear => "mmmm yyyy",   // December 2025
            DateFormatStyle.DayMonth => "d mmmm",       // 31 December
            DateFormatStyle.Year => "yyyy",             // 2025
            DateFormatStyle.Month => "mmmm",            // December
            DateFormatStyle.Day => "d",                 // 31
            _ => "d"                                    // Default to short date
        };
    }

    private static string BuildTimeFormat(NumberFormatOptions options, CultureInfo culture)
    {
        // Use standard Excel time format codes
        // Excel interprets h vs H and AM/PM based on locale settings automatically
        _ = culture; // Culture parameter kept for API consistency

        var timeFormat = options.TimeFormat switch
        {
            TimeFormatStyle.ShortTime => "h:mm",                // Excel shows AM/PM based on locale
            TimeFormatStyle.LongTime => "h:mm:ss",              // With seconds
            TimeFormatStyle.Duration => "[h]:mm:ss",            // Elapsed time (hours can exceed 24)
            TimeFormatStyle.HoursMinutes => "h:mm",
            TimeFormatStyle.HoursMinutesSeconds => "h:mm:ss",
            _ => "h:mm"
        };

        // Optionally include date
        if (options.IncludeDate == true)
        {
            return "d " + timeFormat;  // Short date + time
        }

        return timeFormat;
    }

    private static string BuildFractionFormat(NumberFormatOptions options)
    {
        return options.FractionFormat switch
        {
            FractionStyle.OneDigit => "# ?/?",
            FractionStyle.TwoDigits => "# ??/??",
            FractionStyle.ThreeDigits => "# ???/???",
            FractionStyle.Halves => "# ?/2",
            FractionStyle.Quarters => "# ?/4",
            FractionStyle.Eighths => "# ?/8",
            FractionStyle.Sixteenths => "# ??/16",
            FractionStyle.Tenths => "# ?/10",
            FractionStyle.Hundredths => "# ??/100",
            _ => "# ?/?"
        };
    }

    private static string BuildScientificFormat(NumberFormatOptions options, CultureInfo culture)
    {
        // NumberFormatLocal interprets format codes using the locale's conventions
        var nf = culture.NumberFormat;
        var decimals = options.DecimalPlaces ?? 2;

        var decimalPart = decimals > 0
            ? nf.NumberDecimalSeparator + new string('0', decimals)
            : "";

        return "0" + decimalPart + "E+00";
    }

    private static string BuildSpecialFormat(NumberFormatOptions options)
    {
        return options.SpecialFormat switch
        {
            SpecialFormatType.ZipCode => "00000",
            SpecialFormatType.ZipCodePlus4 => "00000-0000",
            SpecialFormatType.PhoneNumber => "(000) 000-0000",
            SpecialFormatType.SocialSecurityNumber => "000-00-0000",
            _ => "General"
        };
    }
}
