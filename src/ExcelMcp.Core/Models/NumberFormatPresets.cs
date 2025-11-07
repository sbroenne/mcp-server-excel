namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Common Excel number format codes for LLM convenience
/// These are standard Excel format codes that can be used with SetNumberFormatAsync
/// </summary>
public static class NumberFormatPresets
{
    /// <summary>Currency format with two decimal places: $1,234.56</summary>
    public const string Currency = "$#,##0.00";

    /// <summary>Currency format without decimals: $1,235</summary>
    public const string CurrencyNoDecimals = "$#,##0";

    /// <summary>Currency format with negative numbers in red: $1,234.56 or ($1,234.56)</summary>
    public const string CurrencyNegativeRed = "$#,##0.00_);[Red]($#,##0.00)";

    /// <summary>Percentage format with two decimal places: 12.34%</summary>
    public const string Percentage = "0.00%";

    /// <summary>Percentage format without decimals: 12%</summary>
    public const string PercentageNoDecimals = "0%";

    /// <summary>Percentage format with one decimal place: 12.3%</summary>
    public const string PercentageOneDecimal = "0.0%";

    /// <summary>Short date format: 1/15/2025</summary>
    public const string DateShort = "m/d/yyyy";

    /// <summary>Long date format: January 15, 2025</summary>
    public const string DateLong = "mmmm d, yyyy";

    /// <summary>Month and year format: Jan 2025</summary>
    public const string DateMonthYear = "mmm yyyy";

    /// <summary>Day/month/year format (European): 15/01/2025</summary>
    public const string DateDayMonth = "dd/mm/yyyy";

    /// <summary>12-hour time format: 1:30 PM</summary>
    public const string Time12Hour = "h:mm AM/PM";

    /// <summary>24-hour time format: 13:30</summary>
    public const string Time24Hour = "h:mm";

    /// <summary>Date and time format: 1/15/2025 13:30</summary>
    public const string DateTime = "m/d/yyyy h:mm";

    /// <summary>Number format with two decimal places and thousands separator: 1,234.56</summary>
    public const string Number = "#,##0.00";

    /// <summary>Number format without decimals: 1,235</summary>
    public const string NumberNoDecimals = "#,##0";

    /// <summary>Number format with one decimal place: 1,234.6</summary>
    public const string NumberOneDecimal = "#,##0.0";

    /// <summary>Scientific notation format: 1.23E+03</summary>
    public const string Scientific = "0.00E+00";

    /// <summary>Text format (forces value to be treated as text): @</summary>
    public const string Text = "@";

    /// <summary>Fraction format: 1 1/4</summary>
    public const string Fraction = "# ?/?";

    /// <summary>Accounting format with aligned currency symbols and dashes for zeros</summary>
    public const string Accounting = "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)";

    /// <summary>General format (Excel's default automatic format)</summary>
    public const string General = "General";
}
