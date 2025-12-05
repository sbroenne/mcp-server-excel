using System.Globalization;
using System.Text;

namespace Sbroenne.ExcelMcp.ComInterop.Formatting;

/// <summary>
/// Translates DAX formula argument separators between US (English) format (comma) and the
/// locale-specific separator that Excel's Analysis Services engine expects.
/// </summary>
/// <remarks>
/// <para><b>Why This Is Needed:</b></para>
/// <para>
/// Power Pivot's Analysis Services engine uses the SYSTEM CULTURE (not Excel's International
/// settings) to interpret DAX formula separators. This is a critical distinction because:
/// </para>
/// <list type="bullet">
/// <item>Excel's International property may report different settings than the system culture</item>
/// <item>Power Pivot interprets commas based on the system's NumberDecimalSeparator</item>
/// <item>If the system uses comma for decimals (European locales), DAX commas must be semicolons</item>
/// </list>
/// <para><b>Example Problem (Fixed by this translator):</b></para>
/// <para>
/// On a system with culture en-DE (English with German regional settings):
/// </para>
/// <list type="bullet">
/// <item>Excel.International reports: ListSeparator=',' DecimalSeparator='.'</item>
/// <item>System culture (en-DE) has: NumberDecimalSeparator=','</item>
/// <item>Power Pivot uses SYSTEM culture, so it sees comma as decimal separator</item>
/// <item>Formula "DATEADD(Date[Date], -1, MONTH)" becomes "DATEADD(Date[Date], -1. MONTH)" (corrupted!)</item>
/// </list>
/// <para><b>Solution:</b></para>
/// <para>
/// This translator checks the SYSTEM culture's decimal separator. If it's a comma,
/// we translate all DAX function argument commas to semicolons, regardless of what
/// Excel's International property reports.
/// </para>
/// <para><b>Usage:</b></para>
/// <code>
/// var translator = new DaxFormulaTranslator(excelApp);
/// string daxFormula = translator.TranslateToLocale("CALCULATE([ACR], DATEADD(Date[Date], -1, MONTH))");
/// // Returns "CALCULATE([ACR]; DATEADD(Date[Date]; -1; MONTH))" on European systems
/// </code>
/// </remarks>
public sealed class DaxFormulaTranslator
{
    // XlApplicationInternational enum values (for reference, but we prefer system culture)
    private const int XlListSeparator = 5;
    private const int XlDecimalSeparator = 3;

    /// <summary>The Excel International list separator (for logging/diagnostics)</summary>
    public string ExcelListSeparator { get; }

    /// <summary>The Excel International decimal separator (for logging/diagnostics)</summary>
    public string ExcelDecimalSeparator { get; }

    /// <summary>The system culture's decimal separator - THIS is what Power Pivot uses</summary>
    public string SystemDecimalSeparator { get; }

    /// <summary>The system culture's list separator (for diagnostics)</summary>
    public string SystemListSeparator { get; }

    /// <summary>The separator to use for DAX function arguments.
    /// Uses semicolon if system decimal separator is comma (European locales),
    /// otherwise uses the Excel list separator.
    /// </summary>
    public string DaxArgumentSeparator { get; }

    /// <summary>True if system uses comma as decimal separator (meaning DAX commas must become semicolons)</summary>
    public bool SystemUsesCommaDecimal { get; }

    /// <summary>True if translation is needed (system uses comma for decimal, so DAX commas must be semicolons)</summary>
    public bool RequiresTranslation => SystemUsesCommaDecimal;

    /// <summary>
    /// True if the system has an invalid configuration where both decimal and list separator are the same.
    /// This is a Windows Regional Settings misconfiguration that will cause DAX errors.
    /// </summary>
    /// <remarks>
    /// This can happen with the en-DE locale (English with German regional settings) where
    /// Windows may default to comma for both decimal and list separator. The solution is to
    /// change the Windows list separator to semicolon in Regional Settings > Additional settings.
    /// </remarks>
    public bool HasSeparatorConflict => SystemDecimalSeparator == SystemListSeparator;

    /// <summary>
    /// Creates a new DaxFormulaTranslator by reading locale settings from both Excel Application
    /// and the system culture. Power Pivot uses the SYSTEM culture, not Excel's International settings.
    /// </summary>
    /// <param name="excelApp">The Excel.Application COM object (dynamic)</param>
    public DaxFormulaTranslator(dynamic excelApp)
    {
        // Read Excel's International settings (for logging/diagnostics)
        ExcelListSeparator = GetInternationalValue(excelApp, XlListSeparator) ?? ",";
        ExcelDecimalSeparator = GetInternationalValue(excelApp, XlDecimalSeparator) ?? ".";

        // CRITICAL: Power Pivot uses the SYSTEM culture, not Excel's International settings!
        // Get the system culture's decimal and list separators
        SystemDecimalSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
        SystemListSeparator = CultureInfo.CurrentCulture.TextInfo.ListSeparator;
        SystemUsesCommaDecimal = SystemDecimalSeparator == ",";

        // Determine the correct DAX argument separator:
        // - If SYSTEM uses comma for decimal, we MUST use semicolon (regardless of what Excel reports)
        // - Otherwise, use the Excel list separator (which is comma for US/UK locales)
        DaxArgumentSeparator = SystemUsesCommaDecimal ? ";" : ExcelListSeparator;
    }

    /// <summary>
    /// Translates a US (English) DAX formula to the locale-specific format Excel expects.
    /// Converts comma argument separators to the locale-specific list separator.
    /// </summary>
    /// <param name="usDaxFormula">US DAX formula with comma separators (e.g., "DATEADD(Date[Date], -1, MONTH)")</param>
    /// <returns>Locale-specific DAX formula (e.g., "DATEADD(Date[Date]; -1; MONTH)" on German Excel)</returns>
    /// <remarks>
    /// <para>Translation rules:</para>
    /// <list type="bullet">
    /// <item>Commas inside function parentheses are translated to locale list separator</item>
    /// <item>Content inside strings (single or double quotes) is preserved</item>
    /// <item>Content inside square brackets (column references like [Column Name]) is preserved</item>
    /// <item>Commas outside function calls are preserved (though rare in DAX)</item>
    /// </list>
    /// </remarks>
    public string TranslateToLocale(string usDaxFormula)
    {
        if (string.IsNullOrEmpty(usDaxFormula))
            return usDaxFormula;

        // If system uses period as decimal (US/UK locales), no translation needed
        // The comma in the formula is already the correct list separator
        if (!RequiresTranslation)
            return usDaxFormula;

        // Check if formula contains commas at all
        if (!usDaxFormula.Contains(','))
            return usDaxFormula;

        return TranslateFormula(usDaxFormula);
    }

    /// <summary>
    /// Translates DAX formula by converting comma separators to locale-specific separators.
    /// </summary>
    private string TranslateFormula(string formula)
    {
        var result = new StringBuilder(formula.Length);
        int parenDepth = 0;

        for (int i = 0; i < formula.Length; i++)
        {
            char c = formula[i];

            // Skip content in double quotes (string literals)
            if (c == '"')
            {
                int quoteEnd = FindClosingQuote(formula, i, '"');
                result.Append(formula.AsSpan(i, quoteEnd - i + 1));
                i = quoteEnd;
                continue;
            }

            // Skip content in single quotes (string literals in DAX)
            if (c == '\'')
            {
                int quoteEnd = FindClosingQuote(formula, i, '\'');
                result.Append(formula.AsSpan(i, quoteEnd - i + 1));
                i = quoteEnd;
                continue;
            }

            // Skip content in square brackets (column references like [Column Name, With Comma])
            if (c == '[')
            {
                int bracketEnd = FindClosingBracket(formula, i);
                result.Append(formula.AsSpan(i, bracketEnd - i + 1));
                i = bracketEnd;
                continue;
            }

            // Track parentheses depth
            if (c == '(')
            {
                parenDepth++;
                result.Append(c);
                continue;
            }

            if (c == ')')
            {
                parenDepth--;
                result.Append(c);
                continue;
            }

            // Translate commas inside function calls (parenDepth > 0)
            if (c == ',' && parenDepth > 0)
            {
                result.Append(DaxArgumentSeparator);
                continue;
            }

            // All other characters pass through unchanged
            result.Append(c);
        }

        return result.ToString();
    }

    /// <summary>
    /// Finds the closing quote character, handling escaped quotes.
    /// </summary>
    private static int FindClosingQuote(string formula, int startIndex, char quoteChar)
    {
        for (int i = startIndex + 1; i < formula.Length; i++)
        {
            if (formula[i] == quoteChar)
            {
                // Check for escaped quote (doubled quote character)
                if (i + 1 < formula.Length && formula[i + 1] == quoteChar)
                {
                    i++; // Skip the escaped quote
                    continue;
                }
                return i;
            }
        }
        // No closing quote found - return end of string
        return formula.Length - 1;
    }

    /// <summary>
    /// Finds the closing bracket ']', handling nested brackets.
    /// </summary>
    private static int FindClosingBracket(string formula, int startIndex)
    {
        int depth = 0;
        for (int i = startIndex; i < formula.Length; i++)
        {
            if (formula[i] == '[')
            {
                depth++;
            }
            else if (formula[i] == ']')
            {
                depth--;
                if (depth == 0)
                    return i;
            }
        }
        // No closing bracket found - return end of string
        return formula.Length - 1;
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
    /// Returns a summary of the locale settings for debugging/logging.
    /// </summary>
    public override string ToString()
    {
        var conflict = HasSeparatorConflict ? " [CONFLICT: decimal=list!]" : "";
        return $"DaxFormulaTranslator: SystemDecimal='{SystemDecimalSeparator}' SystemList='{SystemListSeparator}' ExcelList='{ExcelListSeparator}' ExcelDecimal='{ExcelDecimalSeparator}' DaxArgSeparator='{DaxArgumentSeparator}' RequiresTranslation={RequiresTranslation}{conflict}";
    }
}
