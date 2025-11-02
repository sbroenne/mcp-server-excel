using System.ComponentModel;
using Microsoft.Extensions.AI;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// MCP prompts for Excel naming conventions and formatting best practices.
/// Teaches LLMs community-standard naming patterns and professional formatting for Excel objects.
/// </summary>
[McpServerPromptType]
public static class ExcelNamingPrompts
{
    private static readonly string _namingGuide;
    private static readonly string _formattingGuide;

    static ExcelNamingPrompts()
    {
        // Load from markdown files
        _namingGuide = MarkdownLoader.LoadPrompt("excel_naming_conventions.md");
        _formattingGuide = MarkdownLoader.LoadPrompt("excel_formatting_best_practices.md");
    }

    [McpServerPrompt(Name = "excel_naming_best_practices")]
    [Description("Excel naming conventions - community best practices for tables, ranges, queries, worksheets")]
    public static ChatMessage NamingBestPractices()
    {
        return new ChatMessage(ChatRole.User, _namingGuide);
    }

    [McpServerPrompt(Name = "excel_formatting_best_practices")]
    [Description("Excel formatting standards - fonts, colors, number formats, table styles, accessibility")]
    public static ChatMessage FormattingBestPractices()
    {
        return new ChatMessage(ChatRole.User, _formattingGuide);
    }

    [McpServerPrompt(Name = "excel_suggest_names")]
    [Description("Get naming suggestions for Excel objects based on context")]
    public static ChatMessage SuggestNames(
        [Description("Object type: table, range, query, worksheet, column")]
        string objectType,
        [Description("Purpose or content description")]
        string purpose)
    {
        var suggestions = GenerateNameSuggestions(objectType.ToLowerInvariant(), purpose);

        return new ChatMessage(ChatRole.User, $@"
# NAME SUGGESTIONS FOR {objectType.ToUpperInvariant()}

**Purpose:** {purpose}

**Recommended Names (PascalCase):**
{suggestions.PascalCase}

**Alternative (snake_case):**
{suggestions.SnakeCase}

**With Prefix (for large workbooks):**
{suggestions.WithPrefix}

**Best Practice:**
- Use PascalCase by default: {suggestions.Recommended}
- Be descriptive and consistent
- Avoid generic names like 'Data' or 'Table1'

**Validation:**
- Max {suggestions.MaxLength} characters
- {suggestions.Rules}
");
    }

    [McpServerPrompt(Name = "excel_suggest_formatting")]
    [Description("Get formatting suggestions for Excel ranges and tables based on purpose")]
    public static ChatMessage SuggestFormatting(
        [Description("Content type: financial, sales, dashboard, report, data-entry")]
        string contentType,
        [Description("Has header row?")]
        bool hasHeaders = true,
        [Description("Has totals row?")]
        bool hasTotals = false)
    {
        var suggestions = GenerateFormattingSuggestions(contentType.ToLowerInvariant(), hasHeaders, hasTotals);

        return new ChatMessage(ChatRole.User, $@"
# FORMATTING SUGGESTIONS FOR {contentType.ToUpperInvariant()}

**Recommended Table Style:** {suggestions.TableStyle}

**Font:**
- Family: {suggestions.FontFamily}
- Size: {suggestions.FontSize}pt
- Headers: Bold

**Colors:**
- Header Background: {suggestions.HeaderColor}
- Header Text: {suggestions.HeaderTextColor}
{(hasTotals ? $"- Totals Background: {suggestions.TotalsColor}" : "")}

**Number Formats:**
{suggestions.NumberFormats}

**Layout:**
- Freeze panes: After row 1 (headers)
- Column alignment: {suggestions.Alignment}
- Row height: {suggestions.RowHeight}pt
{(hasTotals ? "- Totals: Bold with top border (medium)" : "")}

**Quick Apply:**
1. Select table range
2. Format as Table → {suggestions.TableStyle}
3. Apply number formats to numeric columns
{(hasHeaders ? "4. Freeze panes after header row" : "")}
{(hasTotals ? "5. Bold totals row, add top border" : "")}
");
    }

    private static (string PascalCase, string SnakeCase, string WithPrefix, string Recommended, int MaxLength, string Rules) 
        GenerateNameSuggestions(string objectType, string purpose)
    {
        // Generate smart suggestions based on purpose
        var baseNamePascal = ToPascalCase(purpose);
        var baseNameSnake = ToSnakeCase(purpose);

        return objectType switch
        {
            "table" => (
                PascalCase: $"- {baseNamePascal}\n- {baseNamePascal}Data\n- {baseNamePascal}List",
                SnakeCase: $"- {baseNameSnake}\n- {baseNameSnake}_data\n- {baseNameSnake}_list",
                WithPrefix: $"- tbl_{baseNamePascal}\n- tbl_{baseNameSnake}",
                Recommended: baseNamePascal,
                MaxLength: 255,
                Rules: "Must start with letter or underscore, alphanumeric + underscore only"
            ),
            "range" or "parameter" => (
                PascalCase: $"- {baseNamePascal.ToUpperInvariant()}\n- {baseNamePascal.ToUpperInvariant()}_VALUE",
                SnakeCase: $"- {baseNameSnake.ToUpperInvariant()}\n- {baseNameSnake.ToUpperInvariant()}_VALUE",
                WithPrefix: $"- prm_{baseNamePascal}\n- rng_{baseNamePascal}",
                Recommended: baseNamePascal.ToUpperInvariant(),
                MaxLength: 255,
                Rules: "Must start with letter/underscore/backslash, no spaces, cannot look like cell reference"
            ),
            "query" or "powerquery" => (
                PascalCase: $"- Transform_{baseNamePascal}\n- Load_{baseNamePascal}\n- {baseNamePascal}_Extract",
                SnakeCase: $"- transform_{baseNameSnake}\n- load_{baseNameSnake}\n- {baseNameSnake}_extract",
                WithPrefix: $"- pq_{baseNamePascal}\n- pq_{baseNameSnake}",
                Recommended: $"Transform_{baseNamePascal}",
                MaxLength: 80,
                Rules: "Spaces allowed but not recommended, avoid special characters"
            ),
            "worksheet" or "sheet" => (
                PascalCase: $"- {baseNamePascal}\n- {baseNamePascal} Data\n- 01_{baseNamePascal}",
                SnakeCase: $"- {baseNameSnake}\n- {baseNameSnake}_data\n- 01_{baseNameSnake}",
                WithPrefix: $"- _{baseNamePascal} (underscore to hide)\n- 01_{baseNamePascal} (sort order)",
                Recommended: $"{baseNamePascal} Data",
                MaxLength: 31,
                Rules: "Spaces OK! Cannot contain: [ ] * ? / \\ :"
            ),
            "column" => (
                PascalCase: $"- {baseNamePascal}\n- {baseNamePascal}Value\n- Is{baseNamePascal}",
                SnakeCase: $"- {baseNameSnake}\n- {baseNameSnake}_value\n- is_{baseNameSnake}",
                WithPrefix: $"- {baseNamePascal}USD (with unit)\n- {baseNamePascal}Date (with type)",
                Recommended: baseNamePascal,
                MaxLength: 255,
                Rules: "Clear and concise, include units/types if applicable"
            ),
            _ => (
                PascalCase: $"- {baseNamePascal}",
                SnakeCase: $"- {baseNameSnake}",
                WithPrefix: $"- obj_{baseNamePascal}",
                Recommended: baseNamePascal,
                MaxLength: 255,
                Rules: "Follow naming conventions for the object type"
            )
        };
    }

    private static (string TableStyle, string FontFamily, int FontSize, string HeaderColor, 
                    string HeaderTextColor, string TotalsColor, string NumberFormats, 
                    string Alignment, int RowHeight) 
        GenerateFormattingSuggestions(string contentType, bool hasHeaders, bool hasTotals)
    {
        return contentType switch
        {
            "financial" or "finance" => (
                TableStyle: "Table Style Light 9 (minimal borders)",
                FontFamily: "Aptos or Calibri",
                FontSize: 11,
                HeaderColor: "#D6DCE4 (Light Blue)",
                HeaderTextColor: "Black",
                TotalsColor: "#D6DCE4 (Light Blue)",
                NumberFormats: "- Currency: $#,##0 (no decimals)\n- Percentage: 0.0%\n- Variance: $#,##0;[Red]($#,##0)",
                Alignment: "Numbers right, text left",
                RowHeight: 15
            ),
            "sales" => (
                TableStyle: "Table Style Medium 2 (blue header)",
                FontFamily: "Aptos or Calibri",
                FontSize: 11,
                HeaderColor: "#4472C4 (Theme Blue)",
                HeaderTextColor: "White",
                TotalsColor: "#2F5496 (Dark Blue)",
                NumberFormats: "- Units: #,##0\n- Price: $#,##0.00\n- Total: $#,##0.00\n- Discount: 0%",
                Alignment: "Numbers right, text left",
                RowHeight: 15
            ),
            "dashboard" => (
                TableStyle: "Table Style Medium 7 (colorful)",
                FontFamily: "Aptos or Calibri",
                FontSize: 12,
                HeaderColor: "#4472C4 (Theme Blue)",
                HeaderTextColor: "White",
                TotalsColor: "#ED7D31 (Orange)",
                NumberFormats: "- Metrics: #,##0\n- Percentages: 0%\n- Use data bars for visual comparison",
                Alignment: "Numbers center, text left",
                RowHeight: 18
            ),
            "report" => (
                TableStyle: "Table Style Medium 2 (professional)",
                FontFamily: "Aptos or Calibri",
                FontSize: 11,
                HeaderColor: "#4472C4 (Theme Blue)",
                HeaderTextColor: "White",
                TotalsColor: "#2F5496 (Dark Blue)",
                NumberFormats: "- Dates: mmm d, yyyy\n- Numbers: #,##0\n- Currency: $#,##0.00",
                Alignment: "Numbers right, text left, headers center",
                RowHeight: 15
            ),
            "data-entry" or "form" => (
                TableStyle: "Table Style Light 1 (minimal)",
                FontFamily: "Aptos or Calibri",
                FontSize: 11,
                HeaderColor: "#D6DCE4 (Light Blue)",
                HeaderTextColor: "Black",
                TotalsColor: "#D6DCE4 (Light Blue)",
                NumberFormats: "- Use data validation dropdowns\n- Date: m/d/yyyy\n- Required fields: Light yellow (#FFF2CC) background",
                Alignment: "Text left, dates right",
                RowHeight: 18
            ),
            _ => (
                TableStyle: "Table Style Medium 2 (default blue)",
                FontFamily: "Aptos or Calibri",
                FontSize: 11,
                HeaderColor: "#4472C4 (Theme Blue)",
                HeaderTextColor: "White",
                TotalsColor: "#2F5496 (Dark Blue)",
                NumberFormats: "- Currency: $#,##0.00\n- Percentage: 0%\n- Number: #,##0",
                Alignment: "Numbers right, text left",
                RowHeight: 15
            )
        };
    }

    private static string ToPascalCase(string input)
    {
        if (string.IsNullOrWhiteSpace(input)) return "MyObject";

        // Simple conversion: capitalize first letter of each word, remove spaces
        var words = input.Split(new[] { ' ', '_', '-' }, StringSplitOptions.RemoveEmptyEntries);
        return string.Join("", words.Select(w => 
            char.ToUpperInvariant(w[0]) + w.Substring(1).ToLowerInvariant()));
    }

    private static string ToSnakeCase(string input)
    {
        if (string.IsNullOrWhiteSpace(input)) return "my_object";

        // Simple conversion: lowercase with underscores
        var words = input.Split(new[] { ' ', '_', '-' }, StringSplitOptions.RemoveEmptyEntries);
        return string.Join("_", words.Select(w => w.ToLowerInvariant()));
    }
}
