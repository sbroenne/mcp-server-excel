using System.ComponentModel;
using Microsoft.Extensions.AI;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// MCP prompts for Excel range operations - values, formulas, formatting, validation.
/// </summary>
[McpServerPromptType]
public static class ExcelRangePrompts
{
    /// <summary>
    /// Guide for formatting and styling Excel ranges.
    /// </summary>
    [McpServerPrompt(Name = "excel_range_formatting_guide")]
    [Description("Best practices for formatting and styling Excel ranges")]
    public static ChatMessage FormattingGuide()
    {
        return new ChatMessage(ChatRole.User, @"# EXCEL RANGE FORMATTING GUIDE

## FORMATTING CAPABILITIES

The **excel_range** tool with action ""format-range"" provides comprehensive visual formatting:

### FONT FORMATTING
- **fontName**: Font family (""Arial"", ""Calibri"", ""Times New Roman"")
- **fontSize**: Font size in points (8, 10, 11, 12, 14, 16, etc.)
- **bold**: Make text bold
- **italic**: Make text italic
- **underline**: Underline text
- **fontColor**: Font color (#RRGGBB hex, e.g., #FF0000 for red)

### FILL (BACKGROUND) FORMATTING
- **fillColor**: Background color (#RRGGBB hex, e.g., #FFFF00 for yellow)

### BORDER FORMATTING
- **borderStyle**: none, continuous, dash, dot, double (most common: continuous)
- **borderWeight**: hairline, thin, medium, thick (most common: thin)
- **borderColor**: Border color (#RRGGBB hex, e.g., #000000 for black)

### ALIGNMENT FORMATTING
- **horizontalAlignment**: left, center, right, justify, distributed
- **verticalAlignment**: top, center, bottom, justify, distributed
- **wrapText**: Enable text wrapping in cells (true/false)
- **orientation**: Text rotation in degrees (0-90 or -90)

## COMMON FORMATTING PATTERNS

### Pattern 1: Header Row Formatting
```javascript
excel_range(
  action: ""format-range"",
  sheetName: ""Sales"",
  rangeAddress: ""A1:E1"",
  bold: true,
  fontSize: 12,
  horizontalAlignment: ""center"",
  fillColor: ""#4472C4"",
  fontColor: ""#FFFFFF""
)
```

### Pattern 2: Highlight Important Values
```javascript
excel_range(
  action: ""format-range"",
  sheetName: ""Data"",
  rangeAddress: ""D2:D100"",
  fillColor: ""#FFFF00"",  // Yellow background
  bold: true
)
```

### Pattern 3: Add Borders to Table
```javascript
excel_range(
  action: ""format-range"",
  sheetName: ""Report"",
  rangeAddress: ""A1:E50"",
  borderStyle: ""continuous"",
  borderWeight: ""thin"",
  borderColor: ""#000000""
)
```

### Pattern 4: Right-Align Numbers
```javascript
excel_range(
  action: ""format-range"",
  sheetName: ""Financials"",
  rangeAddress: ""D2:D100"",
  horizontalAlignment: ""right""
)
```

### Pattern 5: Wrap Long Text
```javascript
excel_range(
  action: ""format-range"",
  sheetName: ""Comments"",
  rangeAddress: ""C2:C50"",
  wrapText: true,
  verticalAlignment: ""top""
)
```

## COLOR REFERENCES

### Common Colors (hex codes):
- Red: #FF0000
- Green: #00FF00
- Blue: #0000FF
- Yellow: #FFFF00
- Orange: #FFA500
- Purple: #800080
- Black: #000000
- White: #FFFFFF
- Light Gray: #D3D3D3
- Dark Gray: #808080

### Excel Theme Colors:
- Blue: #4472C4 (default Excel header blue)
- Orange: #ED7D31
- Gray: #A5A5A5
- Yellow: #FFC000
- Light Blue: #5B9BD5
- Green: #70AD47

## BEST PRACTICES

1. **Combine Related Formatting**: Apply multiple properties in one call for efficiency
2. **Use Consistent Colors**: Stick to a theme for professional appearance
3. **Headers Stand Out**: Bold + larger font + background color
4. **Alignment Matters**: Numbers → right, Text → left, Headers → center
5. **Borders for Tables**: Use thin continuous borders for grid appearance
6. **Highlight Key Data**: Use yellow/orange background for important cells
7. **Wrap Long Text**: Enable wrapText for comment/description columns

## FORMATTING + NUMBER FORMATS

Combine visual formatting with number formatting for complete styling:

```javascript
// Step 1: Format numbers as currency
excel_range(
  action: ""set-number-format"",
  sheetName: ""Sales"",
  rangeAddress: ""D2:D100"",
  formatCode: ""$#,##0.00""
)

// Step 2: Apply visual formatting
excel_range(
  action: ""format-range"",
  sheetName: ""Sales"",
  rangeAddress: ""D2:D100"",
  horizontalAlignment: ""right"",
  bold: true
)
```

## ANTI-PATTERNS (AVOID)

❌ Applying formatting cell-by-cell in a loop
  → Format entire range in single call

❌ Using color indexes instead of hex codes
  → Use #RRGGBB hex codes for consistency

❌ Over-formatting (too many colors/styles)
  → Keep it simple and professional

❌ Inconsistent alignment
  → Be consistent: numbers right, text left
");
    }

    /// <summary>
    /// Guide for data validation in Excel ranges.
    /// </summary>
    [McpServerPrompt(Name = "excel_range_validation_guide")]
    [Description("Best practices for data validation rules in Excel")]
    public static ChatMessage ValidationGuide()
    {
        return new ChatMessage(ChatRole.User, @"# EXCEL RANGE VALIDATION GUIDE

## VALIDATION TYPES

The **excel_range** tool with action ""validate-range"" provides comprehensive data validation:

### LIST VALIDATION (Dropdown)
- **validationType**: ""list""
- **validationFormula1**: Comma-separated values or range reference
- **showDropdown**: true (shows dropdown arrow)

Example dropdown values:
```javascript
excel_range(
  action: ""validate-range"",
  sheetName: ""Data"",
  rangeAddress: ""F2:F100"",
  validationType: ""list"",
  validationFormula1: ""Active,Inactive,Pending"",
  showDropdown: true
)
```

Example dropdown from range:
```javascript
excel_range(
  action: ""validate-range"",
  sheetName: ""Data"",
  rangeAddress: ""F2:F100"",
  validationType: ""list"",
  validationFormula1: ""=$A$1:$A$10"",  // Reference to list in A1:A10
  showDropdown: true
)
```

### WHOLE NUMBER VALIDATION
- **validationType**: ""whole""
- **validationOperator**: between, notBetween, equal, notEqual, greaterThan, lessThan, greaterThanOrEqual, lessThanOrEqual
- **validationFormula1**: First value
- **validationFormula2**: Second value (for between/notBetween)

Example number range:
```javascript
excel_range(
  action: ""validate-range"",
  sheetName: ""Data"",
  rangeAddress: ""E2:E100"",
  validationType: ""whole"",
  validationOperator: ""between"",
  validationFormula1: ""1"",
  validationFormula2: ""999"",
  errorStyle: ""stop"",
  errorTitle: ""Invalid Value"",
  errorMessage: ""Enter a number between 1 and 999""
)
```

### DECIMAL VALIDATION
- **validationType**: ""decimal""
- Same operators as whole number
- Allows decimal places

Example minimum value:
```javascript
excel_range(
  action: ""validate-range"",
  sheetName: ""Prices"",
  rangeAddress: ""D2:D100"",
  validationType: ""decimal"",
  validationOperator: ""greaterThan"",
  validationFormula1: ""0"",
  errorStyle: ""stop"",
  errorMessage: ""Price must be greater than 0""
)
```

### DATE VALIDATION
- **validationType**: ""date""
- **validationOperator**: Same as numeric
- **validationFormula1/2**: Date values (e.g., ""1/1/2025"")

Example minimum date:
```javascript
excel_range(
  action: ""validate-range"",
  sheetName: ""Events"",
  rangeAddress: ""A2:A100"",
  validationType: ""date"",
  validationOperator: ""greaterThanOrEqual"",
  validationFormula1: ""1/1/2025"",
  errorStyle: ""warning"",
  errorMessage: ""Date should be in 2025 or later""
)
```

### TIME VALIDATION
- **validationType**: ""time""
- Same operators as date
- Use time format (e.g., ""9:00 AM"")

### TEXT LENGTH VALIDATION
- **validationType**: ""textLength""
- **validationOperator**: Same as numeric
- **validationFormula1**: Character count limit

Example maximum length:
```javascript
excel_range(
  action: ""validate-range"",
  sheetName: ""Comments"",
  rangeAddress: ""C2:C100"",
  validationType: ""textLength"",
  validationOperator: ""lessThanOrEqual"",
  validationFormula1: ""100"",
  errorStyle: ""stop"",
  errorMessage: ""Comment cannot exceed 100 characters""
)
```

### CUSTOM FORMULA VALIDATION
- **validationType**: ""custom""
- **validationFormula1**: Excel formula returning TRUE/FALSE

Example custom validation:
```javascript
excel_range(
  action: ""validate-range"",
  sheetName: ""Data"",
  rangeAddress: ""D2:D100"",
  validationType: ""custom"",
  validationFormula1: ""=MOD(D2,5)=0"",  // Must be multiple of 5
  errorStyle: ""stop"",
  errorMessage: ""Value must be a multiple of 5""
)
```

## VALIDATION OPERATORS

| Operator | Usage | Formula1 | Formula2 |
|----------|-------|----------|----------|
| between | Value is between two numbers | Minimum | Maximum |
| notBetween | Value is outside range | Lower bound | Upper bound |
| equal | Value equals specific value | Value | - |
| notEqual | Value doesn't equal | Value | - |
| greaterThan | Value > minimum | Minimum | - |
| lessThan | Value < maximum | Maximum | - |
| greaterThanOrEqual | Value >= minimum | Minimum | - |
| lessThanOrEqual | Value <= maximum | Maximum | - |

## VALIDATION MESSAGES

### Input Message (Help)
Shows when cell is selected:
- **showInputMessage**: true
- **inputTitle**: Short title (e.g., ""Valid Values"")
- **inputMessage**: Helpful description

### Error Alert (Validation Failed)
Shows when invalid value entered:
- **showErrorAlert**: true
- **errorStyle**: ""stop"" (block), ""warning"" (allow), ""information"" (notify)
- **errorTitle**: Error title
- **errorMessage**: Detailed error description

Example with messages:
```javascript
excel_range(
  action: ""validate-range"",
  sheetName: ""Data"",
  rangeAddress: ""F2:F100"",
  validationType: ""list"",
  validationFormula1: ""Active,Inactive,Pending"",
  showDropdown: true,
  showInputMessage: true,
  inputTitle: ""Status Selection"",
  inputMessage: ""Choose Active, Inactive, or Pending"",
  showErrorAlert: true,
  errorStyle: ""stop"",
  errorTitle: ""Invalid Status"",
  errorMessage: ""Please select a valid status from the dropdown""
)
```

## VALIDATION OPTIONS

- **ignoreBlank**: Allow empty cells (default: true)
- **showDropdown**: Show dropdown arrow for list validation (default: true)

## COMMON VALIDATION PATTERNS

### Pattern 1: Status Dropdown
```javascript
excel_range(
  action: ""validate-range"",
  sheetName: ""Tasks"",
  rangeAddress: ""E2:E100"",
  validationType: ""list"",
  validationFormula1: ""Not Started,In Progress,Complete,Blocked"",
  showDropdown: true
)
```

### Pattern 2: Positive Numbers Only
```javascript
excel_range(
  action: ""validate-range"",
  sheetName: ""Sales"",
  rangeAddress: ""D2:D100"",
  validationType: ""decimal"",
  validationOperator: ""greaterThan"",
  validationFormula1: ""0"",
  errorStyle: ""stop"",
  errorMessage: ""Amount must be greater than 0""
)
```

### Pattern 3: Date Range
```javascript
excel_range(
  action: ""validate-range"",
  sheetName: ""Schedule"",
  rangeAddress: ""A2:A100"",
  validationType: ""date"",
  validationOperator: ""between"",
  validationFormula1: ""1/1/2025"",
  validationFormula2: ""12/31/2025"",
  errorStyle: ""warning"",
  errorMessage: ""Date should be within 2025""
)
```

### Pattern 4: Email Format (Custom)
```javascript
excel_range(
  action: ""validate-range"",
  sheetName: ""Contacts"",
  rangeAddress: ""C2:C100"",
  validationType: ""custom"",
  validationFormula1: ""=AND(ISNUMBER(FIND(""@"",C2)),ISNUMBER(FIND(""."",C2)))"",
  errorStyle: ""warning"",
  errorMessage: ""Invalid email format""
)
```

### Pattern 5: Unique Values (Custom)
```javascript
excel_range(
  action: ""validate-range"",
  sheetName: ""IDs"",
  rangeAddress: ""A2:A100"",
  validationType: ""custom"",
  validationFormula1: ""=COUNTIF($A$2:$A$100,A2)=1"",
  errorStyle: ""stop"",
  errorMessage: ""ID must be unique""
)
```

## BEST PRACTICES

1. **Use Input Messages**: Help users understand what values are expected
2. **Choose Right Error Style**: 
   - stop: Prevent invalid data (strict validation)
   - warning: Allow but warn (flexible validation)
   - information: Just notify (informational)
3. **Test Validation**: Verify formulas work before applying to large ranges
4. **List Validation First**: Dropdowns are most user-friendly
5. **Ignore Blank**: Usually want `ignoreBlank: true` to allow empty cells
6. **Clear Error Messages**: Explain what's wrong and how to fix it
7. **Combine with Formatting**: Format validated cells consistently

## ANTI-PATTERNS (AVOID)

❌ Validating without error messages
  → Always provide helpful error message

❌ Using ""stop"" for everything
  → Use ""warning"" when flexibility needed

❌ Complex custom formulas without testing
  → Test formulas on single cell first

❌ Hardcoded values in formulas
  → Reference named ranges for maintainability

❌ Validation without input message
  → Help users know what's expected
");
    }

    /// <summary>
    /// Guide for comprehensive range workflows combining values, formulas, formatting, and validation.
    /// </summary>
    [McpServerPrompt(Name = "excel_range_complete_workflow")]
    [Description("Complete workflow patterns for range operations")]
    public static ChatMessage CompleteWorkflowGuide()
    {
        return new ChatMessage(ChatRole.User, @"# EXCEL RANGE COMPLETE WORKFLOW GUIDE

## COMPLETE WORKFLOW PATTERNS

Combine multiple range operations for professional, validated, formatted data:

### Workflow 1: Create Formatted Data Entry Table

```javascript
// Step 1: Set up headers
excel_range(
  action: ""set-values"",
  sheetName: ""Employees"",
  rangeAddress: ""A1:E1"",
  values: [[""ID"", ""Name"", ""Department"", ""Salary"", ""Status""]]
)

// Step 2: Format headers
excel_range(
  action: ""format-range"",
  sheetName: ""Employees"",
  rangeAddress: ""A1:E1"",
  bold: true,
  fontSize: 12,
  horizontalAlignment: ""center"",
  fillColor: ""#4472C4"",
  fontColor: ""#FFFFFF""
)

// Step 3: Format salary column as currency
excel_range(
  action: ""set-number-format"",
  sheetName: ""Employees"",
  rangeAddress: ""D2:D100"",
  formatCode: ""$#,##0.00""
)

// Step 4: Add department dropdown
excel_range(
  action: ""validate-range"",
  sheetName: ""Employees"",
  rangeAddress: ""C2:C100"",
  validationType: ""list"",
  validationFormula1: ""Sales,Marketing,Engineering,HR,Finance"",
  showDropdown: true
)

// Step 5: Add status dropdown
excel_range(
  action: ""validate-range"",
  sheetName: ""Employees"",
  rangeAddress: ""E2:E100"",
  validationType: ""list"",
  validationFormula1: ""Active,On Leave,Inactive"",
  showDropdown: true
)

// Step 6: Add borders
excel_range(
  action: ""format-range"",
  sheetName: ""Employees"",
  rangeAddress: ""A1:E100"",
  borderStyle: ""continuous"",
  borderWeight: ""thin"",
  borderColor: ""#000000""
)
```

### Workflow 2: Build Financial Report with Formulas

```javascript
// Step 1: Set up headers
excel_range(
  action: ""set-values"",
  sheetName: ""Report"",
  rangeAddress: ""A1:D1"",
  values: [[""Month"", ""Revenue"", ""Expenses"", ""Profit""]]
)

// Step 2: Format headers
excel_range(
  action: ""format-range"",
  sheetName: ""Report"",
  rangeAddress: ""A1:D1"",
  bold: true,
  fillColor: ""#70AD47"",
  fontColor: ""#FFFFFF""
)

// Step 3: Add profit formulas
excel_range(
  action: ""set-formulas"",
  sheetName: ""Report"",
  rangeAddress: ""D2:D13"",
  formulas: [[""=B2-C2""], [""=B3-C3""], [""=B4-C4""], ..., [""=B13-C13""]]
)

// Step 4: Format currency columns
excel_range(
  action: ""set-number-format"",
  sheetName: ""Report"",
  rangeAddress: ""B2:D13"",
  formatCode: ""$#,##0""
)

// Step 5: Right-align numbers
excel_range(
  action: ""format-range"",
  sheetName: ""Report"",
  rangeAddress: ""B2:D13"",
  horizontalAlignment: ""right""
)

// Step 6: Highlight negative profits in red
excel_range(
  action: ""format-range"",
  sheetName: ""Report"",
  rangeAddress: ""D2:D13"",
  fontColor: ""#FF0000""  // Apply conditionally based on values
)
```

### Workflow 3: Data Validation with Error Prevention

```javascript
// Step 1: Create ID column with unique validation
excel_range(
  action: ""validate-range"",
  sheetName: ""Records"",
  rangeAddress: ""A2:A100"",
  validationType: ""custom"",
  validationFormula1: ""=COUNTIF($A$2:$A$100,A2)=1"",
  errorStyle: ""stop"",
  errorTitle: ""Duplicate ID"",
  errorMessage: ""Each ID must be unique""
)

// Step 2: Add date validation (future dates only)
excel_range(
  action: ""validate-range"",
  sheetName: ""Records"",
  rangeAddress: ""B2:B100"",
  validationType: ""date"",
  validationOperator: ""greaterThanOrEqual"",
  validationFormula1: ""=TODAY()"",
  errorStyle: ""warning"",
  errorMessage: ""Date should be today or in the future""
)

// Step 3: Validate amount is positive
excel_range(
  action: ""validate-range"",
  sheetName: ""Records"",
  rangeAddress: ""C2:C100"",
  validationType: ""decimal"",
  validationOperator: ""greaterThan"",
  validationFormula1: ""0"",
  errorStyle: ""stop"",
  errorMessage: ""Amount must be greater than 0""
)

// Step 4: Format amounts as currency
excel_range(
  action: ""set-number-format"",
  sheetName: ""Records"",
  rangeAddress: ""C2:C100"",
  formatCode: ""$#,##0.00""
)
```

### Workflow 4: Build Dashboard with Batch Mode

Use batch mode for efficiency when creating complex layouts:

```javascript
// Step 1: Start batch session
begin_excel_batch(excelPath: ""Dashboard.xlsx"")

// Step 2-10: All range operations with batchId parameter
excel_range(action: ""set-values"", ..., batchId: ""<batch-id>"")
excel_range(action: ""format-range"", ..., batchId: ""<batch-id>"")
excel_range(action: ""set-formulas"", ..., batchId: ""<batch-id>"")
// ... more operations ...

// Step 11: Commit all changes
commit_excel_batch(batchId: ""<batch-id>"", action: ""save"")
```

## OPERATION ORDER BEST PRACTICES

1. **Data First**: Set values/formulas before formatting
2. **Number Formats**: Apply before visual formatting
3. **Validation Last**: Add validation after data structure is set
4. **Format Once**: Combine multiple formatting properties in single call
5. **Use Batch Mode**: For 3+ operations on same file

## COMBINING WITH OTHER TOOLS

### Range + Table
```javascript
// Step 1: Set up data with range operations
excel_range(action: ""set-values"", ...)
excel_range(action: ""format-range"", ...)
excel_range(action: ""validate-range"", ...)

// Step 2: Convert to Excel Table for advanced features
excel_table(action: ""create"", sourceRange: ""A1:E100"")
```

### Range + Power Query
```javascript
// Step 1: Import data with Power Query
excel_powerquery(action: ""import"", loadDestination: ""worksheet"")

// Step 2: Format loaded data
excel_range(action: ""format-range"", ...)
excel_range(action: ""set-number-format"", ...)
```

### Range + Parameters
```javascript
// Step 1: Create named range parameters
excel_namedrange(action: ""create"", name: ""StartDate"", reference: ""=Settings!A1"")

// Step 2: Use parameters in validation
excel_range(
  action: ""validate-range"",
  validationType: ""date"",
  validationFormula1: ""=StartDate"",  // Reference to parameter
  ...
)
```

## KEY INSIGHTS

1. **Single Cell = 1x1 Range**: ""A1"" is the same as ""A1:A1""
2. **Named Ranges**: Leave sheetName empty when using named ranges
3. **Batch Mode**: Use for 3+ operations to reduce Excel open/close cycles
4. **Format After Data**: Set values/formulas first, then format
5. **Validation Messages**: Always include helpful error messages
6. **Color Consistency**: Use theme colors (#4472C4, #70AD47, etc.)
7. **Number Formats**: Apply before visual formatting for best results
");
    }
}
