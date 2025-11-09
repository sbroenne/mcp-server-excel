# Excel Formatting & Data Validation Specification

> **Comprehensive formatting and validation capabilities for ranges, formulas, and tables**
> 
> **🤖 Primary Audience:** LLMs using MCP Server tools for professional Excel automation

## What This Spec Provides (For LLMs)

This specification consolidates and extends formatting and data validation across all Excel operations:

### **Number Formatting** - Professional Data Display
- **Currency, percentages, dates, custom formats** - Make data readable
- **Get/Set operations** - Read existing formats or apply new ones
- **Bulk operations** - Format entire ranges or table columns efficiently
- **Common use:** Format sales reports, financial dashboards, date columns, percentages

### **Visual Formatting** - Professional Appearance
- **Fonts** - Family, size, bold, italic, color, underline, strikethrough
- **Cell appearance** - Background colors, borders, patterns
- **Alignment** - Horizontal, vertical, text wrapping, indentation, rotation
- **Common use:** Headers, highlighting, color-coding, readability improvements

### **Data Validation** - Data Integrity
- **Dropdown lists** - Restrict input to predefined values
- **Number ranges** - Min/max validation (whole numbers, decimals, dates, times)
- **Text length** - Character limits
- **Custom formulas** - Complex validation rules
- **Error alerts** - Custom messages for invalid data
- **Common use:** Data entry forms, quality control, preventing errors

### **Why You Need These Tools**
When users ask you to "format the sales report" or "add dropdowns for status," you'll use these commands to:
1. Create professional-looking spreadsheets with proper number formats
2. Apply visual styling (colors, fonts, borders) for readability
3. Ensure data quality with validation rules
4. Build user-friendly data entry interfaces

---

## Current State Analysis

### ✅ **What EXISTS Today**

**Range Operations (Phase 1 - Implemented):**
- ✅ Get/Set values and formulas
- ✅ Clear operations (all, contents, formats)
- ✅ Copy operations (all, values, formulas)
- ✅ Insert/delete cells/rows/columns
- ✅ Find/replace
- ✅ Sort
- ✅ Hyperlinks
- ✅ UsedRange, CurrentRegion, RangeInfo

**Table Operations (Implemented):**
- ✅ Lifecycle (create, rename, delete, info, resize)
- ✅ Style management (SetStyleAsync)
- ✅ Totals row management
- ✅ Filtering (apply, clear, get state)
- ✅ Column operations (add, remove, rename)
- ✅ Sorting (single and multi-column)
- ✅ Structured references
- ✅ Append rows

**PivotTable Operations (Phase 1 - Implemented):**
- ✅ Lifecycle (create, delete, list, info)
- ✅ Field management (add/remove/move fields in all areas)
- ✅ Field formatting (SetFieldFormatAsync for number formats)
- ✅ Layout management (SetLayoutAsync - Compact, Outline, Tabular)
- ✅ Style management (SetStyleAsync - 28 built-in styles)
- ✅ Data analysis (refresh, filter, sort)

### ❌ **What's MISSING Today**

**Number Formatting:**
- ❌ Get number formats from ranges
- ❌ Set number formats (uniform or cell-by-cell)
- ❌ Common format presets (currency, percentage, date patterns)
- ❌ Table column number formatting

**Visual Formatting:**
- ❌ Font properties (name, size, bold, italic, color, underline, strikethrough)
- ❌ Cell background colors
- ❌ Borders (styles, weights, colors)
- ❌ Alignment (horizontal, vertical, wrap, indent, rotation)
- ❌ Row height / column width
- ❌ Auto-fit columns/rows

**Data Validation:**
- ❌ Add validation rules to ranges
- ❌ Get validation settings from cells
- ❌ Remove validation
- ❌ All validation types (list, whole, decimal, date, time, text-length, custom)
- ❌ Error alerts and input messages
- ❌ Table column validation

**Advanced:**
- ❌ Conditional formatting
- ❌ Cell merge/unmerge
- ❌ Cell locking for protection

**Note:** PivotTable formatting (field formats, layouts, styles) is **fully implemented** and functional.

---

## Target Architecture

### Design Principles

1. **LLM-First Design** - Optimized for AI automation workflows
2. **Breaking Changes Acceptable** - Clean API > backwards compatibility
3. **Unified Approach** - Consistent patterns across ranges and tables
4. **Excel COM Native** - Use native Excel capabilities, no custom processing
5. **Performance Optimized** - Batch operations where possible

### Proposed Command Structure

```
Formatting & Validation Commands:
├── RangeCommands (extends existing)
│   ├── Number Formatting (3 methods)
│   ├── Visual Formatting (8 methods)
│   └── Data Validation (4 methods)
│
└── TableCommands (extends existing)
    ├── Number Formatting (2 methods)
    ├── Visual Formatting (4 methods - via ranges)
    └── Data Validation (2 methods)
```

**Philosophy:** 
- **RangeCommands** = Low-level, works anywhere
- **TableCommands** = High-level, table-specific convenience methods that delegate to RangeCommands

---

## Proposed API Design

### 1. Range Number Formatting

```csharp
public interface IRangeCommands
{
    // === NUMBER FORMAT OPERATIONS ===
    
    /// <summary>
    /// Gets number format codes from range (2D array matching range dimensions)
    /// Excel COM: Range.NumberFormat
    /// </summary>
    /// <returns>2D array of format codes (e.g., [["$#,##0.00", "0.00%"], ["m/d/yyyy", "General"]])</returns>
    Task<RangeNumberFormatResult> GetNumberFormatsAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    /// <summary>
    /// Sets uniform number format for entire range
    /// Excel COM: Range.NumberFormat = formatCode
    /// </summary>
    /// <param name="formatCode">
    /// Excel format code (e.g., "$#,##0.00", "0.00%", "m/d/yyyy", "General", "@")
    /// See NumberFormatPresets class for common patterns
    /// </param>
    Task<OperationResult> SetNumberFormatAsync(IExcelBatch batch, string sheetName, string rangeAddress, string formatCode);
    
    /// <summary>
    /// Sets number formats cell-by-cell from 2D array
    /// Excel COM: Range.NumberFormat (per cell)
    /// </summary>
    /// <param name="formats">2D array of format codes matching range dimensions</param>
    Task<OperationResult> SetNumberFormatsAsync(IExcelBatch batch, string sheetName, string rangeAddress, List<List<string>> formats);
}

/// <summary>
/// Common Excel number format codes for LLM convenience
/// </summary>
public static class NumberFormatPresets
{
    // Currency
    public const string Currency = "$#,##0.00";
    public const string CurrencyNoDecimals = "$#,##0";
    public const string CurrencyNegativeRed = "$#,##0.00_);[Red]($#,##0.00)";
    
    // Percentages
    public const string Percentage = "0.00%";
    public const string PercentageNoDecimals = "0%";
    public const string PercentageOneDecimal = "0.0%";
    
    // Dates
    public const string DateShort = "m/d/yyyy";
    public const string DateLong = "mmmm d, yyyy";
    public const string DateMonthYear = "mmm yyyy";
    public const string DateDayMonth = "dd/mm/yyyy";
    
    // Times
    public const string Time12Hour = "h:mm AM/PM";
    public const string Time24Hour = "h:mm";
    public const string DateTime = "m/d/yyyy h:mm";
    
    // Numbers
    public const string Number = "#,##0.00";
    public const string NumberNoDecimals = "#,##0";
    public const string NumberOneDecimal = "#,##0.0";
    public const string Scientific = "0.00E+00";
    
    // Special
    public const string Text = "@";
    public const string Fraction = "# ?/?";
    public const string Accounting = "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)";
    public const string General = "General";
}

// Result type
public class RangeNumberFormatResult : ResultBase
{
    public string SheetName { get; set; } = string.Empty;
    public string RangeAddress { get; set; } = string.Empty;
    
    /// <summary>
    /// 2D array of number format codes (matches range dimensions)
    /// </summary>
    public List<List<string>> Formats { get; set; } = [];
    
    public int RowCount { get; set; }
    public int ColumnCount { get; set; }
}
```

---

### 2. Range Visual Formatting

```csharp
public interface IRangeCommands
{
    // === FONT OPERATIONS ===
    
    /// <summary>
    /// Gets font properties from first cell in range
    /// Excel COM: Range.Font
    /// </summary>
    Task<RangeFontResult> GetFontAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    /// <summary>
    /// Sets font properties for entire range
    /// Excel COM: Range.Font
    /// </summary>
    /// <param name="font">Font properties (null values = no change)</param>
    Task<OperationResult> SetFontAsync(IExcelBatch batch, string sheetName, string rangeAddress, FontOptions font);
    
    // === CELL APPEARANCE ===
    
    /// <summary>
    /// Gets background color from first cell in range
    /// Excel COM: Range.Interior.Color
    /// </summary>
    Task<RangeColorResult> GetBackgroundColorAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    /// <summary>
    /// Sets background color for entire range
    /// Excel COM: Range.Interior.Color
    /// </summary>
    /// <param name="color">RGB color as integer: (red) | (green << 8) | (blue << 16)</param>
    Task<OperationResult> SetBackgroundColorAsync(IExcelBatch batch, string sheetName, string rangeAddress, int color);
    
    /// <summary>
    /// Clears background color (resets to no fill)
    /// Excel COM: Range.Interior.ColorIndex = xlColorIndexNone
    /// </summary>
    Task<OperationResult> ClearBackgroundColorAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    /// <summary>
    /// Gets border settings from range
    /// Excel COM: Range.Borders
    /// </summary>
    Task<RangeBorderResult> GetBordersAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    /// <summary>
    /// Sets borders for range
    /// Excel COM: Range.Borders
    /// </summary>
    Task<OperationResult> SetBordersAsync(IExcelBatch batch, string sheetName, string rangeAddress, BorderOptions borders);
    
    /// <summary>
    /// Clears all borders from range
    /// Excel COM: Range.Borders.LineStyle = xlLineStyleNone
    /// </summary>
    Task<OperationResult> ClearBordersAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    // === ALIGNMENT ===
    
    /// <summary>
    /// Gets alignment properties from first cell in range
    /// Excel COM: Range.HorizontalAlignment, Range.VerticalAlignment, etc.
    /// </summary>
    Task<RangeAlignmentResult> GetAlignmentAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    /// <summary>
    /// Sets alignment properties for range
    /// Excel COM: Range alignment properties
    /// </summary>
    Task<OperationResult> SetAlignmentAsync(IExcelBatch batch, string sheetName, string rangeAddress, AlignmentOptions alignment);
    
    // === ROW HEIGHT / COLUMN WIDTH ===
    
    /// <summary>
    /// Auto-fits column widths to content
    /// Excel COM: Range.Columns.AutoFit()
    /// </summary>
    Task<OperationResult> AutoFitColumnsAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    /// <summary>
    /// Auto-fits row heights to content
    /// Excel COM: Range.Rows.AutoFit()
    /// </summary>
    Task<OperationResult> AutoFitRowsAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    /// <summary>
    /// Sets column width in points
    /// Excel COM: Range.ColumnWidth
    /// </summary>
    Task<OperationResult> SetColumnWidthAsync(IExcelBatch batch, string sheetName, string rangeAddress, double width);
    
    /// <summary>
    /// Sets row height in points
    /// Excel COM: Range.RowHeight
    /// </summary>
    Task<OperationResult> SetRowHeightAsync(IExcelBatch batch, string sheetName, string rangeAddress, double height);
}

// Supporting types
public class FontOptions
{
    public string? Name { get; set; }              // Font family (e.g., "Arial", "Calibri")
    public int? Size { get; set; }                 // Font size in points
    public bool? Bold { get; set; }                // Bold text
    public bool? Italic { get; set; }              // Italic text
    public int? Color { get; set; }                // RGB color
    public bool? Underline { get; set; }           // Underline
    public bool? Strikethrough { get; set; }       // Strikethrough
}

public class BorderOptions
{
    public BorderStyle Style { get; set; } = BorderStyle.Continuous;
    public BorderWeight Weight { get; set; } = BorderWeight.Thin;
    public int? Color { get; set; }                // RGB color (null = default black)
    public bool ApplyToAll { get; set; } = true;   // Apply to all edges (top, bottom, left, right)
    
    // Individual edge control (if ApplyToAll = false)
    public bool? Top { get; set; }
    public bool? Bottom { get; set; }
    public bool? Left { get; set; }
    public bool? Right { get; set; }
}

public enum BorderStyle
{
    None,           // xlLineStyleNone
    Continuous,     // xlContinuous
    Dashed,         // xlDash
    Dotted,         // xlDot
    DashDot,        // xlDashDot
    DashDotDot,     // xlDashDotDot
    Double          // xlDouble
}

public enum BorderWeight
{
    Hairline,       // xlHairline
    Thin,           // xlThin
    Medium,         // xlMedium
    Thick           // xlThick
}

public class AlignmentOptions
{
    public HorizontalAlignment? Horizontal { get; set; }
    public VerticalAlignment? Vertical { get; set; }
    public bool? WrapText { get; set; }            // Text wrapping
    public int? Indent { get; set; }               // Indentation level (0-15)
    public int? Orientation { get; set; }          // Text rotation in degrees (-90 to 90)
}

public enum HorizontalAlignment
{
    General,        // xlGeneral (Excel default)
    Left,           // xlLeft
    Center,         // xlCenter
    Right,          // xlRight
    Fill,           // xlFill
    Justify,        // xlJustify
    CenterAcrossSelection,  // xlCenterAcrossSelection
    Distributed     // xlDistributed
}

public enum VerticalAlignment
{
    Top,            // xlTop
    Center,         // xlCenter
    Bottom,         // xlBottom
    Justify,        // xlJustify
    Distributed     // xlDistributed
}

// Result types
public class RangeFontResult : ResultBase
{
    public string SheetName { get; set; } = string.Empty;
    public string RangeAddress { get; set; } = string.Empty;
    public string FontName { get; set; } = string.Empty;
    public int FontSize { get; set; }
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public int Color { get; set; }
    public bool Underline { get; set; }
    public bool Strikethrough { get; set; }
}

public class RangeColorResult : ResultBase
{
    public string SheetName { get; set; } = string.Empty;
    public string RangeAddress { get; set; } = string.Empty;
    public bool HasColor { get; set; }
    public int? Color { get; set; }                // RGB integer (null if no color)
    public int? Red { get; set; }                  // Red component 0-255
    public int? Green { get; set; }                // Green component 0-255
    public int? Blue { get; set; }                 // Blue component 0-255
    public string? HexColor { get; set; }          // #RRGGBB format
}

public class RangeBorderResult : ResultBase
{
    public string SheetName { get; set; } = string.Empty;
    public string RangeAddress { get; set; } = string.Empty;
    public bool HasBorders { get; set; }
    public string? Style { get; set; }
    public string? Weight { get; set; }
    public int? Color { get; set; }
}

public class RangeAlignmentResult : ResultBase
{
    public string SheetName { get; set; } = string.Empty;
    public string RangeAddress { get; set; } = string.Empty;
    public string Horizontal { get; set; } = string.Empty;
    public string Vertical { get; set; } = string.Empty;
    public bool WrapText { get; set; }
    public int Indent { get; set; }
    public int Orientation { get; set; }
}
```

---

### 3. Range Data Validation

```csharp
public interface IRangeCommands
{
    // === DATA VALIDATION OPERATIONS ===
    
    /// <summary>
    /// Gets data validation settings from first cell in range
    /// Excel COM: Range.Validation
    /// </summary>
    Task<RangeValidationResult> GetValidationAsync(IExcelBatch batch, string sheetName, string rangeAddress);
    
    /// <summary>
    /// Adds data validation to range
    /// Excel COM: Range.Validation.Add()
    /// </summary>
    /// <param name="validation">Validation rule configuration</param>
    Task<OperationResult> AddValidationAsync(IExcelBatch batch, string sheetName, string rangeAddress, ValidationRule validation);
    
    /// <summary>
    /// Modifies existing validation rule
    /// Excel COM: Range.Validation.Modify()
    /// </summary>
    Task<OperationResult> ModifyValidationAsync(IExcelBatch batch, string sheetName, string rangeAddress, ValidationRule validation);
    
    /// <summary>
    /// Removes data validation from range
    /// Excel COM: Range.Validation.Delete()
    /// </summary>
    Task<OperationResult> RemoveValidationAsync(IExcelBatch batch, string sheetName, string rangeAddress);
}

// Supporting types
public class ValidationRule
{
    /// <summary>
    /// Type of validation (List, WholeNumber, Decimal, Date, Time, TextLength, Custom)
    /// </summary>
    public ValidationType Type { get; set; }
    
    /// <summary>
    /// Comparison operator (Between, NotBetween, Equal, NotEqual, Greater, Less, GreaterOrEqual, LessOrEqual)
    /// Only used for numeric/date validations (not for List or Custom)
    /// </summary>
    public ValidationOperator Operator { get; set; } = ValidationOperator.Between;
    
    /// <summary>
    /// First formula/value
    /// - List: Comma-separated values "Item1,Item2,Item3" or range reference "=$A$1:$A$10"
    /// - Number/Date/Time: Minimum value or single comparison value
    /// - TextLength: Minimum length
    /// - Custom: Formula expression (must return TRUE/FALSE)
    /// </summary>
    public string Formula1 { get; set; } = string.Empty;
    
    /// <summary>
    /// Second formula/value (only for Between/NotBetween operators)
    /// - Number/Date/Time: Maximum value
    /// - TextLength: Maximum length
    /// </summary>
    public string? Formula2 { get; set; }
    
    /// <summary>
    /// Whether to ignore blank cells (default: true)
    /// </summary>
    public bool IgnoreBlank { get; set; } = true;
    
    /// <summary>
    /// Whether to show input message when cell is selected
    /// </summary>
    public bool ShowInputMessage { get; set; } = false;
    
    /// <summary>
    /// Input message title
    /// </summary>
    public string? InputTitle { get; set; }
    
    /// <summary>
    /// Input message content
    /// </summary>
    public string? InputMessage { get; set; }
    
    /// <summary>
    /// Whether to show error alert on invalid data
    /// </summary>
    public bool ShowErrorAlert { get; set; } = true;
    
    /// <summary>
    /// Error alert style (Stop, Warning, Information)
    /// </summary>
    public ValidationAlertStyle ErrorStyle { get; set; } = ValidationAlertStyle.Stop;
    
    /// <summary>
    /// Error alert title
    /// </summary>
    public string? ErrorTitle { get; set; }
    
    /// <summary>
    /// Error alert message
    /// </summary>
    public string? ErrorMessage { get; set; }
}

public enum ValidationType
{
    List,           // xlValidateList - dropdown list
    WholeNumber,    // xlValidateWholeNumber
    Decimal,        // xlValidateDecimal
    Date,           // xlValidateDate
    Time,           // xlValidateTime
    TextLength,     // xlValidateTextLength
    Custom          // xlValidateCustom - formula-based
}

public enum ValidationOperator
{
    Between,        // xlBetween
    NotBetween,     // xlNotBetween
    Equal,          // xlEqual
    NotEqual,       // xlNotEqual
    Greater,        // xlGreater
    Less,           // xlLess
    GreaterOrEqual, // xlGreaterEqual
    LessOrEqual     // xlLessEqual
}

public enum ValidationAlertStyle
{
    Stop,           // xlValidAlertStop - prevents invalid data
    Warning,        // xlValidAlertWarning - warns but allows
    Information     // xlValidAlertInformation - info only
}

public class RangeValidationResult : ResultBase
{
    public string SheetName { get; set; } = string.Empty;
    public string RangeAddress { get; set; } = string.Empty;
    public bool HasValidation { get; set; }
    public string? Type { get; set; }
    public string? Operator { get; set; }
    public string? Formula1 { get; set; }
    public string? Formula2 { get; set; }
    public bool IgnoreBlank { get; set; }
    public bool ShowInputMessage { get; set; }
    public string? InputTitle { get; set; }
    public string? InputMessage { get; set; }
    public bool ShowErrorAlert { get; set; }
    public string? ErrorStyle { get; set; }
    public string? ErrorTitle { get; set; }
    public string? ErrorMessage { get; set; }
}
```

---

### 4. Table Formatting & Validation

```csharp
public interface ITableCommands
{
    // === NUMBER FORMATTING (delegates to RangeCommands) ===
    
    /// <summary>
    /// Gets number formats for a table column
    /// Delegates to RangeCommands.GetNumberFormatsAsync() on column range
    /// </summary>
    Task<RangeNumberFormatResult> GetColumnNumberFormatAsync(IExcelBatch batch, string tableName, string columnName);
    
    /// <summary>
    /// Sets uniform number format for entire table column
    /// Delegates to RangeCommands.SetNumberFormatAsync() on column data range (excludes header)
    /// </summary>
    /// <param name="formatCode">Excel format code (e.g., "$#,##0.00", "0.00%")</param>
    Task<OperationResult> SetColumnNumberFormatAsync(IExcelBatch batch, string tableName, string columnName, string formatCode);
    
    // === VISUAL FORMATTING (delegates to RangeCommands) ===
    
    /// <summary>
    /// Sets font for table column data cells
    /// Delegates to RangeCommands.SetFontAsync() on column data range
    /// </summary>
    Task<OperationResult> SetColumnFontAsync(IExcelBatch batch, string tableName, string columnName, FontOptions font);
    
    /// <summary>
    /// Sets background color for table column data cells
    /// Delegates to RangeCommands.SetBackgroundColorAsync()
    /// </summary>
    Task<OperationResult> SetColumnBackgroundColorAsync(IExcelBatch batch, string tableName, string columnName, int color);
    
    /// <summary>
    /// Sets alignment for table column data cells
    /// Delegates to RangeCommands.SetAlignmentAsync()
    /// </summary>
    Task<OperationResult> SetColumnAlignmentAsync(IExcelBatch batch, string tableName, string columnName, AlignmentOptions alignment);
    
    /// <summary>
    /// Auto-fits table column width to content
    /// Delegates to RangeCommands.AutoFitColumnsAsync()
    /// </summary>
    Task<OperationResult> AutoFitColumnAsync(IExcelBatch batch, string tableName, string columnName);
    
    // === DATA VALIDATION (delegates to RangeCommands) ===
    
    /// <summary>
    /// Adds data validation to table column data cells
    /// Delegates to RangeCommands.AddValidationAsync() on column data range
    /// </summary>
    Task<OperationResult> AddColumnValidationAsync(IExcelBatch batch, string tableName, string columnName, ValidationRule validation);
    
    /// <summary>
    /// Removes data validation from table column
    /// Delegates to RangeCommands.RemoveValidationAsync()
    /// </summary>
    Task<OperationResult> RemoveColumnValidationAsync(IExcelBatch batch, string tableName, string columnName);
}
```

---

### 5. PivotTable Formatting

**Note:** PivotTable formatting is **ALREADY IMPLEMENTED** in Phase 1 (as of October 30, 2025). See [PIVOTTABLE-API-SPECIFICATION.md](PIVOTTABLE-API-SPECIFICATION.md) for complete details.

```csharp
public interface IPivotTableCommands
{
    // === NUMBER FORMATTING (IMPLEMENTED) ===
    
    /// <summary>
    /// Sets number format for value field in PivotTable
    /// Excel COM: PivotField.NumberFormat
    /// </summary>
    /// <param name="fieldName">Name of value field (e.g., "Sum of Sales")</param>
    /// <param name="numberFormat">Excel format code (e.g., "$#,##0.00", "0.00%")</param>
    Task<PivotFieldResult> SetFieldFormatAsync(IExcelBatch batch, string pivotTableName, 
        string fieldName, string numberFormat);
    
    // === LAYOUT & STYLE (IMPLEMENTED) ===
    
    /// <summary>
    /// Sets PivotTable layout form
    /// Excel COM: PivotTable.RowAxisLayout() or PivotTable.LayoutForm
    /// </summary>
    /// <param name="layout">Compact, Outline, or Tabular</param>
    Task<OperationResult> SetLayoutAsync(IExcelBatch batch, string pivotTableName, 
        PivotTableLayout layout);
    
    /// <summary>
    /// Sets PivotTable visual style
    /// Excel COM: PivotTable.TableStyle2
    /// </summary>
    /// <param name="styleName">
    /// Built-in style name (e.g., "PivotStyleMedium2", "PivotStyleLight16")
    /// See PivotTableStylePresets for common options
    /// </param>
    Task<OperationResult> SetStyleAsync(IExcelBatch batch, string pivotTableName, 
        string styleName);
}

// Layout options
public enum PivotTableLayout
{
    Compact,    // xlCompactForm - hierarchical, space-saving
    Outline,    // xlOutlineForm - hierarchical with subtotals
    Tabular     // xlTabularForm - flat, traditional layout
}

/// <summary>
/// Common PivotTable style names for LLM convenience
/// </summary>
public static class PivotTableStylePresets
{
    // Light Styles (subtle colors)
    public const string Light1 = "PivotStyleLight1";
    public const string Light2 = "PivotStyleLight2";
    public const string Light9 = "PivotStyleLight9";
    public const string Light16 = "PivotStyleLight16";
    public const string Light21 = "PivotStyleLight21";
    
    // Medium Styles (balanced colors)
    public const string Medium1 = "PivotStyleMedium1";
    public const string Medium2 = "PivotStyleMedium2";    // Popular choice
    public const string Medium3 = "PivotStyleMedium3";
    public const string Medium9 = "PivotStyleMedium9";
    public const string Medium10 = "PivotStyleMedium10";
    
    // Dark Styles (high contrast)
    public const string Dark1 = "PivotStyleDark1";
    public const string Dark2 = "PivotStyleDark2";
    public const string Dark11 = "PivotStyleDark11";
    
    // None (remove styling)
    public const string None = "";
}
```

**PivotTable Formatting Capabilities (Implemented):**

✅ **Number Formatting** - Format value fields (Sum of Sales, Average Price, etc.) - **PERSISTENT (survives refresh)**  
✅ **Layout Forms** - Compact, Outline, or Tabular layout - **PERSISTENT**  
✅ **Visual Styles** - 28 built-in PivotTable styles (Light, Medium, Dark) - **PERSISTENT**  

**How It Works:**
```csharp
// SetFieldFormatAsync() sets format on the PivotField object itself
// Excel COM: dataField.NumberFormat = "$#,##0.00"
// This is a FIELD-LEVEL setting, not a cell-level format
// Result: Format persists across RefreshTable(), save/reopen, data updates
```

**Advanced Formatting (NOT in current scope):**
- ❌ Row/column header cell formatting (use RangeCommands, but lost on refresh)
- ❌ Grand total cell formatting (use RangeCommands, but lost on refresh)
- ❌ Conditional formatting within PivotTable (use RangeCommands, but lost on refresh)
- ❌ Individual data cell formatting (recalculated on refresh, formats lost)

**Recommended Approach for PivotTable Formatting:**

1. ✅ **Use PivotTableCommands** for structure and value field formats (PERSISTENT)
   - `SetFieldFormatAsync()` for number formats on value fields
   - `SetLayoutAsync()` for layout form
   - `SetStyleAsync()` for visual appearance
   
2. ⚠️ **Use RangeCommands carefully** for additional formatting (NOT PERSISTENT)
   - Can format header cells, grand totals, or specific data cells
   - **Warning:** These formats will be **lost on refresh**
   - Only use for one-time formatting or static PivotTables that won't refresh

3. 💡 **Best Practice:** Always use field-level formats when possible
   - Field formats are part of the PivotTable definition
   - They persist across all PivotTable operations
   - They're saved with the workbook

---

## MCP Server Integration

### Updated excel_range Tool

```json
{
  "name": "excel_range",
  "actions": [
    // EXISTING (Phase 1)
    "get-values", "set-values",
    "get-formulas", "set-formulas",
    "clear-all", "clear-contents", "clear-formats",
    "copy", "copy-values", "copy-formulas",
    "insert-cells", "delete-cells", "insert-rows", "delete-rows", "insert-columns", "delete-columns",
    "find", "replace", "sort",
    "get-used-range", "get-current-region", "get-range-info",
    "add-hyperlink", "remove-hyperlink", "list-hyperlinks", "get-hyperlink",
    
    // NEW: NUMBER FORMATTING
    "get-number-formats",
    "set-number-format",
    "set-number-formats",
    
    // NEW: VISUAL FORMATTING
    "get-font", "set-font",
    "get-background-color", "set-background-color", "clear-background-color",
    "get-borders", "set-borders", "clear-borders",
    "get-alignment", "set-alignment",
    "auto-fit-columns", "auto-fit-rows",
    "set-column-width", "set-row-height",
    
    // NEW: DATA VALIDATION
    "get-validation",
    "add-validation",
    "modify-validation",
    "remove-validation"
  ]
}
```

### Updated excel_table Tool

```json
{
  "name": "excel_table",
  "actions": [
    // EXISTING
    "list", "create", "rename", "delete", "info", "resize",
    "toggle-totals", "set-column-total", "append-rows", "set-style",
    "add-to-datamodel",
    "apply-filter", "apply-filter-values", "clear-filters", "get-filters",
    "add-column", "remove-column", "rename-column",
    "get-structured-reference",
    "sort", "sort-multi",
    
    // NEW: FORMATTING & VALIDATION
    "get-column-number-format", "set-column-number-format",
    "set-column-font", "set-column-background-color", "set-column-alignment", "auto-fit-column",
    "add-column-validation", "remove-column-validation"
  ]
}
```

### excel_pivottable Tool (Existing - Already Implemented)

```json
{
  "name": "excel_pivottable",
  "description": "PivotTable operations - create, configure, format, and analyze data",
  "actions": [
    // EXISTING (18 actions implemented in Phase 1)
    "create", "delete", "list", "info",
    "list-fields", "add-row-field", "add-column-field", "add-value-field", 
    "add-filter-field", "remove-field", "move-field",
    "set-field-function", "set-field-name", "set-field-format",
    "get-data", "set-field-filter", "sort-field", "refresh",
    
    // FORMATTING (IMPLEMENTED)
    "set-layout",           // Compact, Outline, or Tabular
    "set-style"             // PivotTable visual styles
  ]
}
```

**Note:** PivotTable formatting commands (`set-layout`, `set-style`, `set-field-format`) are **already implemented** and functional. See [PIVOTTABLE-API-SPECIFICATION.md](PIVOTTABLE-API-SPECIFICATION.md) for complete documentation.

---

## MCP Server Parameter Reference (Critical for LLMs)

### Common Parameters Across All Actions

**Required on ALL actions:**
- `excelPath` (string) - Absolute path to Excel file (e.g., "C:\\Users\\user\\sales.xlsx")
- `action` (string) - Action to perform (e.g., "set-number-format", "add-validation")

**Optional batch parameter:**
- `batchId` (string) - If using batch mode, reference batch session ID

### excel_range Actions - Parameter Details

#### Number Formatting

**get-number-formats**
- Required: `excelPath`, `sheetName`, `rangeAddress`
- Returns: `{ success, sheetName, rangeAddress, formats: [[string]], rowCount, columnCount }`

**set-number-format**
- Required: `excelPath`, `sheetName`, `rangeAddress`, `formatCode`
- `formatCode`: Excel format string (e.g., "$#,##0.00", "0.00%", "m/d/yyyy", "@")
- Returns: `{ success, message }`

**set-number-formats**
- Required: `excelPath`, `sheetName`, `rangeAddress`, `formats`
- `formats`: 2D array of format codes `[["$#,##0", "0.00%"], ["m/d/yyyy", "General"]]`
- Returns: `{ success, message }`

#### Visual Formatting

**set-font**
- Required: `excelPath`, `sheetName`, `rangeAddress`, `font`
- `font` (object with ALL properties optional):
  - `name` (string): Font family (e.g., "Arial", "Calibri", "Times New Roman")
  - `size` (number): Font size in points (e.g., 10, 12, 14)
  - `bold` (boolean): Bold text (true/false)
  - `italic` (boolean): Italic text (true/false)
  - `color` (number): RGB integer (see RGB calculation below)
  - `underline` (boolean): Underline text (true/false)
  - `strikethrough` (boolean): Strikethrough text (true/false)
- Returns: `{ success, message }`

**set-background-color**
- Required: `excelPath`, `sheetName`, `rangeAddress`, `color`
- `color` (number): RGB integer (see RGB calculation below)
- Returns: `{ success, message }`

**set-borders**
- Required: `excelPath`, `sheetName`, `rangeAddress`, `borders`
- `borders` (object):
  - `style` (string): "Continuous", "Dashed", "Dotted", "DashDot", "DashDotDot", "Double", "None"
  - `weight` (string): "Hairline", "Thin", "Medium", "Thick"
  - `color` (number): RGB integer (optional, default black)
  - `applyToAll` (boolean): Apply to all edges (true) or specify individual edges (false)
  - If `applyToAll` = false:
    - `top` (boolean): Apply to top edge
    - `bottom` (boolean): Apply to bottom edge
    - `left` (boolean): Apply to left edge
    - `right` (boolean): Apply to right edge
- Returns: `{ success, message }`

**set-alignment**
- Required: `excelPath`, `sheetName`, `rangeAddress`, `alignment`
- `alignment` (object with ALL properties optional):
  - `horizontal` (string): "General", "Left", "Center", "Right", "Fill", "Justify", "CenterAcrossSelection", "Distributed"
  - `vertical` (string): "Top", "Center", "Bottom", "Justify", "Distributed"
  - `wrapText` (boolean): Enable text wrapping (true/false)
  - `indent` (number): Indentation level 0-15
  - `orientation` (number): Text rotation in degrees (-90 to 90)
- Returns: `{ success, message }`

#### Data Validation

**add-validation**
- Required: `excelPath`, `sheetName`, `rangeAddress`, `validation`
- `validation` (object):
  - `type` (string): "List", "WholeNumber", "Decimal", "Date", "Time", "TextLength", "Custom"
  - `operator` (string): "Between", "NotBetween", "Equal", "NotEqual", "Greater", "Less", "GreaterOrEqual", "LessOrEqual"
  - `formula1` (string): First value/formula/list
  - `formula2` (string, optional): Second value (for Between/NotBetween)
  - `ignoreBlank` (boolean, optional): Ignore blank cells (default: true)
  - `showInputMessage` (boolean, optional): Show input message (default: false)
  - `inputTitle` (string, optional): Input message title
  - `inputMessage` (string, optional): Input message text
  - `showErrorAlert` (boolean, optional): Show error alert (default: true)
  - `errorStyle` (string, optional): "Stop", "Warning", "Information" (default: "Stop")
  - `errorTitle` (string, optional): Error alert title
  - `errorMessage` (string, optional): Error alert message
- Returns: `{ success, message }`

### excel_table Actions - Parameter Details

**set-column-number-format**
- Required: `excelPath`, `tableName`, `columnName`, `formatCode`
- Returns: `{ success, message }`

**set-column-font**
- Required: `excelPath`, `tableName`, `columnName`, `font`
- `font`: Same structure as excel_range set-font
- Returns: `{ success, message }`

**set-column-background-color**
- Required: `excelPath`, `tableName`, `columnName`, `color`
- Returns: `{ success, message }`

**add-column-validation**
- Required: `excelPath`, `tableName`, `columnName`, `validation`
- `validation`: Same structure as excel_range add-validation
- Returns: `{ success, message }`

### excel_pivottable Actions - Parameter Details

**set-field-format**
- Required: `excelPath`, `pivotTableName`, `fieldName`, `numberFormat`
- `fieldName`: Use value field name from list-fields (e.g., "Sum of Sales", "Average of Price")
- Returns: `{ success, fieldName, customName, area, numberFormat }`

**set-layout**
- Required: `excelPath`, `pivotTableName`, `layout`
- `layout` (string): "Compact", "Outline", "Tabular"
- Returns: `{ success, message }`

**set-style**
- Required: `excelPath`, `pivotTableName`, `styleName`
- `styleName` (string): Built-in style name (e.g., "PivotStyleMedium2", "PivotStyleLight16")
- Returns: `{ success, message }`

---

## RGB Color Calculation (Critical Reference)

**How to Calculate RGB Integer for Color Parameters:**

```
Formula: RGB(red, green, blue) = red + (green × 256) + (blue × 256²)
Alternative: red + (green << 8) + (blue << 16)

Where: red, green, blue are each 0-255
```

**Common Colors (Ready to Use):**

| Color Name | RGB Values | Integer Value | Use For |
|------------|------------|---------------|---------|
| **Red** | (255, 0, 0) | 255 | Errors, alerts, negative values |
| **Green** | (0, 255, 0) | 65280 | Success, positive values |
| **Blue** | (0, 0, 255) | 16711680 | Headers, links |
| **Yellow** | (255, 255, 0) | 65535 | Highlights, warnings |
| **Orange** | (255, 165, 0) | 42495 | Warnings, important items |
| **Purple** | (128, 0, 128) | 8388736 | Categories, special items |
| **Light Gray** | (211, 211, 211) | 13882323 | Disabled, secondary |
| **Light Blue** | (173, 216, 230) | 15128749 | Header backgrounds |
| **Light Green** | (144, 238, 144) | 9498256 | Positive highlights |
| **Light Yellow** | (255, 255, 224) | 14745599 | Subtle highlights |
| **White** | (255, 255, 255) | 16777215 | Clear/reset background |
| **Black** | (0, 0, 0) | 0 | Text, borders |

**When User Says Color Name:**
```
"red background" → color: 255
"yellow highlight" → color: 65535
"green text" → font.color: 65280
"light blue cells" → color: 15128749
```

---

## Batch Mode Usage (Performance Optimization)

**When to Use Batch Mode:**
- Formatting 3+ ranges/columns
- Multiple operations on same file
- Complete workbook setup workflows

**How to Use Batch Mode with MCP Server:**

```json
// Step 1: Begin batch session
{
  "tool": "begin_excel_batch",
  "excelPath": "C:\\reports\\sales.xlsx"
}
// Returns: { "success": true, "batchId": "batch_abc123" }

// Step 2: Execute multiple operations (use batchId)
{
  "tool": "excel_range",
  "action": "set-number-format",
  "batchId": "batch_abc123",
  "sheetName": "Sales",
  "rangeAddress": "D2:D100",
  "formatCode": "$#,##0.00"
}

{
  "tool": "excel_range",
  "action": "set-font",
  "batchId": "batch_abc123",
  "sheetName": "Sales",
  "rangeAddress": "A1:E1",
  "font": { "bold": true, "size": 12 }
}

{
  "tool": "excel_range",
  "action": "add-validation",
  "batchId": "batch_abc123",
  "sheetName": "Sales",
  "rangeAddress": "F2:F100",
  "validation": {
    "type": "List",
    "formula1": "Active,Inactive"
  }
}

// Step 3: Commit batch (saves all changes)
{
  "tool": "commit_excel_batch",
  "batchId": "batch_abc123",
  "save": true
}
// Returns: { "success": true, "message": "Batch committed, changes saved" }
```

**Benefits:**
- ✅ Excel opened once for all operations
- ✅ Changes saved once at end
- ✅ 5-10x faster for multiple operations
- ✅ Atomic - all succeed or all fail

---

## Common Mistakes & How to Avoid Them

### Range Formatting Mistakes

**❌ Mistake 1: Wrong range address separator**
```json
// WRONG
"rangeAddress": "A1-D10"

// CORRECT
"rangeAddress": "A1:D10"
```

**❌ Mistake 2: Wrong alignment case**
```json
// WRONG
"alignment": { "horizontal": "center" }

// CORRECT
"alignment": { "horizontal": "Center" }  // Capitalize enum values
```

**❌ Mistake 3: Formatting entire columns inefficiently**
```json
// SLOW: Formats millions of cells
"rangeAddress": "A:Z"

// FAST: Format only used range
"rangeAddress": "A1:Z1000"
// Or use get-used-range first to find actual data
```

**❌ Mistake 4: Wrong RGB color calculation**
```json
// WRONG: Using separate R, G, B values
"color": { "red": 255, "green": 0, "blue": 0 }

// CORRECT: Use integer
"color": 255  // For red
```

**❌ Mistake 5: Providing partial 2D array for set-number-formats**
```json
// WRONG: 2x3 range but only 1x2 formats array
"rangeAddress": "A1:C2",
"formats": [["$#,##0", "0.00%"]]

// CORRECT: Match dimensions
"formats": [
  ["$#,##0", "0.00%", "m/d/yyyy"],
  ["$#,##0", "0.00%", "m/d/yyyy"]
]
```

### Validation Mistakes

**❌ Mistake 6: Wrong validation type for dropdown**
```json
// WRONG
"type": "Dropdown"

// CORRECT
"type": "List"
```

**❌ Mistake 7: Forgetting formula2 for Between operator**
```json
// WRONG: Between requires two values
"type": "WholeNumber",
"operator": "Between",
"formula1": "1"
// Missing formula2!

// CORRECT
"type": "WholeNumber",
"operator": "Between",
"formula1": "1",
"formula2": "100"
```

**❌ Mistake 8: Using formula reference without = prefix**
```json
// WRONG: List from range needs = prefix
"formula1": "A1:A10"

// CORRECT
"formula1": "=$A$1:$A$10"
```

### Table Formatting Mistakes

**❌ Mistake 9: Using range address instead of table name**
```json
// WRONG: excel_table needs table name
"tableName": "A1:D100"

// CORRECT
"tableName": "SalesData"
```

**❌ Mistake 10: Formatting table headers (not supported)**
```json
// WRONG: Headers have fixed formatting from table style
{
  "tool": "excel_table",
  "action": "set-column-font",
  "columnName": "Amount",
  "font": { "bold": true }  // Affects data cells only, not header
}

// CORRECT: Use table styles instead
{
  "tool": "excel_table",
  "action": "set-style",
  "tableName": "SalesData",
  "styleName": "TableStyleMedium2"
}
```

---

## LLM Usage Examples

### Example 1: Format Sales Report

```json
// Scenario: LLM formatting a professional sales report

// Step 1: Format currency column
{
  "tool": "excel_range",
  "action": "set-number-format",
  "sheetName": "Sales",
  "rangeAddress": "D2:D100",
  "formatCode": "$#,##0.00"
}

// Step 2: Format percentage column
{
  "tool": "excel_range",
  "action": "set-number-format",
  "sheetName": "Sales",
  "rangeAddress": "E2:E100",
  "formatCode": "0.00%"
}

// Step 3: Format date column
{
  "tool": "excel_range",
  "action": "set-number-format",
  "sheetName": "Sales",
  "rangeAddress": "A2:A100",
  "formatCode": "m/d/yyyy"
}

// Step 4: Bold headers
{
  "tool": "excel_range",
  "action": "set-font",
  "sheetName": "Sales",
  "rangeAddress": "A1:E1",
  "font": { "bold": true, "size": 12 }
}

// Step 5: Center align headers
{
  "tool": "excel_range",
  "action": "set-alignment",
  "sheetName": "Sales",
  "rangeAddress": "A1:E1",
  "alignment": { "horizontal": "Center" }
}

// Step 6: Add borders
{
  "tool": "excel_range",
  "action": "set-borders",
  "sheetName": "Sales",
  "rangeAddress": "A1:E100",
  "borders": { "style": "Continuous", "weight": "Thin" }
}

// Step 7: Auto-fit columns
{
  "tool": "excel_range",
  "action": "auto-fit-columns",
  "sheetName": "Sales",
  "rangeAddress": "A:E"
}
```

### Example 2: Data Entry Form with Validation

```json
// Scenario: LLM creating data entry form with dropdowns

// Step 1: Add status dropdown validation
{
  "tool": "excel_range",
  "action": "add-validation",
  "sheetName": "Orders",
  "rangeAddress": "D2:D1000",
  "validation": {
    "type": "List",
    "formula1": "Pending,Processing,Shipped,Delivered,Cancelled",
    "showErrorAlert": true,
    "errorStyle": "Stop",
    "errorTitle": "Invalid Status",
    "errorMessage": "Please select a status from the dropdown list."
  }
}

// Step 2: Add quantity number validation (1-999)
{
  "tool": "excel_range",
  "action": "add-validation",
  "sheetName": "Orders",
  "rangeAddress": "E2:E1000",
  "validation": {
    "type": "WholeNumber",
    "operator": "Between",
    "formula1": "1",
    "formula2": "999",
    "showErrorAlert": true,
    "errorTitle": "Invalid Quantity",
    "errorMessage": "Quantity must be between 1 and 999."
  }
}

// Step 3: Add email text length validation
{
  "tool": "excel_range",
  "action": "add-validation",
  "sheetName": "Orders",
  "rangeAddress": "C2:C1000",
  "validation": {
    "type": "TextLength",
    "operator": "LessOrEqual",
    "formula1": "100",
    "showInputMessage": true,
    "inputTitle": "Email Address",
    "inputMessage": "Enter customer email (max 100 characters)"
  }
}
```

### Example 3: Table Column Formatting

```json
// Scenario: LLM formatting table columns professionally

// Step 1: Format amount column as currency
{
  "tool": "excel_table",
  "action": "set-column-number-format",
  "tableName": "SalesData",
  "columnName": "Amount",
  "formatCode": "$#,##0.00"
}

// Step 2: Format growth column as percentage
{
  "tool": "excel_table",
  "action": "set-column-number-format",
  "tableName": "SalesData",
  "columnName": "Growth",
  "formatCode": "0.0%"
}

// Step 3: Add status dropdown validation
{
  "tool": "excel_table",
  "action": "add-column-validation",
  "tableName": "SalesData",
  "columnName": "Status",
  "validation": {
    "type": "List",
    "formula1": "Active,Inactive,Pending"
  }
}

// Step 4: Center align status column
{
  "tool": "excel_table",
  "action": "set-column-alignment",
  "tableName": "SalesData",
  "columnName": "Status",
  "alignment": { "horizontal": "Center" }
}

// Step 5: Auto-fit all columns
{
  "tool": "excel_table",
  "action": "auto-fit-column",
  "tableName": "SalesData",
  "columnName": "Amount"
}
// Repeat for other columns or use range auto-fit for all at once
```

### Example 4: PivotTable Professional Formatting

```json
// Scenario: LLM creating and formatting a professional PivotTable

// Step 1: Create PivotTable from table
{
  "tool": "excel_pivottable",
  "action": "create",
  "sourceType": "table",
  "sourceTable": "SalesData",
  "destinationSheet": "Analysis",
  "destinationCell": "A1",
  "pivotTableName": "SalesPivot"
}

// Step 2: Configure fields
{
  "tool": "excel_pivottable",
  "action": "add-row-field",
  "pivotTableName": "SalesPivot",
  "fieldName": "Region"
}

{
  "tool": "excel_pivottable",
  "action": "add-column-field",
  "pivotTableName": "SalesPivot",
  "fieldName": "Quarter"
}

{
  "tool": "excel_pivottable",
  "action": "add-value-field",
  "pivotTableName": "SalesPivot",
  "fieldName": "Amount",
  "function": "Sum",
  "customName": "Total Sales"
}

// Step 3: Format value field as currency
{
  "tool": "excel_pivottable",
  "action": "set-field-format",
  "pivotTableName": "SalesPivot",
  "fieldName": "Total Sales",
  "numberFormat": "$#,##0"
}

// Step 4: Set layout to Tabular (easier to read)
{
  "tool": "excel_pivottable",
  "action": "set-layout",
  "pivotTableName": "SalesPivot",
  "layout": "Tabular"
}

// Step 5: Apply professional style
{
  "tool": "excel_pivottable",
  "action": "set-style",
  "pivotTableName": "SalesPivot",
  "styleName": "PivotStyleMedium2"
}

// Step 6: Refresh to show data
{
  "tool": "excel_pivottable",
  "action": "refresh",
  "pivotTableName": "SalesPivot"
}
```

---

## LLM Decision Logic

### 1. Number Format Selection

When user says a format type, use these codes:

| User Request | Format Code | Use For |
|--------------|-------------|---------|
| "currency" | `$#,##0.00` | Money amounts |
| "percentage" | `0.00%` | Percentages with 2 decimals |
| "percent" | `0%` | Percentages without decimals |
| "date" | `m/d/yyyy` | US date format |
| "date long" | `mmmm d, yyyy` | Full date (January 1, 2025) |
| "time" | `h:mm AM/PM` | 12-hour time |
| "number" | `#,##0.00` | General numbers with commas |
| "text" | `@` | Force text format |
| "accounting" | See `NumberFormatPresets.Accounting` | Accounting format with alignment |

### 2. Visual Formatting Decisions

**Font Formatting:**
```
User says "make headers bold" → SetFontAsync({ bold: true })
User says "increase font size" → SetFontAsync({ size: 14 })
User says "red text" → SetFontAsync({ color: RGB(255, 0, 0) })
User says "italicize" → SetFontAsync({ italic: true })
```

**Color Application:**
```
User says "highlight in yellow" → SetBackgroundColorAsync(RGB(255, 255, 0))
User says "green background" → SetBackgroundColorAsync(RGB(0, 255, 0))
User says "remove color" → ClearBackgroundColorAsync()
```

**Borders:**
```
User says "add borders" → SetBordersAsync({ style: Continuous, weight: Thin })
User says "thick border" → SetBordersAsync({ weight: Thick })
User says "remove borders" → ClearBordersAsync()
```

**Alignment:**
```
User says "center align" → SetAlignmentAsync({ horizontal: Center })
User says "wrap text" → SetAlignmentAsync({ wrapText: true })
User says "indent" → SetAlignmentAsync({ indent: 2 })
```

### 3. Validation Type Selection

```
User says "dropdown list" or "select from list"
  → Type: List, Formula1: "Item1,Item2,Item3"

User says "number between X and Y"
  → Type: WholeNumber/Decimal, Operator: Between, Formula1: "X", Formula2: "Y"

User says "date range" or "date after X"
  → Type: Date, Operator: Greater, Formula1: "1/1/2025"

User says "maximum length" or "max X characters"
  → Type: TextLength, Operator: LessOrEqual, Formula1: "X"

User says "custom rule" or "formula validation"
  → Type: Custom, Formula1: "=AND(A1>0, A1<100)"
```

### 4. Range vs Table vs PivotTable Decision

```
If user mentions PivotTable or pivot table
  → Use excel_pivottable actions
  → For value field formatting: use set-field-format
  → For layout: use set-layout (Compact, Outline, Tabular)
  → For visual style: use set-style (PivotStyleMedium2, etc.)
  → Note: Cell-level formatting lost on refresh - use field formats

If user mentions table name explicitly
  → Use excel_table actions (e.g., set-column-number-format)

If user mentions specific range or worksheet
  → use excel_range actions

For new tables being created
  → Create table first, then use table formatting actions
  → More efficient than formatting range then converting to table

For PivotTable creation + formatting workflow
  → Create PivotTable → Add fields → Format value fields → Set layout → Set style → Refresh
```

### 5. PivotTable Formatting Guide (Critical for LLMs)

**When User Says "Format the PivotTable":**

```
Step 1: Identify what needs formatting
  → Value fields (numbers) → Use set-field-format (PERSISTENT)
  → Layout/structure → Use set-layout (PERSISTENT)
  → Visual appearance → Use set-style (PERSISTENT)
  → Specific cells → DANGER: Use excel_range (NOT PERSISTENT - see pitfalls)

Step 2: Format value fields FIRST (before layout/style)
  → Locate value field name (e.g., "Sum of Sales", "Average Price")
  → Apply number format using set-field-format
  → Examples:
    - Currency: "$#,##0.00" or "$#,##0"
    - Percentage: "0.00%" or "0%"
    - Numbers: "#,##0.00" or "#,##0"
    - Dates: "m/d/yyyy" (rarely needed in values)

Step 3: Set layout for readability
  → Compact: Best for space-saving, hierarchical data
  → Outline: Best for subtotals and grouping
  → Tabular: Best for flat data, easier to read (RECOMMENDED for most cases)

Step 4: Apply visual style
  → Medium styles most popular (PivotStyleMedium2)
  → Light styles for subtle appearance
  → Dark styles for high contrast

Step 5: ALWAYS refresh after formatting
  → Refresh materializes the formatting changes
  → Without refresh, formats may not appear correctly
```

**Critical Pitfalls to Avoid:**

```
❌ PITFALL 1: Using excel_range to format PivotTable data cells
  Problem: Formats lost on next refresh
  Example: set-number-format on "C5:C20" (data area)
  Why: PivotTable regenerates cells on refresh
  Solution: Use set-field-format on value field instead

❌ PITFALL 2: Formatting before adding all fields
  Problem: Layout changes as fields added, formats misaligned
  Example: Format cells, then add column field
  Solution: Add ALL fields first, THEN format

❌ PITFALL 3: Forgetting to refresh after formatting
  Problem: Formatting doesn't appear or incomplete
  Example: set-field-format, then immediately read data
  Solution: ALWAYS call refresh action after formatting

❌ PITFALL 4: Trying to format row/column headers persistently
  Problem: Header cells regenerate on refresh
  Example: Bold the "Region" header cells
  Solution: Not supported persistently - use styles instead
           OR format with excel_range knowing it's temporary

❌ PITFALL 5: Using wrong field name
  Problem: Field name is display name, not source name
  Example: Field is "Sum of Sales" not "Sales"
  Solution: Use list-fields to see actual value field names
           Value fields are usually "Sum of X", "Count of Y", etc.

❌ PITFALL 6: Applying validation to PivotTable cells
  Problem: Validation lost on refresh, cells are calculated
  Example: Add dropdown to data cells
  Solution: Don't add validation to PivotTables - they're read-only summaries
```

**Correct PivotTable Formatting Workflow:**

```javascript
// ✅ CORRECT: Field-level formatting (persistent)
{
  "tool": "excel_pivottable",
  "action": "set-field-format",
  "pivotTableName": "SalesPivot",
  "fieldName": "Sum of Amount",      // Use actual field name from list-fields
  "numberFormat": "$#,##0"
}

// ✅ CORRECT: Layout for readability
{
  "tool": "excel_pivottable",
  "action": "set-layout",
  "pivotTableName": "SalesPivot",
  "layout": "Tabular"                // Most readable for most users
}

// ✅ CORRECT: Professional style
{
  "tool": "excel_pivottable",
  "action": "set-style",
  "pivotTableName": "SalesPivot",
  "styleName": "PivotStyleMedium2"   // Popular, professional
}

// ✅ CORRECT: Refresh to materialize
{
  "tool": "excel_pivottable",
  "action": "refresh",
  "pivotTableName": "SalesPivot"
}

// ❌ WRONG: Cell-level formatting (lost on refresh)
{
  "tool": "excel_range",
  "action": "set-number-format",
  "sheetName": "Analysis",
  "rangeAddress": "C5:C20",          // Don't format PivotTable data cells this way!
  "formatCode": "$#,##0.00"
}
```

**When to Use Each Formatting Approach:**

```
Use set-field-format when:
  ✅ User wants to format numbers in PivotTable
  ✅ Format needs to persist across refreshes
  ✅ Formatting sum, average, count, etc. values
  ✅ User says "format the sales column" (value field)

Use set-layout when:
  ✅ User wants to change PivotTable structure
  ✅ User says "make it easier to read"
  ✅ User wants subtotals shown differently
  ✅ Default: Choose Tabular (most readable)

Use set-style when:
  ✅ User wants professional appearance
  ✅ User mentions colors, banding, headers
  ✅ User says "make it look nice"
  ✅ Default: PivotStyleMedium2 or PivotStyleMedium9

Use excel_range when (RARE):
  ✅ User explicitly wants one-time formatting
  ✅ PivotTable will never refresh
  ✅ Formatting grand total row/column specifically
  ⚠️  WARN USER: Format will be lost on refresh
```

**Field Name Discovery Pattern:**

```javascript
// ALWAYS discover field names first if unsure
{
  "tool": "excel_pivottable",
  "action": "list-fields",
  "pivotTableName": "SalesPivot"
}

// Response shows:
// - Fields in Values area: "Sum of Amount", "Average of Price", "Count of Orders"
// - These are the names to use in set-field-format

// Then format using discovered names:
{
  "tool": "excel_pivottable",
  "action": "set-field-format",
  "pivotTableName": "SalesPivot",
  "fieldName": "Sum of Amount",       // Use exact name from list-fields
  "numberFormat": "$#,##0.00"
}
```

**Common User Requests Translation:**

```
User says: "Format the sales numbers as currency"
  → set-field-format with fieldName="Sum of Sales" or "Total Sales"
  → numberFormat="$#,##0.00"

User says: "Make the PivotTable easier to read"
  → set-layout with layout="Tabular"
  → set-style with styleName="PivotStyleMedium2"

User says: "Make it look professional"
  → set-style with styleName="PivotStyleMedium2" or "PivotStyleMedium9"
  → Consider set-layout to "Tabular" if currently Compact

User says: "Format the totals row"
  → WARNING: Not supported persistently
  → Can use excel_range but warn it's temporary
  → Better: Use styles that format totals automatically

User says: "Add percentage formatting"
  → set-field-format with numberFormat="0.00%" or "0%"
  → Common for growth, margin, conversion rate fields
```

### 5. Batch Operations

```
Formatting 3+ columns/ranges
  → Use batch mode (begin_excel_batch → operations → commit_excel_batch)
  → Example: Format entire sales report (currency, percentages, dates, fonts, borders)

Single column/range formatting
  → Direct action call is fine

Complete workbook setup (create + format + validate)
  → Always use batch mode for consistency
```

---

## Error Response Structure (Critical for LLM Error Handling)

### Success Response Format

All operations return consistent success response:

```json
{
  "success": true,
  "sheetName": "Sales",  // Echo back for confirmation
  "rangeAddress": "D2:D100",  // Echo back for confirmation
  "message": "Number format applied successfully"
}
```

### Error Response Format

All operations return consistent error response:

```json
{
  "success": false,
  "errorMessage": "Sheet 'Sales' not found in workbook",
  "errorCode": "SHEET_NOT_FOUND"
}
```

### Common Error Codes

| Error Code | Meaning | User Action |
|------------|---------|-------------|
| `SHEET_NOT_FOUND` | Sheet name doesn't exist | List sheets first with excel_worksheet.list |
| `INVALID_RANGE` | Range address malformed | Use "A1:D10" format, check column/row exists |
| `TABLE_NOT_FOUND` | Table name doesn't exist | List tables with excel_table.list |
| `COLUMN_NOT_FOUND` | Column name not in table | List columns with excel_table.get-structured-reference |
| `PIVOTTABLE_NOT_FOUND` | PivotTable name doesn't exist | List pivot tables with excel_pivottable.list |
| `FIELD_NOT_FOUND` | Field name not in PivotTable | List fields with excel_pivottable.list-fields |
| `FIELD_NOT_IN_VALUES` | Field not in Values area | Check area with list-fields, only Values area supports number formats |
| `INVALID_FORMAT_CODE` | Number format code invalid | Use valid Excel format code (e.g., "$#,##0.00", not "currency") |
| `INVALID_VALIDATION_TYPE` | Validation type not recognized | Use: "List", "WholeNumber", "Decimal", "Date", "Time", "TextLength", "Custom" |
| `FILE_NOT_FOUND` | Excel file doesn't exist | Check excelPath is absolute path, file exists |
| `FILE_LOCKED` | Excel file open by another process | Close Excel file first |
| `BATCH_NOT_FOUND` | Batch ID doesn't exist | Check batchId from begin_excel_batch response |

### Get Operations - Return Value Structure

**get-number-formats** returns:
```json
{
  "success": true,
  "sheetName": "Sales",
  "rangeAddress": "A1:C3",
  "formats": [
    ["General", "$#,##0.00", "0.00%"],
    ["m/d/yyyy", "General", "General"],
    ["General", "#,##0", "@"]
  ],
  "rowCount": 3,
  "columnCount": 3
}
```

**get-font** returns:
```json
{
  "success": true,
  "sheetName": "Sales",
  "rangeAddress": "A1",
  "font": {
    "name": "Calibri",
    "size": 11,
    "bold": false,
    "italic": false,
    "color": 0,
    "underline": false,
    "strikethrough": false
  }
}
```

**get-background-color** returns:
```json
{
  "success": true,
  "sheetName": "Sales",
  "rangeAddress": "A1:E1",
  "colors": [
    [15128749, 15128749, 15128749, 15128749, 15128749]
  ],
  "rowCount": 1,
  "columnCount": 5
}
```

**get-borders** returns:
```json
{
  "success": true,
  "sheetName": "Sales",
  "rangeAddress": "A1",
  "borders": {
    "top": { "style": "Continuous", "weight": "Thin", "color": 0 },
    "bottom": { "style": "Continuous", "weight": "Thin", "color": 0 },
    "left": { "style": "None", "weight": "Hairline", "color": 0 },
    "right": { "style": "None", "weight": "Hairline", "color": 0 }
  }
}
```

**get-alignment** returns:
```json
{
  "success": true,
  "sheetName": "Sales",
  "rangeAddress": "A1:E1",
  "alignment": {
    "horizontal": "Center",
    "vertical": "Bottom",
    "wrapText": false,
    "indent": 0,
    "orientation": 0
  }
}
```

**get-validation** returns:
```json
{
  "success": true,
  "sheetName": "Sales",
  "rangeAddress": "F2:F100",
  "validation": {
    "type": "List",
    "operator": "Between",
    "formula1": "Active,Inactive",
    "formula2": null,
    "ignoreBlank": true,
    "showInputMessage": false,
    "inputTitle": null,
    "inputMessage": null,
    "showErrorAlert": true,
    "errorStyle": "Stop",
    "errorTitle": "Invalid Entry",
    "errorMessage": "Please select from dropdown"
  }
}
// If no validation exists:
{
  "success": true,
  "sheetName": "Sales",
  "rangeAddress": "A1",
  "validation": null
}
```

---

## Breaking Changes from Current API

### ✅ **Acceptable Breaking Changes**

1. **RangeNumberFormatResult** - New result type (previously planned but not implemented)
2. **Font/Border/Alignment enums** - New types for better type safety
3. **ValidationRule class** - Comprehensive validation configuration (cleaner than multiple parameters)
4. **Table formatting methods** - New methods, no existing API to break

### ❌ **No Breaking Changes To**

1. **Existing RangeCommands Phase 1** - All value/formula/clear/copy operations unchanged
2. **Existing TableCommands** - All lifecycle/filter/sort operations unchanged
3. **Result types** - Existing ResultBase, OperationResult, RangeValueResult unchanged

---

## Implementation Strategy

### Phase 2A: Number Formatting (Priority 1)
**Timeline:** 2-3 days

**Core Implementation:**
- ✅ Add `GetNumberFormatsAsync`, `SetNumberFormatAsync`, `SetNumberFormatsAsync` to IRangeCommands
- ✅ Implement in RangeCommands.cs using Excel COM Range.NumberFormat
- ✅ Add `NumberFormatPresets` static class
- ✅ Add `RangeNumberFormatResult` type
- ✅ Add table methods: `GetColumnNumberFormatAsync`, `SetColumnNumberFormatAsync`

**MCP Server:**
- ✅ Add actions to excel_range tool: `get-number-formats`, `set-number-format`, `set-number-formats`
- ✅ Add actions to excel_table tool: `get-column-number-format`, `set-column-number-format`

**Tests:**
- ✅ Currency format tests ($#,##0.00)
- ✅ Percentage format tests (0.00%)
- ✅ Date format tests (m/d/yyyy)
- ✅ Custom format tests
- ✅ Bulk format tests (2D array)
- ✅ Table column format tests

**Why First:** Most commonly requested by users, simplest to implement, highest ROI.

---

### Phase 2B: Visual Formatting (Priority 2)
**Timeline:** 3-4 days

**Core Implementation:**
- ✅ Add 14 font/color/border/alignment methods to IRangeCommands
- ✅ Implement using Excel COM Range.Font, Range.Interior, Range.Borders, alignment properties
- ✅ Add supporting types: FontOptions, BorderOptions, AlignmentOptions, enums
- ✅ Add result types: RangeFontResult, RangeColorResult, RangeBorderResult, RangeAlignmentResult
- ✅ Add table methods for column formatting

**MCP Server:**
- ✅ Add 14+ actions to excel_range tool
- ✅ Add 4 formatting actions to excel_table tool

**Tests:**
- ✅ Font tests (bold, italic, size, color)
- ✅ Background color tests (set, get, clear)
- ✅ Border tests (styles, weights)
- ✅ Alignment tests (horizontal, vertical, wrap, indent, rotation)
- ✅ Auto-fit tests
- ✅ Table column visual format tests

**Why Second:** High user demand, professional appearance, builds on number formatting foundation.

---

### Phase 2C: Data Validation (Priority 3)
**Timeline:** 2-3 days

**Core Implementation:**
- ✅ Add 4 validation methods to IRangeCommands
- ✅ Implement using Excel COM Range.Validation
- ✅ Add ValidationRule class with all types/operators
- ✅ Add ValidationOperator, ValidationType, ValidationAlertStyle enums
- ✅ Add RangeValidationResult type
- ✅ Add table methods for column validation

**MCP Server:**
- ✅ Add 4 actions to excel_range tool
- ✅ Add 2 actions to excel_table tool

**Tests:**
- ✅ List validation tests (dropdown)
- ✅ Number validation tests (whole, decimal, between)
- ✅ Date/time validation tests
- ✅ Text length validation tests
- ✅ Custom formula validation tests
- ✅ Error alert tests
- ✅ Table column validation tests

**Why Third:** Data quality critical, complex API, requires careful Excel COM handling.

---

### Phase 2D: CLI Implementation (Priority 4)
**Timeline:** 2 days

**CLI Commands:**
- ✅ `range-get-number-formats`, `range-set-number-format`, `range-set-number-formats`
- ✅ `range-set-font`, `range-set-background-color`, `range-set-borders`, `range-set-alignment`
- ✅ `range-auto-fit-columns`, `range-auto-fit-rows`
- ✅ `range-add-validation`, `range-get-validation`, `range-remove-validation`
- ✅ `table-set-column-number-format`, `table-add-column-validation`, etc.

**Documentation:**
- ✅ Update README.md
- ✅ Update copilot instructions

---

## Testing Strategy

### Unit Tests
- Number format code validation
- RGB color conversion
- Enum mapping (BorderStyle, ValidationOperator, etc.)
- ValidationRule validation

### Integration Tests (Requires Excel)

**Number Formatting:**
- Set currency format, read back, verify
- Set custom format, verify rendering
- Bulk format 2D array, verify each cell
- Table column format, verify data cells only (not headers)

**Visual Formatting:**
- Set font properties, verify each property
- Set background color, verify RGB components
- Set borders, verify style/weight/color
- Set alignment, verify horizontal/vertical/wrap
- Auto-fit, verify column widths adjusted

**Data Validation:**
- Add list validation, verify dropdown appears
- Add number range validation, test invalid input blocked
- Add date validation, verify date picker
- Add custom formula validation, verify evaluation
- Remove validation, verify dropdown removed

**Table Operations:**
- Format table column, verify only data cells affected
- Add validation to table column, verify auto-expands with new rows
- Auto-fit table columns, verify all columns sized

---

## Success Criteria

**Phase 2A - Number Formatting:**
- [ ] All 3 range number format methods implemented and tested
- [ ] NumberFormatPresets class with 20+ common patterns
- [ ] 2 table methods working
- [ ] MCP actions functional
- [ ] 10+ integration tests passing

**Phase 2B - Visual Formatting:**
- [ ] All 14 visual formatting methods implemented and tested
- [ ] Font, border, alignment options working
- [ ] RGB color handling correct
- [ ] Auto-fit functioning
- [ ] 20+ integration tests passing

**Phase 2C - Data Validation:**
- [ ] All 4 validation methods implemented and tested
- [ ] All validation types working (List, Number, Date, TextLength, Custom)
- [ ] Error alerts and input messages functional
- [ ] Table column validation auto-expanding
- [ ] 15+ integration tests passing

**Phase 2D - CLI:**
- [ ] All CLI commands implemented
- [ ] Documentation complete
- [ ] CLI tests passing

**Overall:**
- [ ] All 21 new range methods working
- [ ] All 8 new table methods working
- [ ] 95%+ test coverage
- [ ] MCP Server integration complete
- [ ] Documentation comprehensive
- [ ] Zero regression in existing features

---

## Future Enhancements (Phase 3+)

### Conditional Formatting (Complex)
- Rule-based formatting (data bars, color scales, icon sets)
- Formula-based rules
- Manage existing conditional formats

### Cell Merge/Unmerge
- Merge cells in range
- Unmerge cells
- Check merge status

### Cell Protection
- Lock/unlock cells
- Protect worksheet with password
- Check protection status

### Advanced Formatting
- Patterns (diagonal lines, dots, etc.)
- Gradient fills
- Custom number formats with conditions

These can be considered in future iterations based on user demand and complexity vs value analysis.

---

## Summary

This specification provides a comprehensive, LLM-first approach to Excel formatting and data validation:

**Key Benefits:**
1. **Professional Output** - LLMs can create polished, formatted reports
2. **Data Quality** - Validation ensures data integrity
3. **User Experience** - Dropdowns and input messages guide users
4. **Consistency** - Unified API across ranges and tables
5. **Performance** - Batch operations for efficiency

**Total API Surface:**
- 21 new range methods (number formatting: 3, visual: 14, validation: 4)
- 8 new table methods (number formatting: 2, visual: 4, validation: 2)
- 0 new PivotTable methods (already implemented in Phase 1)
- 29+ new MCP actions (range + table only)

**Implementation Timeline:**
- Phase 2A (Number): 2-3 days
- Phase 2B (Visual): 3-4 days
- Phase 2C (Validation): 2-3 days
- Phase 2D (CLI): 2 days
- **Total: 9-12 days for complete Phase 2**

This represents a significant enhancement to ExcelMcp's capabilities, enabling LLMs to create production-quality Excel workbooks programmatically.
