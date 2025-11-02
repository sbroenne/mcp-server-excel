# Excel Formatting Best Practices

**Professional formatting standards for Excel 2019+ tables, ranges, and reports.**

> **✅ BUILT-IN CELL STYLES NOW SUPPORTED!**
> 
> Use `excel_range` with `action: 'set-style'` to apply built-in Excel styles.
> These are FASTER, more CONSISTENT, and more PROFESSIONAL than manual formatting.
> 
> **Recommended Workflow:**
> 1. **First:** Try built-in styles (Heading 1, Accent1, Total, Good/Bad/Neutral, etc.)
> 2. **Only if needed:** Use manual formatting via `format-range` action

## Excel Built-in Cell Styles (USE THESE FIRST!)

**Why built-in styles?**
- ✅ **Faster:** 1 line of code vs 5-10 lines manual formatting
- ✅ **Consistent:** Same style = same look everywhere
- ✅ **Theme-aware:** Auto-adjust when theme changes
- ✅ **Professional:** Tested and polished by Microsoft
- ✅ **Maintainable:** Change definition once, all cells update

### How to Apply Built-in Styles

```javascript
// MCP Server - RECOMMENDED
excel_range(action: 'set-style', excelPath: 'report.xlsx', sheetName: 'Sheet1', rangeAddress: 'A1', styleName: 'Heading 1')

// Manual formatting - use only when styles don't meet needs
excel_range(action: 'format-range', excelPath: 'report.xlsx', sheetName: 'Sheet1', rangeAddress: 'A1', bold: true, fontSize: 14, fontColor: '#0000FF')
```

### Available Built-in Styles (47+ styles)

#### **Good, Bad and Neutral** (Status Indicators)
```
Good           - Green background, dark green text (✅ positive, completed)
Bad            - Red background, dark red text (❌ negative, errors)
Neutral        - Orange background, dark orange text (⚠️ warnings, pending)
```

**Use for:** KPIs, status indicators, conditional formatting, traffic lights

#### **Data and Model** (Cell Purpose Markers)
```
Input          - Orange background (👤 user input cells)
Calculation    - Bold, orange background (➗ formula cells)
Output         - Bold, gray background (📊 read-only results)
Check Cell     - Bold, white text on green background (✓ validation passed)
Linked Cell    - Bold, orange text (🔗 linked to other sheets)
Note           - Yellow background (📝 annotations, comments)
Explanatory Text - Italic (📄 instructions, help text)
Warning Text   - Red text (⚠️ warnings, errors)
```

**Use for:** Data entry forms, workbooks with formulas, templates

#### **Titles and Headings** (Document Structure)
```
Title          - 18pt Cambria, blue (📰 main report title)
Heading 1      - 15pt Calibri, bold, blue (📑 major sections)
Heading 2      - 13pt Calibri, bold, blue (📄 subsections)
Heading 3      - 11pt Calibri, bold, blue (📝 minor headings)
Heading 4      - 11pt Calibri, bold, blue (📌 smallest heading)
```

**Use for:** Report titles, section headers, table captions

#### **Themed Cell Styles** (Professional Accents)
```
Accent1        - Full theme color 1 (typically blue), white text
Accent2        - Full theme color 2 (typically orange), white text
Accent3        - Full theme color 3 (typically gray), white text
Accent4        - Full theme color 4, white text
Accent5        - Full theme color 5, white text
Accent6        - Full theme color 6, white text

60% - Accent1  - Medium intensity, white text
60% - Accent2  - Medium intensity, white text
... (Accent3-6 similar)

40% - Accent1  - Light background, dark text
40% - Accent2  - Light background, dark text
... (Accent3-6 similar)

20% - Accent1  - Very light background, dark text
20% - Accent2  - Very light background, dark text
... (Accent3-6 similar)
```

**Use for:** Table headers, alternating rows, highlights

**Note:** Default Office theme: Accent1=blue, Accent2=orange, Accent3=gray

#### **Number Format Styles** (Quick Number Formatting)
```
Currency       - $#,##0.00 (with dollar sign and 2 decimals)
Currency [0]   - $#,##0 (no decimals)
Comma          - #,##0.00 (thousands separator, 2 decimals)
Comma [0]      - #,##0 (no decimals)
Percent        - 0.00% (2 decimal percentage)
```

**Use for:** Quick number formatting without manual format codes

#### **Other Essential Styles**
```
Normal         - 11pt Aptos (Excel 2019+), no formatting (default cell)
Total          - Bold, top border (📊 totals/subtotals rows)
```

---

## Built-in Style Recommendations by Use Case

### **Financial Reports**
```
Title:           Title or Heading 1
Section Headers: Heading 2
Column Headers:  Accent1 (full blue, white text)
Data:            Normal with Currency or Comma [0] style
Input Cells:     Input (orange - user enters data)
Calculated:      Calculation (orange - formulas)
Subtotals:       Total style
Grand Total:     Total style (consider double top border manually)
Assumptions:     Explanatory Text or Note (yellow)
```

**Example:**
```csharp
excel_range(action: 'set-style', rangeAddress: 'A1', styleName: 'Title')              // Q4 2024 Financial Report
excel_range(action: 'set-style', rangeAddress: 'A3', styleName: 'Heading 1')          // Revenue Statement
excel_range(action: 'set-style', rangeAddress: 'A5:E5', styleName: 'Accent1')         // Column headers
excel_range(action: 'set-style', rangeAddress: 'B6:E10', styleName: 'Currency [0]')   // Dollar amounts (no decimals)
excel_range(action: 'set-style', rangeAddress: 'B10:E10', styleName: 'Total')         // Totals row
```

### **Sales Dashboards**
```
Dashboard Title: Title (18pt)
KPI Headers:     Heading 2 or Accent1
Positive Trend:  Good (green)
Negative Trend:  Bad (red)
Neutral/Flat:    Neutral (orange)
Data Tables:     20% - Accent1 for headers (light blue)
Data:            Normal with Comma [0]
```

**Example:**
```csharp
excel_range(action: 'set-style', rangeAddress: 'A1', styleName: 'Title')                    // Sales Dashboard
excel_range(action: 'set-style', rangeAddress: 'A3', styleName: 'Heading 2')                // Q4 Performance
excel_range(action: 'set-style', rangeAddress: 'A5:D5', styleName: '20% - Accent1')         // Table headers (light)
if (salesGrowth > 0) excel_range(action: 'set-style', rangeAddress: 'B10', styleName: 'Good')      // Green
else if (salesGrowth < 0) excel_range(action: 'set-style', rangeAddress: 'B10', styleName: 'Bad')  // Red
```

### **Data Entry Forms**
```
Form Title:      Heading 1
Section Labels:  Heading 3 or Heading 4
Required Input:  Input (orange)
Optional Input:  20% - Accent1 (light blue)
Calculated:      Calculation (orange) or Output (gray)
Instructions:    Explanatory Text (italic)
Warnings:        Warning Text (red) or Bad
Validation OK:   Check Cell (green) or Good
```

**Example:**
```csharp
excel_range(action: 'set-style', rangeAddress: 'A1', styleName: 'Heading 1')              // Employee Information Form
excel_range(action: 'set-style', rangeAddress: 'A3', styleName: 'Heading 3')              // Personal Details
excel_range(action: 'set-style', rangeAddress: 'B5:B8', styleName: 'Input')               // User fills these (orange)
excel_range(action: 'set-style', rangeAddress: 'B10:B12', styleName: '20% - Accent1')     // Optional fields (light blue)
excel_range(action: 'set-style', rangeAddress: 'B15', styleName: 'Calculation')           // Formula (age from birthdate)
excel_range(action: 'set-style', rangeAddress: 'A20', styleName: 'Explanatory Text')      // Instructions
```

### **Project Reports**
```
Report Title:    Title
Major Sections:  Heading 1
Subsections:     Heading 2
Table Headers:   Accent1 or 40% - Accent1
Completed Tasks: Good (green)
Delayed/Issues:  Bad (red)
In Progress:     Neutral (orange)
Notes:           Note (yellow background)
```

### **Budget/Variance Reports**
```
Headers:         Heading 2 or Accent1
Actuals:         Normal with Currency
Budget:          20% - Accent1 (light background)
Positive Var:    Good (green)
Negative Var:    Bad (red)
Totals:          Total
Assumptions:     Explanatory Text or Note
```

---

## When Built-in Styles Are Not Enough

**Use manual formatting ONLY when:**
- ✅ Specific brand colors required (not Office theme colors)
- ✅ Very specific formatting not covered by 47+ built-in styles
- ✅ Custom charts/graphics
- ✅ One-off unique design

**For everything else, use built-in styles!**

---

## Manual Formatting Reference (Fallback Only)

**Only use these when built-in styles don't apply.**

### Excel 2019+ Defaults
- **Font:** Aptos, 11pt
- **Theme Colors:** #4472C4 (blue), #ED7D31 (orange), #A5A5A5 (gray)

### Common Number Format Codes

**Note:** Use built-in number format styles (Currency, Comma, Percent) when possible!
Only use these codes when you need custom formatting.

**Currency:**
```
$#,##0.00          // $1,234.56 (2 decimals)
$#,##0             // $1,235 (no decimals, rounded)
```

**Accounting (aligned decimals with parentheses for negatives):**
```
_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)
```

**Percentage:**
```
0%                 // 45% (no decimals)
0.0%               // 45.2% (1 decimal)
0.00%              // 45.23% (2 decimals)
```

**Date:**
```
m/d/yyyy           // 1/29/2025 (US format)
mm/dd/yyyy         // 01/29/2025 (leading zeros)
yyyy-mm-dd         // 2025-01-29 (ISO format, best for sorting)
mmm d, yyyy        // Jan 29, 2025 (readable)
```

**Number:**
```
#,##0              // 1,235 (thousands separator, no decimals)
#,##0.00           // 1,234.56 (2 decimals)
0.00               // 1234.56 (no thousands separator)
```

**Custom:**
```
[>0]#,##0;[Red](#,##0);"-"     // Green positive, red negative in parens, dash for zero
```

### Manual Table Formatting (When Built-in Styles Don't Apply)

**Headers (Row 1):**
- Font: Bold, 11pt
- Background: Theme Blue (#4472C4) or Light Blue (#D6DCE4)
- Text Color: White (with dark background) or Black (with light background)
- Alignment: Center horizontal, center vertical
- Borders: All borders (thin)
- **Freeze panes** after header row

**Data Rows:**
- Font: Regular, 11pt
- Alternating row colors (optional): White and Light Gray (#F2F2F2)
- Alignment: Numbers right, text left, dates right
- Borders: Thin bottom border or no borders (cleaner look)

**Totals Row:**
- Font: Bold, 11pt
- Background: Light Blue (#D6DCE4) or same as header
- Border: Top border (medium weight)
- Alignment: Right (for totals)

**Column Widths:**
- Auto-fit for most columns
- Minimum 10 characters for readability
- Maximum 50 characters (wrap text if needed)

---

## Excel Table Styles (For Excel Tables/ListObjects)

**When using Excel Tables** (Insert → Table), apply these built-in table styles:

**Most Common (Data Tables):**
- **Table Style Medium 2** - Blue header, white/light blue alternating rows

**Financial Reports:**
- **Table Style Light 9** - Minimal borders, professional
- Use accounting number format
- Bold totals with top border

**Dashboards:**
- **Table Style Medium 7** - Colorful, modern
- Larger fonts (12pt)
- More spacing (row height 18pt)

**To Apply:** Select table → Table Design tab → Table Styles

---

## Conditional Formatting (Advanced)

**Negative Numbers:**
```
Format: Font color Red (#C55A11)
Or: Add (parentheses) with accounting format
```

**Top/Bottom Values:**
```
Top 10: Light green fill (#C6E0B4), dark green text (#70AD47)
Bottom 10: Light red fill (#F4B084), dark red text (#C55A11)
```

**Data Bars:**
```
Use for quick visual comparison
Color: Theme blue (#4472C4) or green (#70AD47)
```

**Color Scales:**
```
3-Color: Red → Yellow → Green
2-Color: White → Blue (less distracting)
```

**Note:** For simple status indicators, use Good/Bad/Neutral styles instead!

---

## Layout Best Practices (Universal)

**Font Size:**
- Minimum: **10pt** (preferably 11pt)
- Headers: 11pt bold (or 12pt)

**Color Contrast:**
- Text on background: **4.5:1 ratio minimum**
- White text on theme blue (#4472C4): ✅ Good
- Black text on yellow (#FFC000): ✅ Good
- Gray text on light gray: ❌ Bad (poor contrast)

**Colorblind-Safe:**
- ✅ Use blue/orange instead of red/green
- ✅ Add icons or patterns, not just color
- ❌ Don't use red/green only for positive/negative

**Note:** Built-in Good/Bad/Neutral styles are already accessible!

---

## Accessibility Guidelines

**Alignment:**
- **Text:** Left-aligned
- **Numbers:** Right-aligned (easier to compare)
- **Headers:** Centered (optional) or left-aligned
- **Dates:** Right-aligned (treat as numbers)

**Row Heights:**
- Default: 15pt
- More spacing: 18pt (better readability)
- Headers: Can be taller (20-22pt)

**Column Widths:**
- Use auto-fit as starting point
- Adjust manually for visual balance
- Keep consistent across similar columns

**Merge Cells:**
- ❌ **Avoid!** Breaks sorting, filtering, formulas
- ✅ Use "Center Across Selection" instead (Format Cells → Alignment)

**Borders:**
- Headers: All borders (thin)
- Data: Bottom border only (cleaner) or no borders
- Totals: Top border (medium), bottom border (double for final total)
- Avoid heavy borders (too busy)

**Note:** Total style already includes top border!

---

## Summary: Recommended Approach

❌ **Don't:**
- Use Comic Sans or decorative fonts
- Use font sizes < 10pt (hard to read)
- Use bright, saturated colors (#FF0000, #00FF00)
- Center-align numbers (hard to compare)
- Merge cells (breaks functionality)
- Use red/green only (colorblind issue)
- Mix currency symbols ($1,000 and 1000 USD in same column)

✅ **Do:**
- **Use built-in cell styles first!** (Heading 1, Accent1, Total, etc.)
- Use Excel 2019+ default font (Aptos, 11pt)
- Use 11pt as standard size
- Use theme colors (#4472C4, #ED7D31, etc.)
- Right-align numbers for easy comparison
- Use Center Across Selection instead of merge
- Add icons/patterns for colorblind accessibility
- Be consistent with number formats

### **Quick Formatting Workflow**

**For Professional Documents (USE BUILT-IN STYLES):**
1. **Apply styles to structure:**
   - Title → `Title` or `Heading 1`
   - Sections → `Heading 2`, `Heading 3`
2. **Apply styles to table headers:**
   - Column headers → `Accent1` (full blue)
   - Or lighter headers → `40% - Accent1`
3. **Mark cell purposes:**
   - User input → `Input` (orange)
   - Formulas → `Calculation` (orange, bold)
   - Results → `Output` (gray)
4. **Apply number formats:**
   - Currency → `Currency` or `Currency [0]` style
   - Numbers → `Comma` or `Comma [0]` style
   - Percentages → `Percent` style
5. **Apply totals:**
   - Totals row → `Total` style (bold, top border)
6. **Apply status (if needed):**
   - Positive/Complete → `Good` (green)
   - Negative/Error → `Bad` (red)
   - Warning/Pending → `Neutral` (orange)

**Result:** Professional, consistent, theme-aware document in 2-3 minutes using built-in styles!

---

## Examples: Built-in Styles in Action

### **Financial Report**
```
Revenue: $#,##0  (no decimals for large numbers)
EBITDA: $#,##0  (no decimals)
Margin: 0.0%  (1 decimal percentage)
```

**Sales Reports:**
```
Units: #,##0  (thousands separator)
Price: $#,##0.00  (2 decimals)
Total: $#,##0.00  (2 decimals)
```

**Project Tracking:**
```
Progress: 0%  (no decimals)
Budget: $#,##0  (no decimals)
Due Date: mmm d, yyyy  (Jan 29, 2025)
```

**Default to:** Standard theme colors, 11pt Aptos/Calibri, proper alignment, and clear number formats.

## Excel Built-in Cell Styles (Recommended!)

**Why use built-in styles?** Faster, consistent, theme-aware, professional.

### How to Apply (COM)

```csharp
// C# COM Interop
range.Style = "Heading 1";  // Note: Space in name!

// MCP Server
excel_range(action: 'set-style', rangeAddress: 'A1', styleName: 'Heading 1')
```

### Style Categories and Names

#### **Good, Bad and Neutral**
```
Good           - Green background, dark green text (positive values, completed)
Bad            - Red background, dark red text (negative values, errors)
Neutral        - Orange background, dark orange text (warnings, pending)
```

#### **Data and Model**
```
Calculation    - Bold, orange background (calculated/formula cells)
Check Cell     - Bold, white text on green background (validation/checks)
Explanatory Text - Italic, left-aligned (notes, instructions)
Input          - Orange background (user input cells)
Linked Cell    - Bold, orange text (cells linked to other sheets)
Note           - Yellow background (comments, annotations)
Output         - Bold, gray background (formula results, read-only)
Warning Text   - Red text (warnings, error messages)
```

#### **Titles and Headings**
```
Title          - 18pt Cambria, blue (main report title)
Heading 1      - 15pt Calibri, bold, blue (major sections)
Heading 2      - 13pt Calibri, bold, blue (subsections)
Heading 3      - 11pt Calibri, bold, blue (minor headings)
Heading 4      - 11pt Calibri, bold, blue (smallest heading)
```

#### **Themed Cell Styles (20% variations)**
```
20% - Accent1  - Very light theme color 1 background
20% - Accent2  - Very light theme color 2 background
20% - Accent3  - Very light theme color 3 background
20% - Accent4  - Very light theme color 4 background
20% - Accent5  - Very light theme color 5 background
20% - Accent6  - Very light theme color 6 background
```

#### **Themed Cell Styles (40% variations)**
```
40% - Accent1  - Light theme color 1 background
40% - Accent2  - Light theme color 2 background
40% - Accent3  - Light theme color 3 background
40% - Accent4  - Light theme color 4 background
40% - Accent5  - Light theme color 5 background
40% - Accent6  - Light theme color 6 background
```

#### **Themed Cell Styles (60% variations)**
```
60% - Accent1  - Medium theme color 1 background, white text
60% - Accent2  - Medium theme color 2 background, white text
60% - Accent3  - Medium theme color 3 background, white text
60% - Accent4  - Medium theme color 4 background, white text
60% - Accent5  - Medium theme color 5 background, white text
60% - Accent6  - Medium theme color 6 background, white text
```

#### **Themed Cell Styles (Accent variations)**
```
Accent1        - Full theme color 1 background, white text
Accent2        - Full theme color 2 background, white text
Accent3        - Full theme color 3 background, white text
Accent4        - Full theme color 4 background, white text
Accent5        - Full theme color 5 background, white text
Accent6        - Full theme color 6 background, white text
```

**Note:** Accent1 is typically blue, Accent2 orange, Accent3 gray in default Office theme.

#### **Number Format Styles**
```
Comma          - #,##0.00 format with thousands separator
Comma [0]      - #,##0 format (no decimals)
Currency       - $#,##0.00 format
Currency [0]   - $#,##0 format (no decimals)
Percent        - 0.00% format (2 decimal percentage)
```

#### **Other Built-in Styles**
```
Normal         - Default cell style (11pt Calibri, no formatting)
Total          - Bold, top border (for totals rows)
```

### Style Recommendations by Use Case

#### **Financial Reports:**
```
Title:          Title or Heading 1
Section Headers: Heading 2
Column Headers: Accent1 (full, white text)
Data:           Normal with Comma or Currency number format
Subtotals:      Total style
Grand Total:    Total style with double top border
Input Cells:    Input (orange background)
Calculated:     Calculation or Output
```

#### **Sales Dashboards:**
```
Dashboard Title: Title (18pt)
KPI Headers:    Heading 2 or Accent1
Positive Trend: Good (green)
Negative Trend: Bad (red)
Neutral/Flat:   Neutral (orange)
Data Tables:    Normal with 20% - Accent1 for headers
```

#### **Data Entry Forms:**
```
Form Title:     Heading 1
Section Labels: Heading 3 or Heading 4
Required Input: Input (orange background)
Optional Input: 20% - Accent1 (light blue)
Calculated:     Calculation (orange) or Output (gray)
Instructions:   Explanatory Text (italic)
Warnings:       Warning Text (red) or Bad
Validation OK:  Check Cell (green) or Good
```

#### **Project Reports:**
```
Report Title:   Title
Major Sections: Heading 1
Subsections:    Heading 2
Table Headers:  Accent1 or 40% - Accent1
Completed:      Good (green)
Delayed/Issue:  Bad (red)
In Progress:    Neutral (orange)
Notes:          Note (yellow background)
```

#### **Budget/Variance Reports:**
```
Headers:        Heading 2 or Accent1
Actuals:        Normal with Currency
Budget:         20% - Accent1
Variance:       Good (positive) or Bad (negative)
Totals:         Total
Assumptions:    Explanatory Text or Note
```

### When to Use Styles vs Manual Formatting

**Use Built-in Styles When:**
- ✅ Creating professional reports/forms
- ✅ Want theme consistency
- ✅ Need quick, standard formatting
- ✅ Document will be shared/reused
- ✅ Using common patterns (headers, totals, input cells)

**Use Manual Formatting When:**
- ✅ Specific brand colors required (not theme colors)
- ✅ One-off custom design
- ✅ Very specific formatting not covered by styles
- ✅ Charts/graphics with custom colors

**Best Practice:** Start with built-in styles, customize only when necessary.

### Quick Style Application Examples

**Financial Report Header:**
```csharp
// Apply Heading 1 to title
excel_range(action: 'set-style', rangeAddress: 'A1', styleName: 'Heading 1')

// Apply Accent1 to column headers
excel_range(action: 'set-style', rangeAddress: 'A2:E2', styleName: 'Accent1')

// Apply Total to totals row
excel_range(action: 'set-style', rangeAddress: 'A10:E10', styleName: 'Total')
```

**Data Entry Form:**
```csharp
// Mark input cells
excel_range(action: 'set-style', rangeAddress: 'B5:B10', styleName: 'Input')  // Orange background

// Mark calculated cells
excel_range(action: 'set-style', rangeAddress: 'B15:B20', styleName: 'Calculation')  // Orange with bold

// Add instructions
excel_range(action: 'set-style', rangeAddress: 'A1', styleName: 'Explanatory Text')  // Italic
```

**Dashboard KPIs:**
```csharp
// Good/bad/neutral for metrics
if (value > target) range.Style = "Good";       // Green
else if (value < target * 0.9) range.Style = "Bad";   // Red
else range.Style = "Neutral";  // Orange
```

**Common Mistake to Avoid:**
```csharp
// ❌ WRONG: No space in style name
range.Style = "Heading1";  // Error!

// ✅ CORRECT: Space in multi-word styles
range.Style = "Heading 1";  // Works!

// ✅ CORRECT: No space in single-word or hyphenated
range.Style = "Total";
range.Style = "20% - Accent1";
```
