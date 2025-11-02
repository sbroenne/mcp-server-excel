# BEFORE FORMATTING RANGES - GATHER THIS INFO

REQUIRED:
☐ Excel file path
☐ Worksheet name
☐ Range address (e.g., 'A1:D10', 'SalesData')
☐ **Purpose of formatting** (header, total, input, status, data, etc.)

**STEP 1: CHECK IF BUILT-IN STYLE WORKS (99% of cases)**

Ask user what they're formatting:
- Headers/Titles? → Try 'Heading 1', 'Heading 2', 'Title', 'Accent1'
- Totals/Subtotals? → Try 'Total'
- Input cells? → Try 'Input'
- Calculated cells? → Try 'Calculation' or 'Output'
- Status indicators? → Try 'Good', 'Bad', 'Neutral'
- Notes/Comments? → Try 'Note', 'Explanatory Text', 'Warning Text'
- Data tables? → Try '20% - Accent1', '40% - Accent1', '60% - Accent1'

**STEP 2: ONLY IF NO BUILT-IN STYLE FITS, gather custom formatting details:**

FONT (only if not using built-in style):
☐ Font name (Arial, Calibri, Times New Roman)
☐ Font size (8, 10, 11, 12, 14, 16)
☐ Bold, italic, underline
☐ Font color (#RRGGBB hex code, e.g., #FF0000 for red)

FILL (only if not using built-in style):
☐ Fill color (#RRGGBB hex code, e.g., #FFFF00 for yellow)

BORDERS (only if not using built-in style):
☐ Border style (none, continuous, dash, dot, double)
☐ Border weight (hairline, thin, medium, thick)
☐ Border color (#RRGGBB hex code)

ALIGNMENT (only if not using built-in style):
☐ Horizontal (left, center, right, justify)
☐ Vertical (top, center, bottom)
☐ Wrap text (true/false)

NUMBER FORMAT (can combine with styles):
☐ Format code (e.g., '$#,##0.00', '0.00%', 'mm/dd/yyyy')

**WORKFLOW:**
1. **ASK user what they're formatting** (purpose, not colors)
2. **RECOMMEND built-in style** based on purpose (see style_names.md completions)
3. **ONLY IF user needs custom formatting**, gather the details above
4. **Use set-style action** for built-in styles (faster, theme-aware)
5. **Use format-range action** only for custom formatting (one-off, brand-specific)
