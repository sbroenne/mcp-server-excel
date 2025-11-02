# Excel Formatting Best Practices

**Professional formatting standards for Excel tables, ranges, and reports.**

## Excel Defaults (Use These Unless You Have a Reason Not To)

**Fonts:**
- Excel 2019+: **Aptos** (default)
- Excel 2007-2016: **Calibri** (default)
- Font Size: **11pt** (default)
- Alternative: Arial 10pt (more compact)

**Colors (Excel Theme Colors with Hex Codes):**
- Blue (headers): **#4472C4**
- Orange (accent): **#ED7D31**
- Gray (subtle): **#A5A5A5**
- Light Blue (alt rows): **#D6DCE4**
- Dark Blue (totals): **#2F5496**
- Green (positive): **#70AD47**
- Red (negative): **#C55A11**

**Why theme colors?** They automatically adjust if user changes workbook theme.

## Common Number Format Codes

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

## Professional Table Formatting

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

## Table Style Quick Reference

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

## Conditional Formatting Rules

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

## Accessibility Guidelines

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

## Layout Best Practices

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

## Common Mistakes to Avoid

❌ **Don't:**
- Use Comic Sans or decorative fonts
- Use font sizes < 10pt (hard to read)
- Use bright, saturated colors (#FF0000, #00FF00)
- Center-align numbers (hard to compare)
- Merge cells (breaks functionality)
- Use red/green only (colorblind issue)
- Mix currency symbols ($1,000 and 1000 USD in same column)

✅ **Do:**
- Use Excel default fonts (Aptos, Calibri)
- Use 11pt as standard size
- Use theme colors (#4472C4, #ED7D31, etc.)
- Right-align numbers for easy comparison
- Use Center Across Selection instead of merge
- Add icons/patterns for colorblind accessibility
- Be consistent with number formats

## Quick Formatting Workflow

**For Professional Tables:**
1. Apply Table Style Medium 2 (or similar)
2. Bold headers, consider background color
3. Right-align number columns
4. Apply number formats ($#,##0.00, 0%, etc.)
5. Auto-fit column widths, adjust as needed
6. Freeze header row
7. Bold totals row with top border

**Result:** Clean, professional, accessible table in 2-3 minutes.

## Format Code Examples by Use Case

**Financial Statements:**
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
