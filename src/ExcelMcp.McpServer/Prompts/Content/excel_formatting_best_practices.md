# Excel Formatting Best Practices

**When formatting Excel ranges, ALWAYS prefer built-in cell styles over manual formatting.**

## Why Built-in Styles First?

**Built-in styles are:**
- ✅ **Faster** - 1 action vs 5-10 manual formatting calls
- ✅ **Consistent** - Same style = same look everywhere
- ✅ **Theme-aware** - Auto-adjust when workbook theme changes
- ✅ **Professional** - Tested and polished by Microsoft
- ✅ **Maintainable** - Change style definition once, all cells update

**Manual formatting is:**
- ⚠️ **One-off** - Doesn't update when theme changes
- ⚠️ **Inconsistent** - Easy to apply slightly different formats
- ⚠️ **Slower** - Multiple API calls for font, color, borders, etc.

## Decision Guide: Styles vs Manual Formatting

### ✅ Use Built-in Styles When:
- Creating professional reports/dashboards
- Want theme consistency across workbook
- Using common patterns (headers, totals, input cells, status indicators)
- Document will be shared/reused
- Need quick, standard formatting

### ✅ Use Manual Formatting When:
- Specific brand colors required (not theme colors)
- One-off custom design
- Very specific formatting not covered by styles
- Charts/graphics with custom colors

**Best Practice:** Start with built-in styles, customize only when necessary.

## Style Recommendations by Use Case

### Financial Reports
- **Title**: Heading 1 (15pt blue, bold)
- **Column Headers**: Accent1 (blue background, white text)
- **Input Cells**: Input (orange background - user enters data)
- **Calculated Cells**: Calculation (orange background, bold - formulas)
- **Subtotals/Totals**: Total (bold, top border)
- **Data**: Normal with Currency or Comma [0] number format

### Sales Dashboards
- **Dashboard Title**: Title (18pt)
- **KPI Headers**: Heading 2 or Accent1
- **Positive Metrics**: Good (green background, dark green text)
- **Negative Metrics**: Bad (red background, dark red text)
- **Neutral/Warning**: Neutral (orange background, dark orange text)
- **Data Tables**: 20% - Accent1 for headers (light blue)

### Data Entry Forms
- **Form Title**: Heading 1
- **Required Input**: Input (orange background)
- **Optional Input**: 20% - Accent1 (light blue)
- **Calculated**: Calculation (orange, bold) or Output (gray)
- **Instructions**: Explanatory Text (italic)
- **Warnings**: Warning Text (red) or Bad
- **Validation OK**: Check Cell (green) or Good

### Project Reports
- **Report Title**: Title
- **Major Sections**: Heading 1
- **Table Headers**: Accent1 or 40% - Accent1
- **Completed**: Good (green)
- **Delayed/Issue**: Bad (red)
- **In Progress**: Neutral (orange)
- **Notes**: Note (yellow background)

## Common Mistakes

❌ **Using manual formatting first**
- Try built-in styles before writing custom formatting code

❌ **Inconsistent formatting**
- Same purpose = same style (e.g., all headers use Accent1)

❌ **Forgetting spaces in style names**
- ❌ 'Heading1' → Error!
- ✅ 'Heading 1' → Works!

❌ **Reinventing standard styles**
- Don't manually create bold + blue + 14pt when Heading 2 exists

❌ **Not using status styles**
- Good/Bad/Neutral are perfect for KPIs, status indicators, traffic lights
- More professional than custom red/green/yellow

## Server-Specific Behavior

**Style application:**
- Use xcel_range with ction: 'set-style' to apply built-in styles
- Use ction: 'format-range' only for custom formatting
- Batch mode recommended for formatting multiple ranges (3+ operations)

**Available styles:**
- See completions for styleName parameter (47+ built-in styles)
- Use elicitation ange_formatting to gather formatting requirements

**Workflow:**
1. Check if built-in style meets needs (99% of cases)
2. Apply style with single set-style action
3. Only use ormat-range for brand-specific colors or one-off designs

## Quick Examples

**Apply heading:**
```javascript
excel_range(action: 'set-style', rangeAddress: 'A1', styleName: 'Heading 1')
```

**Format financial report headers:**
```javascript
// Headers with built-in style (RECOMMENDED)
excel_range(action: 'set-style', rangeAddress: 'A2:E2', styleName: 'Accent1')

// Totals row
excel_range(action: 'set-style', rangeAddress: 'A10:E10', styleName: 'Total')
```

**Mark input cells in data entry form:**
```javascript
excel_range(action: 'set-style', rangeAddress: 'B5:B10', styleName: 'Input')
```

**Dashboard KPIs with status colors:**
```javascript
// Good/Bad/Neutral based on performance
excel_range(action: 'set-style', rangeAddress: 'B3', styleName: 'Good')   // Positive KPI
excel_range(action: 'set-style', rangeAddress: 'B4', styleName: 'Bad')    // Negative KPI
excel_range(action: 'set-style', rangeAddress: 'B5', styleName: 'Neutral') // Warning KPI
```

**Remember:** Built-in styles first, manual formatting only when necessary!
