# BEFORE FORMATTING RANGES - GATHER THIS INFO

REQUIRED:
☐ Excel file path
☐ Worksheet name
☐ Range address (e.g., 'A1:D10', 'SalesData')

FORMATTING OPTIONS (choose what applies):

FONT:
☐ Font name (Arial, Calibri, Times New Roman)
☐ Font size (8, 10, 11, 12, 14, 16)
☐ Bold, italic, underline
☐ Font color (#RRGGBB hex code, e.g., #FF0000 for red)

FILL:
☐ Fill color (#RRGGBB hex code, e.g., #FFFF00 for yellow)

BORDERS:
☐ Border style (none, continuous, dash, dot, double)
☐ Border weight (hairline, thin, medium, thick)
☐ Border color (#RRGGBB hex code)

ALIGNMENT:
☐ Horizontal (left, center, right, justify)
☐ Vertical (top, center, bottom)
☐ Wrap text (true/false)

NUMBER FORMAT:
☐ Format code (e.g., '$#,##0.00', '0.00%', 'mm/dd/yyyy')

COMMON PATTERNS:
- Headers: bold, fillColor: '#4472C4', fontColor: '#FFFFFF', center aligned
- Currency: formatCode: '$#,##0.00', right aligned
- Dates: formatCode: 'mm/dd/yyyy'
- Percentages: formatCode: '0.00%'

ASK USER what formatting they want before calling excel_range(action: 'format-range')
