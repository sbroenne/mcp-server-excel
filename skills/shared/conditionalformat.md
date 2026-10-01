# Conditional formatting

Read existing rules before adding or replacing them. Rules are returned in
priority order with their applies-to ranges and type-specific settings.
Clear rules only when replacing the existing formatting is intended.

## Examples

For an existing `Data` sheet and captured session, highlight amounts greater
than 100, or separately highlight rows whose first column says Active:

```mcp
conditionalformat(action: 'add-rule', session_id: sessionId, sheet_name: 'Data', range_address: 'B2:B100', rule_type: 'cell-value', operator_type: 'greater', formula1: '100', interior_color: '#FFFF00')
conditionalformat(action: 'add-rule', session_id: sessionId, sheet_name: 'Data', range_address: 'A2:E100', rule_type: 'expression', formula1: '=$A2="Active"', interior_color: '#90EE90')
conditionalformat(action: 'list-worksheet-rules', session_id: sessionId, sheet_name: 'Data')
```

```cli
excelcli -q conditionalformat add-rule --session $sessionId --sheet Data --range B2:B100 --rule-type cell-value --operator-type greater --formula1 '100' --interior-color '#FFFF00'
excelcli -q conditionalformat add-rule --session $sessionId --sheet Data --range A2:E100 --rule-type expression --formula1 '=$A2="Active"' --interior-color '#90EE90'
excelcli -q conditionalformat list-worksheet-rules --session $sessionId --sheet Data
```

Check each result. Expression formulas use the top-left target cell's perspective.
Use `$A2` for a fixed column and relative row, or `$A$2` for one fixed cell.
PowerShell single quotes preserve dollar signs in formulas.

Use the tool schema or native help for rule types and thresholds.
Supply only the properties for the selected rule type, not a blanket payload
containing settings for every kind of rule.

Read-back details include color-scale stops, data-bar limits and direction,
icon criteria, top/bottom rank, average mode, or date period only for the matching
rule type. Numeric formulas may be normalized (100 becomes `=100`); compare their
meaning, not just their original spelling.
