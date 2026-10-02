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
excelcli -q conditionalformat add-rule --session $sessionId --sheet-name Data --range-address B2:B100 --rule-type cell-value --operator-type greater --formula1 '100' --interior-color '#FFFF00'
excelcli -q conditionalformat add-rule --session $sessionId --sheet-name Data --range-address A2:E100 --rule-type expression --formula1 '=$A2="Active"' --interior-color '#90EE90'
excelcli -q conditionalformat list-worksheet-rules --session $sessionId --sheet-name Data
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

## Change only the selected rule

Use `update-rule`, `delete-rule`, or `set-rule-priority` instead of clearing all
rules. Select the rule's current worksheet-wide `priority` and `fingerprint`
from `list-rules` or `list-worksheet-rules`. Priority is not its index in a
range's collection, and native priorities can have gaps for disjoint rules.
Fingerprints describe the listed settings; they are not permanent rule IDs.
Changed or removed selections fail before edits. List again after changes or
reordering and use fresh selection values.

```mcp
conditionalformat(action: 'update-rule', session_id: sessionId, sheet_name: 'Data', rule_priority: selectedPriority, expected_fingerprint: selectedFingerprint, options: {formula1: '150', stopIfTrue: false, appliesTo: 'B2:B100'})
conditionalformat(action: 'set-rule-priority', session_id: sessionId, sheet_name: 'Data', rule_priority: freshPriority, expected_fingerprint: freshFingerprint, new_priority: 1)
```

```cli
excelcli -q conditionalformat update-rule --session $sessionId --sheet-name Data --rule-priority $selectedPriority --expected-fingerprint $selectedFingerprint --options '{"formula1":"150","stopIfTrue":false,"appliesTo":"B2:B100"}'
excelcli -q conditionalformat set-rule-priority --session $sessionId --sheet-name Data --rule-priority $freshPriority --expected-fingerprint $freshFingerprint --new-priority 1
```

Only supplied nested settings change; their JSON names stay camelCase in both
entry points. The rule's type is retained. Use the matching visual settings to
change scale stops, bar limits/colors, icons/thresholds, top/bottom ranks, average
comparisons, date periods, or duplicate/unique selection. A two-color scale
cannot gain a midpoint in place. Changing an icon set may reset native thresholds.
`stopIfTrue` is not available for color scales, data bars, or icon sets.
`add-rule` can set a new rule's priority and stop flag directly.

Native failures do not promise rollback. Inspect current rules after a failed
write instead of retrying a stale selection or recreating unrelated rules.
