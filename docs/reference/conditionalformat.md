# Conditional formatting

Read existing rules before adding or replacing them. Rules are returned in
priority order with their applies-to ranges and type-specific settings.
Clear rules only when replacing the existing formatting is intended.
Current commands and inputs come from CLI help or MCP tool descriptions.

## Examples

For highlighting rows whose first column says Active, the worksheet expression
`=$A2="Active"` fixes the column but lets the row vary. Conditional formulas use
the top-left target cell's perspective. Use `$A$2` instead when every row should
refer to one fixed cell.
PowerShell single quotes preserve dollar signs in formulas.

Use the tool schema or native help for rule types and thresholds.
Supply only the properties for the selected rule type, not a blanket payload
containing settings for every kind of rule.

Read-back details include color-scale stops, data-bar limits and direction,
icon criteria, top/bottom rank, average mode, or date period only for the matching
rule type. Numeric formulas may be normalized (100 becomes `=100`); compare their
meaning, not just their original spelling.

## Change only the selected rule

Use targeted edits instead of clearing all rules. Select the rule's current
worksheet-wide priority and fingerprint from a fresh listing. Priority is not its index in a
range's collection, and native priorities can have gaps for disjoint rules.
Fingerprints describe the listed settings; they are not permanent rule IDs.
Changed or removed selections fail before edits. List again after changes or
reordering and use fresh selection values.

```mcp
conditionalformat(action: 'update-rule', session_id: sessionId, sheet_name: 'Data', rule_priority: selectedPriority, expected_fingerprint: selectedFingerprint, options: {formula1: '150', stopIfTrue: false, appliesTo: 'B2:B100'})
conditionalformat(action: 'set-rule-priority', session_id: sessionId, sheet_name: 'Data', rule_priority: freshPriority, expected_fingerprint: freshFingerprint, new_priority: 1)
```

```cli
excelcli -q conditionalformat update-rule --session $sessionId --sheet Data --rule-priority $selectedPriority --expected-fingerprint $selectedFingerprint --options '{"formula1":"150","stopIfTrue":false,"appliesTo":"B2:B100"}'
excelcli -q conditionalformat set-rule-priority --session $sessionId --sheet Data --rule-priority $freshPriority --expected-fingerprint $freshFingerprint --new-priority 1
```

The example assumes the selection came from a current listing and is refreshed
after the first edit. Changing priority changes how rules interact; read the
final rules and displayed cells.

A targeted update retains the rule type. A two-color scale cannot gain a
midpoint in place, and changing an icon set may reset native thresholds.
Stop-if-true behavior does not apply to scales, bars, or icons.

Native failures do not promise rollback. Inspect current rules after a failed
write instead of retrying a stale selection or recreating unrelated rules.
