# Excel Data Model Tool - Troubleshooting Guide

## Quick Reference: Available Actions

**Discovery (safe, read-only):**
- `list-tables` - See all tables in Data Model
- `list-measures` - See all measures (call before any delete!)
- `list-columns` - See columns in a table
- `list-relationships` - See all relationships
- `read-table` - Get table details
- `read` - Get measure formula

**Create/Update (non-destructive):**
- `create-measure` - Add new DAX measure
- `update-measure` - Modify existing measure formula/format
- `create-relationship` - Link tables
- `update-relationship` - Change relationship active state

**Delete (DESTRUCTIVE - use with caution):**
- `delete-measure` - Remove single measure (table preserved)
- `delete-relationship` - Remove relationship (tables preserved)
- `delete-table` - **DANGER: Removes table AND ALL its measures**

---

## CRITICAL: Never Delete Tables to "Start Fresh"

Deleting a Data Model table **cascades to delete ALL measures** associated with that table. This is Excel behavior that cannot be undone.

**Before considering delete-table, ALWAYS:**
1. Call `list-measures` to see what will be lost
2. Consider if you can fix the issue without deletion
3. Ask the user for explicit confirmation

---

## Common Error Recovery Patterns

### 1. "Measure already exists"

**Error**: Trying to create a measure that already exists.

**Fix**: Use `update-measure` instead of `create-measure`:
```
action: update-measure
measureName: "ExistingMeasure"
daxFormula: "=SUM(Sales[Amount])"
```

### 2. "Invalid DAX formula" / Formula syntax error

**Common causes:**
- Missing quotes around text: `"Category"` not `Category`
- Wrong column reference: `Table[Column]` format required
- Missing table prefix: `SUM(Sales[Amount])` not `SUM(Amount)`
- Locale issues: Use US format with commas, not semicolons

**Fix**: Check formula syntax, verify column names with `list-columns`:
```
action: list-columns
tableName: "Sales"
```

### 3. "Table not found"

**Fix**: Verify table name with `list-tables`:
```
action: list-tables
```

### 4. "Measure not found"

**Fix**: Verify measure name with `list-measures`:
```
action: list-measures
```

### 5. "Column not found" (for relationships)

**Fix**: Check both tables have the expected columns:
```
action: list-columns
tableName: "SourceTable"
```

---

## Recommended Workflow Pattern

**Always: Discovery → Action → Verify**

1. **Before creating**: `list-measures` to avoid duplicates
2. **Before updating**: `read` to see current formula
3. **Before deleting**: `list-measures` to understand impact
4. **After any change**: Verify with appropriate list action

---

## When Truly Stuck

If operations keep failing:

1. **Check session is valid** - reopen file if needed
2. **Verify object names exactly** - case-sensitive, check spelling
3. **Check for hidden objects** - user may need to unhide in Power Pivot UI
4. **DAX locale** - formulas should use US format (commas, not semicolons)

**Do NOT delete and recreate tables** - you will lose all measures.
