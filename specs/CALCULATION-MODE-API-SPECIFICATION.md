# Calculation Mode API Specification

> **User-controllable calculation mode for performance optimization and workflow control**
> 
> **🤖 Primary Audience:** LLMs using MCP Server/CLI to automate Excel workbooks with complex formulas, Data Models, or bulk operations

## What This Spec Provides (For LLMs)

This specification defines a **Calculation Mode API** that lets you (an LLM) explicitly control when Excel recalculates formulas. This unlocks:

### **Performance Optimization** - Batch Without Overhead
- Current state: Every write operation toggles calculation mode internally (50 writes = 50 toggles)
- With this API: Set manual mode once → 50 writes → Calculate once → ~10-50% faster

### **Predictable Timing** - Know Exactly When Recalc Happens
- No surprise delays mid-operation
- No COM timeouts from unexpected DAX/Data Model recalculations
- Better token efficiency (fewer retries, fewer status checks)

### **Preview Before Commit** - Validate Formulas Without Side Effects
- Write formulas → Read them back → Verify correct → THEN calculate
- Useful for complex financial models, debugging formula chains

### **Step-Through Debugging** - Systematic Formula Analysis
- Change one input → Recalculate → Check dependent cells → Repeat
- Methodical debugging impossible with automatic mode

---

## Background: Current State

### Internal Optimization (Invisible to LLMs)

The server already uses calculation mode internally to prevent COM timeouts:

```csharp
// Current pattern in RangeCommands.Values.cs, RangeCommands.Formulas.cs, etc.
public OperationResult SetValues(IExcelBatch batch, string sheetName, string rangeAddress, List<List<object?>> values)
{
    return batch.Execute((ctx, ct) =>
    {
        int originalCalculation = ctx.App.Calculation;  // Save
        ctx.App.Calculation = -4135;  // xlCalculationManual
        
        try
        {
            range.Value2 = arrayValues;  // Write without triggering recalc
        }
        finally
        {
            ctx.App.Calculation = originalCalculation;  // Restore (triggers recalc)
        }
    });
}
```

**Problem:** This per-operation toggling:
1. Adds overhead on every write (even if user wants many writes before recalc)
2. Triggers recalculation after EVERY operation (expensive for Data Model workbooks)
3. Invisible to LLM - can't be optimized at workflow level

### CHANGELOG Reference (Issue #412)

```markdown
- **COM Timeout with Data Model Dependencies** (#412): Fixed timeout when setting 
  formulas/values that trigger Data Model recalculation
  - ROOT CAUSE: Excel's automatic calculation blocks COM interface during DAX recalculation
  - FIX: Temporarily disable calculation mode (xlCalculationManual) during write operations
```

---

## Research: Excel Calculation Modes

### Application.Calculation Property

**Excel COM API:**
- `Application.Calculation` - Get/set calculation mode
- Type: `XlCalculation` enum (integer values)

**Calculation Modes:**

| Mode | Value | Behavior |
|------|-------|----------|
| `xlCalculationAutomatic` | -4105 | Recalculates when any value changes (default) |
| `xlCalculationManual` | -4135 | Only recalculates when explicitly requested |
| `xlCalculationSemiautomatic` | 2 | Auto except data tables (recalc-intensive) |

**Official Reference:**
- [Application.Calculation Property](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculation)
- [XlCalculation Enumeration](https://learn.microsoft.com/en-us/office/vba/api/excel.xlcalculation)

### Triggering Calculation

**Scope Options:**

| Method | Scope | Use Case |
|--------|-------|----------|
| `Application.Calculate()` | All open workbooks | After batch operations across files |
| `Workbook.Calculate()` | Single workbook (undocumented but works) | Not recommended - use Application |
| `Worksheet.Calculate()` | Single sheet | Targeted recalc after sheet changes |
| `Range.Calculate()` | Specific range | Surgical recalc for formula debugging |

**Official Reference:**
- [Application.Calculate Method](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculate)
- [Worksheet.Calculate Method](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.calculate)
- [Range.Calculate Method](https://learn.microsoft.com/en-us/office/vba/api/excel.range.calculate)

### Calculation State

**Properties to track:**

| Property | Type | Purpose |
|----------|------|---------|
| `Application.CalculationState` | `XlCalculationState` | Current calculation status |
| - `xlDone` (0) | | Calculation complete |
| - `xlPending` (-4108) | | Calculation needed (dirty cells exist) |
| - `xlCalculating` (1) | | Currently calculating |

**Official Reference:**
- [Application.CalculationState Property](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculationstate)

---

## Proposed API Design

### New Tool: `excel_calculation_mode`

A dedicated tool for calculation mode control, separate from existing tools.

### Actions

#### 1. `get-mode` - Query Current Calculation Mode

**Purpose:** Check current mode and whether calculation is pending.

**Parameters:** None (applies to Excel Application)

**Returns:**
```json
{
  "success": true,
  "mode": "automatic",
  "modeValue": -4105,
  "calculationState": "done",
  "calculationStateValue": 0,
  "isPending": false
}
```

**Use Case:** Before starting batch operations, check if already in manual mode (avoid redundant toggle).

---

#### 2. `set-mode` - Set Calculation Mode

**Purpose:** Switch between automatic, manual, or semi-automatic calculation.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `mode` | string | Yes | `automatic`, `manual`, or `semi-automatic` |

**Mode Details:**

| Mode | When to Use |
|------|-------------|
| `automatic` | Normal operation (default) - formulas recalculate on any change |
| `manual` | Batch operations, performance-critical workflows, formula debugging |
| `semi-automatic` | Workbooks with large data tables (Excel `xlCalculationSemiautomatic`) |

**Returns:**
```json
{
  "success": true,
  "previousMode": "automatic",
  "newMode": "manual",
  "message": "Calculation mode set to manual. Call calculate action when ready to recalculate.",
  "suggestedNextActions": [
    "Perform your write operations (values, formulas, data)",
    "excel_calculation_mode(action: 'calculate') when ready to recalculate"
  ]
}
```

**Important:** When session closes, mode should be restored to original state (safety net).

---

#### 3. `calculate` - Trigger Calculation

**Purpose:** Explicitly recalculate when in manual mode.

**Parameters:**

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `scope` | string | No | `workbook` | `workbook`, `sheet`, or `range` |
| `sheetName` | string | No* | - | Required when scope = `sheet` or `range` |
| `rangeAddress` | string | No* | - | Required when scope = `range` |

**Scope Details:**

| Scope | Excel COM Method | Use Case |
|-------|------------------|----------|
| `workbook` | `Application.Calculate()` | After batch operations, general recalc |
| `sheet` | `Worksheet.Calculate()` | Targeted recalc, avoid touching other sheets |
| `range` | `Range.Calculate()` | Formula debugging, step-through analysis |

**Returns:**
```json
{
  "success": true,
  "scope": "sheet",
  "sheetName": "Calculations",
  "calculationState": "done",
  "message": "Calculation complete for sheet 'Calculations'"
}
```

---

### Core Interface (ICalculationCommands)

```csharp
namespace Sbroenne.ExcelMcp.Core.Commands.Calculation;

public interface ICalculationCommands
{
    /// <summary>
    /// Gets the current calculation mode and state
    /// </summary>
    CalculationModeResult GetMode(IExcelBatch batch);
    
    /// <summary>
    /// Sets the calculation mode (automatic, manual, semi-automatic)
    /// </summary>
    OperationResult SetMode(IExcelBatch batch, CalculationMode mode);
    
    /// <summary>
    /// Triggers calculation for the specified scope
    /// </summary>
    OperationResult Calculate(IExcelBatch batch, CalculationScope scope, string? sheetName = null, string? rangeAddress = null);
}

public enum CalculationMode
{
    Automatic = -4105,      // xlCalculationAutomatic
    Manual = -4135,         // xlCalculationManual  
    SemiAutomatic = 2       // xlCalculationSemiautomatic
}

public enum CalculationScope
{
    Workbook,   // Application.Calculate()
    Sheet,      // Worksheet.Calculate()
    Range       // Range.Calculate()
}

public class CalculationModeResult : OperationResult
{
    public string Mode { get; set; } = string.Empty;
    public int ModeValue { get; set; }
    public string CalculationState { get; set; } = string.Empty;
    public int CalculationStateValue { get; set; }
    public bool IsPending { get; set; }
}
```

---

## CLI Commands

### Command Structure

```powershell
excelcli calculation <action> [options]
```

### Actions

```powershell
# Get current mode
excelcli calculation get-mode --session 1
excelcli calculation get-mode --file "C:\Reports\Sales.xlsx"

# Set mode
excelcli calculation set-mode --session 1 --mode manual
excelcli calculation set-mode --session 1 --mode automatic
excelcli calculation set-mode --session 1 --mode semi-automatic

# Trigger calculation
excelcli calculation calculate --session 1
excelcli calculation calculate --session 1 --scope sheet --sheet "Calculations"
excelcli calculation calculate --session 1 --scope range --sheet "Data" --range "E2:E100"
```

---

## LLM Workflow Examples

### Example 1: High-Throughput Batch (Performance Optimization)

**Scenario:** LLM needs to populate 500 cells with data and formulas.

**Before (Current - Internal Toggling):**
```
# 500 operations = 500 internal mode toggles + 500 recalculations
excelcli range set-values A1 ...    # Toggle → Write → Toggle → Recalc
excelcli range set-values A2 ...    # Toggle → Write → Toggle → Recalc
... × 500
# Total: ~500 recalculations during operation
```

**After (With Calculation API):**
```powershell
# Step 1: Disable recalculation
excelcli calculation set-mode --session 1 --mode manual

# Step 2: Batch operations (NO recalculations, internal toggle skipped)
excelcli range set-values --session 1 --sheet Data --range A1 --values '[["Value1"]]'
excelcli range set-values --session 1 --sheet Data --range A2 --values '[["Value2"]]'
... × 500

# Step 3: Single recalculation at the end
excelcli calculation calculate --session 1
excelcli calculation set-mode --session 1 --mode automatic
# Total: 1 recalculation
```

**Expected Improvement:** 10-50% faster depending on formula complexity.

---

### Example 2: Formula Debugging (Step-Through)

**Scenario:** User reports formula `=INDEX(MATCH(...))` returns wrong value. LLM debugs.

```powershell
# Step 1: Enter debug mode
excelcli calculation set-mode --session 1 --mode manual

# Step 2: Check current formula
excelcli range get-formulas --session 1 --sheet Lookup --range E5
# Returns: =INDEX(Products!B:B,MATCH(D5,Products!A:A,0))

# Step 3: Check lookup value
excelcli range get-values --session 1 --sheet Lookup --range D5
# Returns: "Widget-A"

# Step 4: Check if lookup value exists in source
excelcli range get-values --session 1 --sheet Products --range A1:A100
# LLM scans: "Widget-A" is at row 15

# Step 5: Manually set D5 to a known good value
excelcli range set-values --session 1 --sheet Lookup --range D5 --values '[["Widget-B"]]'

# Step 6: Recalculate JUST that cell
excelcli calculation calculate --session 1 --scope range --sheet Lookup --range E5

# Step 7: Check result
excelcli range get-values --session 1 --sheet Lookup --range E5
# LLM: "Now it returns the correct value. The issue was the original D5 had trailing whitespace."

# Step 8: Restore automatic mode
excelcli calculation set-mode --session 1 --mode automatic
```

---

### Example 3: Data Model Workbook (Timeout Prevention)

**Scenario:** Workbook has Power Pivot Data Model. Writing values triggers expensive DAX recalculation.

```powershell
# Problem: Without manual mode, this times out
excelcli range set-values --session 1 --sheet Input --range A2 --values '[[1000000]]'
# COM timeout: DAX measures recalculating across 5M rows

# Solution: Batch inputs, then recalc
excelcli calculation set-mode --session 1 --mode manual

excelcli range set-values --session 1 --sheet Input --range A2 --values '[[1000000]]'
excelcli range set-values --session 1 --sheet Input --range B2 --values '[["East"]]'
excelcli range set-values --session 1 --sheet Input --range C2 --values '[["2025-Q1"]]'

# Now recalculate (user expects this to take time)
excelcli calculation calculate --session 1
# Output: { "success": true, "message": "Calculation complete", "duration": "12.3s" }

excelcli calculation set-mode --session 1 --mode automatic
```

---

### Example 4: Preview Formulas Before Committing

**Scenario:** LLM builds complex formula, wants to verify syntax before recalculation.

```powershell
excelcli calculation set-mode --session 1 --mode manual

# Write formula (no immediate recalc)
excelcli range set-formulas --session 1 --sheet Analysis --range F2 \
  --formulas '[["=SUMPRODUCT((Region=\"East\")*(Year=2025)*Sales)"]]'

# Read it back to verify
excelcli range get-formulas --session 1 --sheet Analysis --range F2
# Returns: "=SUMPRODUCT((Region=\"East\")*(Year=2025)*Sales)"
# LLM: "Formula syntax looks correct."

# Now calculate
excelcli calculation calculate --session 1 --scope range --sheet Analysis --range F2

# Check result
excelcli range get-values --session 1 --sheet Analysis --range F2
# Returns: 1250000

excelcli calculation set-mode --session 1 --mode automatic
```

---

## Implementation Notes

### Session Cleanup

When a session closes (via `session close` or daemon shutdown), the calculation mode MUST be restored to its original value:

```csharp
// In ExcelSession.Close() or Dispose()
if (_originalCalculationMode != null)
{
    try
    {
        _app.Calculation = _originalCalculationMode.Value;
    }
    catch
    {
        // Best effort - workbook closing anyway
    }
}
```

### Integration with Existing Write Operations

When calculation mode is already manual (set by user via this API), internal toggles should be skipped:

```csharp
// In RangeCommands.SetValues, etc.
int currentMode = ctx.App.Calculation;
bool needsToggle = currentMode != -4135; // Only toggle if not already manual

if (needsToggle)
{
    ctx.App.Calculation = -4135;
}

try
{
    // ... write operation ...
}
finally
{
    if (needsToggle)
    {
        ctx.App.Calculation = currentMode;
    }
}
```

### Response Enrichment

When calculation mode is manual, include `calculationPending: true` in write operation responses:

```json
{
  "success": true,
  "action": "set-values",
  "message": "Values written to Sheet1!A1:D10",
  "calculationPending": true,
  "hint": "Call 'calculate' action to recalculate formulas"
}
```

---

## MCP Tool Schema

### Tool Definition

```csharp
[McpServerToolType]
public static class ExcelCalculationTool
{
    /// <summary>
    /// Control when Excel recalculates formulas (automatic vs manual mode).
    /// 
    /// ACTIONS:
    /// - get-mode: Query current calculation mode and state
    /// - set-mode: Switch between automatic, manual, or semi-automatic
    /// - calculate: Trigger recalculation (required when in manual mode)
    /// 
    /// CALCULATION MODES:
    /// - automatic: Recalculates on any change (default, standard behavior)
    /// - manual: Only recalculates when you call 'calculate' action
    /// - semi-automatic: Auto except data tables (for recalc-intensive workbooks)
    /// 
    /// WHEN TO USE MANUAL MODE:
    /// - Batch operations: Set manual → many writes → calculate once
    /// - Data Model workbooks: Prevent DAX recalc timeouts
    /// - Formula debugging: Change input → calculate → check output → repeat
    /// 
    /// IMPORTANT: Mode is restored to original when session closes.
    /// </summary>
    [McpServerTool(Name = "excel_calculation_mode")]
    public static async Task<string> ExecuteAsync(
        [Description("Action: get-mode, set-mode, calculate")]
        CalculationModeAction action,
        
        [Description("Session ID (required)")]
        int session,
        
        [Description("Calculation mode for set-mode: automatic, manual, semi-automatic")]
        string? mode = null,
        
        [Description("Calculation scope for calculate: workbook (default), sheet, range")]
        string? scope = null,
        
        [Description("Sheet name (required when scope=sheet or scope=range)")]
        string? sheetName = null,
        
        [Description("Range address (required when scope=range)")]
        string? rangeAddress = null
    )
    {
        // Implementation...
    }
}
```

---

## Testing Strategy

### Test Categories

1. **Mode Toggle Tests**
   - Get mode returns correct values
   - Set mode changes Application.Calculation
   - Set mode preserves original for session cleanup

2. **Calculate Scope Tests**
   - Workbook scope calculates all sheets
   - Sheet scope calculates only target sheet
   - Range scope calculates only target range

3. **Integration Tests**
   - Manual mode + batch writes + single calculate
   - Manual mode persists across multiple operations
   - Session close restores original mode

4. **Performance Tests** (OnDemand)
   - Measure time: 100 writes with internal toggling vs manual mode
   - Document actual improvement percentage

5. **Data Model Tests**
   - Verify no timeout when writing to cells with DAX dependencies in manual mode
   - Verify DAX recalculates correctly after explicit calculate call

### Test Traits

```csharp
[Trait("Feature", "Calculation")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
```

---

## Migration Path

### Phase 1: Core Implementation
- Add `ICalculationCommands` interface
- Implement `CalculationCommands` class
- Add unit tests

### Phase 2: MCP Server
- Add `excel_calculation_mode` tool
- Add `CalculationModeAction` enum
- Update tool count in docs

### Phase 3: CLI
- Add `calculation` command group
- Implement `get-mode`, `set-mode`, `calculate` subcommands
- Update CLI documentation

### Phase 4: Optimization
- Modify existing write operations to skip internal toggle when already manual
- Add `calculationPending` to write operation responses
- Update skill files with calculation mode guidance

---

## Acceptance Criteria

### Functional
- [ ] `get-mode` returns current mode and calculation state
- [ ] `set-mode` changes calculation mode
- [ ] `calculate` triggers recalculation at specified scope
- [ ] Session close restores original calculation mode
- [ ] Write operations skip internal toggle when already in manual mode

### Performance
- [ ] Batch of 100 writes is measurably faster with manual mode
- [ ] No COM timeouts on Data Model workbooks in manual mode

### Documentation
- [ ] MCP tool has comprehensive XML documentation
- [ ] CLI has help text for all commands
- [ ] Skill files updated with calculation mode workflows
- [ ] FEATURES.md updated with new tool/actions

### Testing
- [ ] All calculate scopes tested (workbook, sheet, range)
- [ ] Mode persistence across operations verified
- [ ] Session cleanup verified
- [ ] Integration with existing write operations verified

---

## Open Questions

1. **Should we track original mode per-session or globally?**
   - Recommendation: Per-session. Multiple sessions might want different modes.

2. **Should calculate action report duration?**
   - Recommendation: Yes. Useful for LLMs to understand workbook complexity.

3. **Should we add a `force-calculate` that works even in automatic mode?**
   - Recommendation: No. `Application.Calculate()` already works in automatic mode.

4. **Should response include dirty cell count when in manual mode?**
   - Recommendation: Nice to have. Low priority. `calculationPending: true` is sufficient.

---

## Summary

The Calculation Mode API transforms calculation control from an invisible internal optimization to an explicit, user-controllable feature. This enables:

| Capability | Current | With This API |
|------------|---------|---------------|
| Batch performance | 50 writes = 50 recalcs | 50 writes = 1 recalc |
| Data Model safety | Internal toggle (helps) | Explicit control (better) |
| Formula debugging | Not possible | Step-through analysis |
| Timing predictability | Surprises possible | Fully controlled |

**Estimated Effort:** Medium (3-5 days)
- Core: 1 day
- MCP Server: 0.5 day
- CLI: 0.5 day
- Tests: 1-2 days
- Documentation: 0.5 day

---

## When NOT to Use Manual Mode

Manual mode adds complexity. Use automatic mode (default) when:

| Scenario | Why Automatic is Better |
|----------|------------------------|
| **Single-cell reads/writes** | Overhead of mode toggle > recalc cost |
| **Simple workbooks** | No complex formulas to recalculate anyway |
| **When you need immediate feedback** | Reading calculated values right after write |
| **Unfamiliar workbooks** | Don't know if formula dependencies exist |
| **Interactive user sessions** | User expects Excel to update in real-time |

**Rule of Thumb:** If you're doing < 10 write operations, don't bother with manual mode.

---

## Integration with Other Tools

### Power Query (`excel_powerquery`)

Manual mode is **highly recommended** when:
- Importing multiple queries in sequence
- Refreshing queries that feed into Data Model
- Building query chains (one query references another)

```powershell
excelcli calculation set-mode --session 1 --mode manual
excelcli powerquery import --session 1 --name "Sales" --formula "..."
excelcli powerquery import --session 1 --name "Products" --formula "..."
excelcli powerquery import --session 1 --name "Combined" --formula "..." # References Sales, Products
excelcli calculation calculate --session 1
excelcli calculation set-mode --session 1 --mode automatic
```

### Data Model (`excel_datamodel`)

**Critical for timeout prevention:**
- Writing to cells that feed DAX measures
- Adding multiple measures in sequence
- Refreshing Data Model connections

```powershell
excelcli calculation set-mode --session 1 --mode manual
excelcli datamodel add-measure --session 1 --table Sales --name "Total Revenue" --formula "SUM(Sales[Amount])"
excelcli datamodel add-measure --session 1 --table Sales --name "YoY Growth" --formula "..."
excelcli range set-values --session 1 --sheet Input --range A2 --values '[[1000000]]'  # Feeds DAX
excelcli calculation calculate --session 1
excelcli calculation set-mode --session 1 --mode automatic
```

### PivotTables (`excel_pivottable`)

Consider manual mode when:
- Creating multiple PivotTables from same source
- Configuring PivotTable then adding slicers
- Batch field configuration

---

## Documentation Update Checklist

When this feature is implemented, **ALL** these files require updates:

### Tool/Operation Counts (22 → 23 tools, 211 → 214 operations)

| File | Location | Change |
|------|----------|--------|
| `README.md` | Line ~21, ~84, ~100 | Update tool/operation counts |
| `src/ExcelMcp.McpServer/README.md` | Lines ~18, ~56, ~72 | Update counts |
| `src/ExcelMcp.CLI/README.md` | Lines ~9, ~102 | Update counts (14 → 15 categories) |
| `vscode-extension/README.md` | Line ~19 | Update counts |
| `gh-pages/index.md` | Feature reference link | Update counts |
| `gh-pages/features.md` | Add new section | Add `excel_calculation_mode` tool |
| `gh-pages/404.md` | Line ~20 | Update counts |
| `mcpb/manifest.json` | `long_description` | Update counts |
| `src/ExcelMcp.McpServer/Program.cs` | Line ~280 | Update server description |
| `FEATURES.md` | Add new section | Full tool documentation |
| `tests/.../McpServerSmokeTests.cs` | Line ~179 | Update expected tool count (22 → 23) |

### Skills (LLM Guidance)

| File | Change |
|------|--------|
| `skills/shared/workflows.md` | Add "Batch Operations with Calculation Mode" workflow |
| `skills/shared/behavioral-rules.md` | Add rule: "Use manual mode for batch operations" |
| `skills/excel-mcp/SKILL.md` | Add `excel_calculation_mode` to tool reference |
| `skills/excel-cli/SKILL.md` | Add `calculation` command reference |
| **NEW:** `skills/shared/excel_calculation.md` | Dedicated calculation mode guidance |

### MCP Prompts

| File | Change |
|------|--------|
| `src/ExcelMcp.McpServer/Prompts/Content/excel_datamodel.md` | Add: "Use manual mode when writing to cells with DAX dependencies" |
| `src/ExcelMcp.McpServer/Prompts/Content/excel_powerquery.md` | Add: "Consider manual mode during batch query operations" |

### Code Changes

| File | Change |
|------|--------|
| `src/ExcelMcp.Core/ToolActions.cs` | Add `CalculationModeAction` enum |
| `src/ExcelMcp.Core/ActionExtensions.cs` | Add `ToActionString()` for new enum |
| Write operation commands | Skip internal toggle when already manual |
| Write operation results | Add `calculationPending: true` when in manual mode |
