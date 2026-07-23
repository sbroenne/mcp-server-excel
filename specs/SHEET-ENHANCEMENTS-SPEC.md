# Sheet Enhancement Specification - Tab Color & Visibility

> **Enhanced worksheet commands for tab color and visibility management**
> 
> **🤖 Primary Audience:** LLMs using MCP Server tools to automate Excel workbook organization

## What This Spec Provides (For LLMs)

This specification defines 8 new MCP Server actions that let you (an LLM) programmatically:

### **Tab Colors** - Visual Organization
- **Set colors** using RGB values (0-255 each) - no need to know about BGR conversion
- **Read colors** to audit existing workbooks or preserve color schemes
- **Clear colors** to reset tabs to default appearance
- **Common use:** Color-code by department, project status, data category, priority level

### **Sheet Visibility** - Control What Users See
- **Show/Hide** sheets with three levels of visibility
- **Hidden** - Users can unhide via Excel UI (for archive/reference data)
- **VeryHidden** - Only code can unhide (for templates, calculations, sensitive data)
- **Common use:** Protect formulas, hide templates, secure salary data, manage multi-user workbooks

### **Why You Need These Tools**
When users ask you to "organize the sales workbook" or "set up a dashboard with calculations," you'll use these commands to:
1. Create professional-looking workbooks with color-coded tabs
2. Hide internal worksheets users shouldn't modify
3. Protect sensitive data from casual viewing
4. Implement visual workflows (red=todo, yellow=in-progress, green=complete)

---

## Technical Overview

This specification extends the existing SheetCommands to support two key worksheet appearance features:

1. **Tab Color Management** - Set and retrieve worksheet tab colors (RGB → BGR conversion handled automatically)
2. **Visibility Control** - Show/hide worksheets with two protection levels (Hidden and VeryHidden)

---

## Research: Excel Worksheet Capabilities

### Tab Color (Worksheet.Tab.Color)

**Color Format:**
- Excel uses **BGR (Blue-Green-Red) format** stored as integer
- RGB(255, 0, 0) becomes 0x0000FF (red in BGR)
- RGB(0, 255, 0) becomes 0x00FF00 (green)
- RGB(0, 0, 255) becomes 0xFF0000 (blue)
- Formula: `BGR = (Blue << 16) | (Green << 8) | Red`

**ColorIndex Alternative:**
- Excel also supports `Tab.ColorIndex` property using `XlColorIndex` enum (legacy, limited palette)
- Modern approach: Use `Tab.Color` with RGB values

**Official Reference:**
- [Tab.Color Property](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.tab.color)
- [Tab.ColorIndex Property](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.tab.colorindex)

### Worksheet Visibility (Worksheet.Visible)

**Excel COM API:**
- Access via `worksheet.Visible` property
- Set to `XlSheetVisibility` enum values
- Get returns current visibility state

**Visibility Levels:**
1. **xlSheetVisible (-1)** - Normal visible state
2. **xlSheetHidden (0)** - Hidden via UI (right-click → Hide), user can unhide via UI
3. **xlSheetVeryHidden (2)** - Programmatically hidden, requires code to unhide (security/protection)

**Official Reference:**
- [Worksheet.Visible Property](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel._worksheet.visible)
- [XlSheetVisibility Enumeration](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlsheetvisibility)

---

## Proposed API Design

### Core Commands (ISheetCommands)

```csharp
public interface ISheetCommands
{
    // === EXISTING LIFECYCLE COMMANDS (No changes) ===
    Task<WorksheetListResult> ListAsync(IExcelBatch batch);
    Task<OperationResult> CreateAsync(IExcelBatch batch, string sheetName);
    Task<OperationResult> RenameAsync(IExcelBatch batch, string oldName, string newName);
    Task<OperationResult> CopyAsync(IExcelBatch batch, string sourceSheet, string newSheet);
    Task<OperationResult> DeleteAsync(IExcelBatch batch, string sheetName);
    
    // === NEW: TAB COLOR OPERATIONS ===
    
    /// <summary>
    /// Sets the tab color for a worksheet using RGB values
    /// Excel uses BGR format internally (Blue-Green-Red)
    /// </summary>
    void SetTabColor(IExcelBatch batch, string sheetName, int red, int green, int blue);
    
    /// <summary>
    /// Gets the tab color for a worksheet
    /// Returns RGB values or null if no color is set
    /// </summary>
    (int R, int G, int B)? GetTabColor(IExcelBatch batch, string sheetName);
    
    /// <summary>
    /// Clears the tab color for a worksheet (resets to default)
    /// </summary>
    void ClearTabColor(IExcelBatch batch, string sheetName);
    
    // === NEW: VISIBILITY OPERATIONS ===
    
    /// <summary>
    /// Sets worksheet visibility level
    /// </summary>
    void SetVisibility(IExcelBatch batch, string sheetName, SheetVisibility visibility);
    
    /// <summary>
    /// Gets worksheet visibility level
    /// </summary>
    SheetVisibility GetVisibility(IExcelBatch batch, string sheetName);
    
    /// <summary>
    /// Shows a hidden or very hidden worksheet
    /// Convenience method equivalent to SetVisibility(..., SheetVisibility.Visible)
    /// </summary>
    void Show(IExcelBatch batch, string sheetName);
    
    /// <summary>
    /// Hides a worksheet (user can unhide via UI)
    /// Convenience method equivalent to SetVisibility(..., SheetVisibility.Hidden)
    /// </summary>
    void Hide(IExcelBatch batch, string sheetName);
    
    /// <summary>
    /// Very hides a worksheet (requires code to unhide)
    /// Convenience method equivalent to SetVisibility(..., SheetVisibility.VeryHidden)
    /// </summary>
    void VeryHide(IExcelBatch batch, string sheetName);
}

// === SUPPORTING TYPES ===

public enum SheetVisibility
{
    Visible = -1,      // xlSheetVisible
    Hidden = 0,        // xlSheetHidden (can unhide via UI)
    VeryHidden = 2     // xlSheetVeryHidden (requires code to unhide)
}

public class TabColorResult : OperationResult
{
    public bool HasColor { get; set; }  // False if no color is set
    public int? Red { get; set; }       // 0-255, null if no color
    public int? Green { get; set; }     // 0-255, null if no color
    public int? Blue { get; set; }      // 0-255, null if no color
    public string? HexColor { get; set; }  // #RRGGBB format for convenience
}

public class SheetVisibilityResult : OperationResult
{
    public SheetVisibility Visibility { get; set; }
    public string VisibilityName { get; set; } = string.Empty;  // "Visible", "Hidden", "VeryHidden"
}
```

### Implementation Strategy

**Tab Color Operations:**
- Convert RGB (0-255 each) to BGR format: `(blue << 16) | (green << 8) | red`
- Set via `worksheet.Tab.Color` property
- Get returns integer, convert back to RGB components
- Clear by setting to 0 or `XlColorIndex.xlColorIndexNone`
- Validate RGB values are 0-255

**Visibility Operations:**
- Map `SheetVisibility` enum to Excel's `XlSheetVisibility` constants
- Set via `worksheet.Visible` property
- Get returns integer, cast to `SheetVisibility` enum
- Convenience methods call `SetVisibility` with appropriate enum value

---

## MCP Server Integration (Primary Use Case)

### Updated worksheet Tool

```typescript
{
  "name": "worksheet",
  "description": "Worksheet lifecycle and appearance management",
  "parameters": {
    "action": "string",
    "excelPath": "string",
    "sheetName": "string",
    "newSheetName": "string",  // for rename/copy
    "red": "number",           // 0-255 for set-tab-color
    "green": "number",         // 0-255 for set-tab-color
    "blue": "number",          // 0-255 for set-tab-color
    "visibility": "string"     // "visible" | "hidden" | "veryhidden"
  },
  "actions": [
    // Existing lifecycle operations
    "list",
    "create",
    "rename",
    "copy",
    "delete",
    
    // New tab color operations
    "set-tab-color",    // Set tab color with RGB values
    "get-tab-color",    // Get current tab color
    "clear-tab-color",  // Clear tab color (reset to default)
    
    // New visibility operations
    "set-visibility",   // Set visibility level (visible/hidden/veryhidden)
    "get-visibility",   // Get current visibility level
    "show",             // Convenience: make visible
    "hide",             // Convenience: hide (user can unhide)
    "very-hide"         // Convenience: very hide (requires code)
  ]
}
```

### MCP Action Examples

**Set Tab Color:**
```json
{
  "action": "set-tab-color",
  "excelPath": "Report.xlsx",
  "sheetName": "Sales",
  "red": 255,
  "green": 0,
  "blue": 0
}
// Response: { "success": true }
```

**Get Tab Color:**
```json
{
  "action": "get-tab-color",
  "excelPath": "Report.xlsx",
  "sheetName": "Sales"
}
// Response: { "success": true, "hasColor": true, "red": 255, "green": 0, "blue": 0, "hexColor": "#FF0000" }
```

**Set Visibility:**
```json
{
  "action": "set-visibility",
  "excelPath": "Report.xlsx",
  "sheetName": "Data",
  "visibility": "hidden"
}
// Response: { "success": true }
```

**Get Visibility:**
```json
{
  "action": "get-visibility",
  "excelPath": "Report.xlsx",
  "sheetName": "Data"
}
// Response: { "success": true, "visibility": "Hidden", "visibilityName": "Hidden" }
```

---

## MCP Action Reference (Quick Lookup for LLMs)

**Use this table when deciding which action to call:**

| Action | Required Parameters | Returns | When To Use |
|--------|-------------------|---------|-------------|
| `set-tab-color` | `sheetName`, `red`, `green`, `blue` | `{success}` | User wants to color-code sheets |
| `get-tab-color` | `sheetName` | `{success, hasColor, red, green, blue, hexColor}` | Need to read existing color or audit workbook |
| `clear-tab-color` | `sheetName` | `{success}` | User wants to remove color/reset to default |
| `set-visibility` | `sheetName`, `visibility` | `{success}` | Need precise control over visibility level |
| `get-visibility` | `sheetName` | `{success, visibility, visibilityName}` | Check current visibility state |
| `show` | `sheetName` | `{success}` | Make hidden/very-hidden sheet visible |
| `hide` | `sheetName` | `{success}` | Hide sheet (user can still unhide via UI) |
| `very-hide` | `sheetName` | `{success}` | Protect sheet from users (only code can unhide) |

---

## Visibility Decision Guide (For LLMs)

**When user says... → You should use:**

| User Request | Action to Call | Visibility Level | Reason |
|--------------|---------------|------------------|---------|
| "hide the calculations" | `very-hide` | VeryHidden | User shouldn't see internal formulas |
| "hide the template" | `very-hide` | VeryHidden | Templates should be protected from editing |
| "protect lookup tables" | `very-hide` | VeryHidden | Reference data shouldn't be modified |
| "hide salary/sensitive data" | `very-hide` | VeryHidden | Security - prevent casual viewing |
| "hide archive data" | `hide` | Hidden | User may need to reference old data later |
| "hide temporary sheets" | `hide` | Hidden | May need manual cleanup/review |
| "hide for now" | `hide` | Hidden | Temporary hiding, user can unhide |
| "show all sheets" | `show` | Visible | Make previously hidden sheets visible |
| "unhide everything" | `show` (loop all) | Visible | Reveal all hidden worksheets |

**Key Distinction:**
- **`hide`** (Hidden) - Users can right-click tabs → Unhide in Excel UI
- **`very-hide`** (VeryHidden) - Only code/automation can unhide (true protection)

**Visibility Level Characteristics:**

| Level | Excel Value | User Can See? | User Can Unhide? | When To Use |
|-------|-------------|---------------|------------------|-------------|
| **Visible** | `-1` | ✅ Yes | N/A | Normal sheets |
| **Hidden** | `0` | ❌ No | ✅ Yes (via UI) | Temporary hiding, archives |
| **VeryHidden** | `2` | ❌ No | ❌ No (code only) | Templates, formulas, sensitive data |

---

## Use Cases (LLM Workflows)

### 1. Color-Coding by Category (Single Operations)
**Scenario:** LLM organizing financial workbook by department

```json
// Color code departments
{ "action": "set-tab-color", "sheetName": "Sales", "red": 0, "green": 176, "blue": 240 }      // Blue
{ "action": "set-tab-color", "sheetName": "Marketing", "red": 255, "green": 192, "blue": 0 }  // Orange
{ "action": "set-tab-color", "sheetName": "Operations", "red": 146, "green": 208, "blue": 80 } // Green
{ "action": "set-tab-color", "sheetName": "Summary", "red": 192, "green": 0, "blue": 0 }      // Red
```

### 1b. Color-Coding by Category (Batch Mode - Recommended)
**Scenario:** LLM organizing financial workbook - efficient batch approach

```json
// Step 1: Begin batch session
{ "tool": "begin_excel_batch", "excelPath": "Financial-Report.xlsx", "batchId": "color-coding" }

// Step 2: Apply all colors in one session (no file saves between operations)
{ "tool": "worksheet", "action": "set-tab-color", "batchId": "color-coding", "sheetName": "Sales", "red": 0, "green": 176, "blue": 240 }
{ "tool": "worksheet", "action": "set-tab-color", "batchId": "color-coding", "sheetName": "Marketing", "red": 255, "green": 192, "blue": 0 }
{ "tool": "worksheet", "action": "set-tab-color", "batchId": "color-coding", "sheetName": "Operations", "red": 146, "green": 208, "blue": 80 }
{ "tool": "worksheet", "action": "set-tab-color", "batchId": "color-coding", "sheetName": "HR", "red": 112, "green": 48, "blue": 160 }
{ "tool": "worksheet", "action": "set-tab-color", "batchId": "color-coding", "sheetName": "Finance", "red": 255, "green": 217, "blue": 102 }
{ "tool": "worksheet", "action": "set-tab-color", "batchId": "color-coding", "sheetName": "Summary", "red": 192, "green": 0, "blue": 0 }

// Step 3: Commit batch (saves once)
{ "tool": "commit_excel_batch", "batchId": "color-coding", "saveChanges": true }

// Result: 6 sheets colored in ~2 seconds vs ~12 seconds with individual operations
```

### 2. Template Management
**Scenario:** LLM setting up workbook with hidden calculation sheets

```json
// Hide template/calculation sheets from end users
{ "action": "very-hide", "sheetName": "Template" }
{ "action": "very-hide", "sheetName": "Calculations" }
{ "action": "very-hide", "sheetName": "LookupTables" }

// Show only user-facing sheets
{ "action": "show", "sheetName": "Dashboard" }
{ "action": "show", "sheetName": "Summary" }
```

### 3. Workflow Status Indication
**Scenario:** LLM tracking project status with colors

```json
// Use colors to indicate workflow status
{ "action": "set-tab-color", "sheetName": "ToDo", "red": 255, "green": 0, "blue": 0 }        // Red - not started
{ "action": "set-tab-color", "sheetName": "InProgress", "red": 255, "green": 165, "blue": 0 } // Orange - in progress
{ "action": "set-tab-color", "sheetName": "Complete", "red": 0, "green": 255, "blue": 0 }     // Green - done
```

### 4. Data Security
**Scenario:** LLM protecting sensitive calculations

```json
// Very hide sheets containing sensitive calculations
{ "action": "very-hide", "sheetName": "SalaryData" }
{ "action": "very-hide", "sheetName": "Formulas" }

// Can only unhide via code
{ "action": "show", "sheetName": "SalaryData" }  // When authorized user needs access
```

### 5. Complete Workbook Setup Workflow
**Scenario:** LLM creating and organizing new workbook from scratch

```json
// Step 1: Create sheets
{ "tool": "worksheet", "action": "create", "excelPath": "Q1-Report.xlsx", "sheetName": "Dashboard" }
{ "tool": "worksheet", "action": "create", "excelPath": "Q1-Report.xlsx", "sheetName": "Sales Data" }
{ "tool": "worksheet", "action": "create", "excelPath": "Q1-Report.xlsx", "sheetName": "Calculations" }
{ "tool": "worksheet", "action": "create", "excelPath": "Q1-Report.xlsx", "sheetName": "Lookup Tables" }

// Step 2: Color-code by purpose
{ "tool": "worksheet", "action": "set-tab-color", "sheetName": "Dashboard", "red": 68, "green": 114, "blue": 196 }     // Blue - user-facing
{ "tool": "worksheet", "action": "set-tab-color", "sheetName": "Sales Data", "red": 112, "green": 173, "blue": 71 }   // Green - data
{ "tool": "worksheet", "action": "set-tab-color", "sheetName": "Calculations", "red": 255, "green": 192, "blue": 0 }  // Orange - internal
{ "tool": "worksheet", "action": "set-tab-color", "sheetName": "Lookup Tables", "red": 158, "green": 158, "blue": 158 } // Gray - reference

// Step 3: Hide internal sheets
{ "tool": "worksheet", "action": "very-hide", "sheetName": "Calculations" }
{ "tool": "worksheet", "action": "very-hide", "sheetName": "Lookup Tables" }

// Step 4: Populate data (using other tools)
// ... range operations, Power Query, etc.

// Result: Organized workbook with color-coded tabs and protected internal sheets
```

---

## Error Handling

**Common Error Scenarios:**

| Error Case | API Response | LLM Should... |
|------------|--------------|---------------|
| Sheet doesn't exist | `{success: false, errorMessage: "Sheet 'XYZ' not found"}` | Verify sheet exists with `list` action first |
| RGB out of range | `{success: false, errorMessage: "RGB values must be 0-255"}` | Validate RGB values before calling |
| Last visible sheet | `{success: false, errorMessage: "Cannot hide last visible sheet"}` | Check visibility of other sheets first |
| Invalid visibility value | `{success: false, errorMessage: "Invalid visibility: 'xyz'"}` | Use only: `visible`, `hidden`, `veryhidden` |

**Best Practice for LLMs:**
```json
// Always check if operation succeeded
const result = worksheet({action: "set-tab-color", sheetName: "Sales", red: 255, green: 0, blue: 0});
if (!result.success) {
  // Handle error - maybe sheet was renamed or deleted
  console.error(result.errorMessage);
}
```

---

## LLM Decision Logic (Your Automation Rules)

**Apply these rules when processing user requests:**

### 1. Color Keyword → RGB Mapping
When user says a color name, use these RGB values:

| Color Name | RGB Values | Hex | Use For |
|------------|-----------|-----|---------|
| Red | `(255, 0, 0)` | `#FF0000` | Urgent, errors, high priority |
| Green | `(0, 255, 0)` | `#00FF00` | Complete, approved, success |
| Blue | `(0, 0, 255)` | `#0000FF` | Information, primary data |
| Orange | `(255, 165, 0)` | `#FFA500` | In progress, warnings |
| Yellow | `(255, 255, 0)` | `#FFFF00` | Pending, caution |
| Purple | `(128, 0, 128)` | `#800080` | Special, VIP, custom |
| Light Blue | `(173, 216, 230)` | `#ADD8E6` | Secondary, reference |
| Light Green | `(144, 238, 144)` | `#90EE90` | Safe, verified |
| Gray | `(128, 128, 128)` | `#808080` | Inactive, archived |
| Pink | `(255, 192, 203)` | `#FFC0CB` | Special attention |

See [Common Color Presets](#common-color-presets) for complete list.

### 2. Visibility Intent Detection
Parse user intent and map to correct action:

```
User says "hide calculations/formulas/templates" 
  → Call: very-hide (VeryHidden)
  → Reason: Internal sheets shouldn't be user-accessible

User says "hide for now/temporarily" 
  → Call: hide (Hidden)
  → Reason: User may need to access later

User says "protect/secure sensitive data" 
  → Call: very-hide (VeryHidden)
  → Reason: Security requirement

User says "show/unhide/make visible" 
  → Call: show (Visible)
  → Reason: Make accessible
```

### 3. Batch vs Single Operations
Optimize performance based on operation count:

```
Coloring/hiding 3+ sheets
  → Use: Batch mode (begin_excel_batch → operations → commit_excel_batch)
  → Benefit: ~5-6x faster (one file save vs N saves)

Single sheet operation
  → Use: Direct action call
  → Benefit: Simpler, adequate for single operation

Mixed operations (create + color + hide)
  → Use: Batch mode
  → Benefit: Transactional consistency
```

### 4. Integration with Other Commands
Chain operations for complete workflows:

```
When creating new sheet:
  1. worksheet(action: "create", sheetName: "Sales")
  2. worksheet(action: "set-tab-color", sheetName: "Sales", red: 0, green: 176, blue: 240)
  → Result: New sheet with color applied immediately

When organizing workbook:
  1. Color code by category (batch operation)
  2. Hide internal sheets (very-hide templates/calculations)
  3. Set visibility (hide archives)
  → Result: Professional, organized workbook

Before deleting sheet:
  1. Check if it's the last visible sheet (get-visibility on all sheets)
  2. If last visible, don't delete (error prevention)
  → Result: Avoid Excel error
```

### 5. Error Prevention Strategies
Validate before calling:

```
RGB color validation:
  if (red < 0 || red > 255 || green < 0 || green > 255 || blue < 0 || blue > 255) {
    → Error: "RGB values must be 0-255"
  }

Sheet existence check:
  1. Call: worksheet(action: "list")
  2. Verify sheetName exists in list
  3. Then call: set-tab-color or set-visibility
  → Prevents "sheet not found" errors

Last visible sheet protection:
  1. Get visibility of all sheets
  2. Count visible sheets
  3. If count === 1 and attempting to hide that sheet:
     → Error: "Cannot hide last visible sheet"
```

---

## CLI Commands (Secondary Use Case)

### Tab Color Commands

**sheet-set-tab-color** - Set worksheet tab color
```powershell
excelcli sheet-set-tab-color <file.xlsx> <sheet-name> <red> <green> <blue>

# Examples
excelcli sheet-set-tab-color "Report.xlsx" "Sales" 255 0 0       # Red
excelcli sheet-set-tab-color "Report.xlsx" "Expenses" 0 255 0   # Green
excelcli sheet-set-tab-color "Report.xlsx" "Summary" 0 0 255    # Blue
excelcli sheet-set-tab-color "Report.xlsx" "Data" 255 165 0     # Orange
```

**sheet-get-tab-color** - Get worksheet tab color
```powershell
excelcli sheet-get-tab-color <file.xlsx> <sheet-name>

# Example output
Sheet: Sales
Color: #FF0000 (Red: 255, Green: 0, Blue: 0)
```

**sheet-clear-tab-color** - Clear worksheet tab color
```powershell
excelcli sheet-clear-tab-color <file.xlsx> <sheet-name>

# Example
excelcli sheet-clear-tab-color "Report.xlsx" "Sales"
```

### Visibility Commands

**sheet-set-visibility** - Set worksheet visibility level
```powershell
excelcli sheet-set-visibility <file.xlsx> <sheet-name> <visible|hidden|veryhidden>

# Examples
excelcli sheet-set-visibility "Report.xlsx" "Data" hidden
excelcli sheet-set-visibility "Report.xlsx" "Data" veryhidden
excelcli sheet-set-visibility "Report.xlsx" "Data" visible
```

**sheet-get-visibility** - Get worksheet visibility level
```powershell
excelcli sheet-get-visibility <file.xlsx> <sheet-name>

# Example output
Sheet: Data
Visibility: Hidden
```

**sheet-show** - Show a hidden worksheet
```powershell
excelcli sheet-show <file.xlsx> <sheet-name>

# Example
excelcli sheet-show "Report.xlsx" "Data"
```

**sheet-hide** - Hide a worksheet (user can unhide via UI)
```powershell
excelcli sheet-hide <file.xlsx> <sheet-name>

# Example
excelcli sheet-hide "Report.xlsx" "Data"
```

**sheet-very-hide** - Very hide a worksheet (requires code to unhide)
```powershell
excelcli sheet-very-hide <file.xlsx> <sheet-name>

# Example
excelcli sheet-very-hide "Report.xlsx" "Calculations"
```

---

## Testing Strategy

### Unit Tests
- RGB to BGR conversion logic
- Enum mapping for SheetVisibility
- Input validation (RGB range 0-255)

### Integration Tests

**Tab Color Tests:**
- Test setting valid RGB values and verify color is set correctly
- Test RGB to BGR conversion accuracy
- Test clearing tab color removes color
- Test getting color when none is set returns HasColor = false
- Test invalid RGB values (< 0 or > 255) return error

**Visibility Tests:**
- Test setting each visibility level (Visible, Hidden, VeryHidden)
- Test getting visibility returns correct state
- Test VeryHidden can be unhidden programmatically
- Test convenience methods (Show, Hide, VeryHide) call SetVisibility correctly

---

## Breaking Changes

**None.** All new functionality is additive to existing SheetCommands.

---

## Implementation Checklist

### Phase 1: Core Implementation
- [ ] Add `SheetVisibility` enum to Core
- [ ] Add `TabColorResult` and `SheetVisibilityResult` classes
- [ ] Update `ISheetCommands` interface with new methods
- [ ] Implement tab color operations in `SheetCommands.cs`
- [ ] Implement visibility operations in `SheetCommands.cs`
- [ ] Add integration tests for tab color
- [ ] Add integration tests for visibility

### Phase 2: CLI Implementation
- [ ] Create CLI wrapper for tab color commands
- [ ] Create CLI wrapper for visibility commands
- [ ] Add tab color commands to `Program.cs` routing
- [ ] Add visibility commands to `Program.cs` routing
- [ ] Add CLI tests for new commands
- [ ] Update user documentation

### Phase 3: MCP Server Implementation
- [ ] Add tab color actions to `ExcelWorksheetTool`
- [ ] Add visibility actions to `ExcelWorksheetTool`
- [ ] Update `server.json` configuration
- [ ] Add MCP integration tests
- [ ] Update MCP prompts documentation

### Phase 4: Documentation
- [ ] Update README.md with new features
- [ ] Update copilot instructions
- [ ] Add examples to documentation
- [ ] Update INSTALLATION.md if needed

---

## Success Criteria

- [ ] All 8 new Core methods implemented and tested
- [ ] RGB ↔ BGR conversion working correctly
- [ ] All 3 visibility levels (Visible, Hidden, VeryHidden) working
- [ ] CLI commands functional for both features
- [ ] MCP Server actions working via protocol
- [ ] Integration tests passing (95%+ coverage)
- [ ] Documentation complete

---

## Common Color Presets

For user convenience, here are common Excel tab colors:

| Color Name | RGB | Hex | BGR (Excel) |
|------------|-----|-----|-------------|
| Red | 255, 0, 0 | #FF0000 | 0x0000FF |
| Green | 0, 255, 0 | #00FF00 | 0x00FF00 |
| Blue | 0, 0, 255 | #0000FF | 0xFF0000 |
| Yellow | 255, 255, 0 | #FFFF00 | 0x00FFFF |
| Orange | 255, 165, 0 | #FFA500 | 0x00A5FF |
| Purple | 128, 0, 128 | #800080 | 0x800080 |
| Pink | 255, 192, 203 | #FFC0CB | 0xCBC0FF |
| Teal | 0, 128, 128 | #008080 | 0x808000 |
| Light Blue | 173, 216, 230 | #ADD8E6 | 0xE6D8AD |
| Light Green | 144, 238, 144 | #90EE90 | 0x90EE90 |

**Note:** BGR values shown for reference. API accepts RGB values and handles conversion internally.

---

## Future Enhancements (Out of Scope)

- **Hex Color Input** - Accept `#FF0000` format directly (currently requires RGB conversion)
- **Get All Colors Action** - Single call to retrieve all sheet colors: `get-all-tab-colors` → `[{sheetName, red, green, blue, hexColor}]`
- **Theme Color Support** - Use Excel's theme colors instead of RGB
- **Tab Icons** - Excel 365 supports custom tab icons (not widely used)
- **Tab Position** - Reorder tabs programmatically
- **Tab Group Protection** - Protect groups of tabs together
- **Bulk Color Operations** - Set same color for multiple sheets in one call

These features can be considered in future iterations if user demand exists.
