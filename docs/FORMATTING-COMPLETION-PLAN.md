# FORMATTING & VALIDATION - COMPLETION PLAN

**Status:** Core implementation complete (1,770 lines), MCP/CLI/Tests/Docs pending  
**Date:** 2025-01-20  
**Estimated Remaining:** 6-8 hours

---

## ✅ COMPLETED (Phase 1: Core Implementation)

### Files Created
1. `NumberFormatPresets.cs` - 39 preset constants
2. `FormattingEnums.cs` - 7 enums
3. `FormattingOptions.cs` - 4 option classes  
4. `ResultTypes.cs` - 6 new result types added
5. `IRangeCommands.cs` - 21 new method signatures added
6. `RangeCommands.NumberFormatting.cs` - 3 methods (220 lines)
7. `RangeCommands.VisualFormatting.cs` - 5 methods (280 lines)
8. `RangeCommands.Borders.cs` - 3 methods + helpers (240 lines)
9. `RangeCommands.Alignment.cs` - 2 methods + helpers (180 lines)
10. `RangeCommands.AutoFit.cs` - 4 methods (170 lines)
11. `RangeCommands.Validation.cs` - 4 methods + helpers (350 lines)
12. `ITableCommands.cs` - 8 new method signatures added
13. `TableCommands.Formatting.cs` - 8 methods (370 lines)

### Build Status
✅ Core project builds successfully (with TreatWarningsAsErrors=false)

---

## 📝 TODO: Phase 2 - MCP Server Integration

### Update ExcelRangeTool.cs (21 new actions)

**File:** `src/ExcelMcp.McpServer/Tools/ExcelRangeTool.cs`

**Step 1:** Update action RegularExpression to include:
```csharp
[RegularExpression(@"^(get-values|...|get-number-formats|set-number-format|set-number-formats|get-font|set-font|get-background-color|set-background-color|clear-background-color|get-borders|set-borders|clear-borders|get-alignment|set-alignment|autofit-columns|autofit-rows|set-column-width|set-row-height|get-validation|add-validation|modify-validation|remove-validation)$")]
```

**Step 2:** Add new action descriptions to Description attribute

**Step 3:** Add parameters:
```csharp
[Description("Number format code (for set-number-format)")]
string? formatCode = null,

[Description("2D array of format codes (for set-number-formats) - JSON string")]
string? formatsJson = null,

[Description("Font options JSON (for set-font): { name, size, bold, italic, color, underline, strikethrough }")]
string? fontOptionsJson = null,

[Description("RGB color as integer (for set-background-color)")]
int? color = null,

[Description("Border options JSON (for set-borders): { style, weight, color, applyToAll, top, bottom, left, right }")]
string? borderOptionsJson = null,

[Description("Alignment options JSON (for set-alignment): { horizontal, vertical, wrapText, indent, orientation }")]
string? alignmentOptionsJson = null,

[Description("Column width in points (for set-column-width)")]
double? width = null,

[Description("Row height in points (for set-row-height)")]
double? height = null,

[Description("Validation rule JSON (for add-validation/modify-validation)")]
string? validationRuleJson = null
```

**Step 4:** Add switch cases (example for number formatting):
```csharp
"get-number-formats" => await GetNumberFormatsAsync(commands, excelPath, sheetName, rangeAddress, batchId),
"set-number-format" => await SetNumberFormatAsync(commands, excelPath, sheetName, rangeAddress, formatCode!, batchId),
"set-number-formats" => await SetNumberFormatsAsync(commands, excelPath, sheetName, rangeAddress, formatsJson!, batchId),
// ... repeat for all 21 new actions
```

**Step 5:** Implement helper methods (example):
```csharp
private static async Task<string> GetNumberFormatsAsync(
    RangeCommands commands, string excelPath, string? sheetName, string rangeAddress, string? batchId)
{
    if (string.IsNullOrEmpty(sheetName))
        throw new McpException("sheetName required for get-number-formats");
    
    var batch = await GetOrCreateBatchAsync(excelPath, batchId);
    try
    {
        var result = await commands.GetNumberFormatsAsync(batch, sheetName, rangeAddress);
        if (!result.Success)
            throw new McpException($"get-number-formats failed: {result.ErrorMessage}");
        return JsonSerializer.Serialize(result, JsonOptions);
    }
    finally
    {
        await DisposeBatchIfNotSessionAsync(batch, batchId);
    }
}

// ... 20 more helper methods
```

**Estimated Time:** 3-4 hours (21 actions × 10-15 min each)

---

### Update ExcelTableTool.cs (8 new actions)

**File:** `src/ExcelMcp.McpServer/Tools/ExcelTableTool.cs`

**Step 1:** Update action RegularExpression to include:
```csharp
set-column-number-format, set-header-font, set-data-font, set-header-color, set-banded-rows, autofit-columns, set-column-validation, remove-column-validation
```

**Step 2:** Add parameters (reuse from ExcelRangeTool):
- `formatCode`, `fontOptionsJson`, `color`, `validationRuleJson`
- Additional: `color1`, `color2` for `set-banded-rows`

**Step 3:** Add switch cases and helper methods

**Estimated Time:** 2 hours (8 actions × 15 min each)

---

## 📝 TODO: Phase 3 - CLI Commands

### Update Program.cs

**File:** `src/ExcelMcp.CLI/Program.cs`

**Add routing for 29 new commands:**
```csharp
"range-get-number-formats" => await rangeCommands.GetNumberFormatsAsync(...),
"range-set-number-format" => await rangeCommands.SetNumberFormatAsync(...),
// ... 27 more commands
```

### Create CLI Helper Methods

**Pattern for each command (example):**
```csharp
private static async Task<int> RangeGetNumberFormatsAsync(
    IRangeCommands commands, string[] args)
{
    var parser = new ArgParser(args);
    var excelPath = parser.GetRequired("--file");
    var sheetName = parser.GetRequired("--sheet");
    var rangeAddress = parser.GetRequired("--range");
    
    await using var batch = await ExcelSession.BeginBatchAsync(excelPath);
    var result = await commands.GetNumberFormatsAsync(batch, sheetName, rangeAddress);
    
    if (!result.Success)
    {
        Console.Error.WriteLine($"Error: {result.ErrorMessage}");
        return 1;
    }
    
    // Print 2D array as CSV or JSON
    Console.WriteLine(JsonSerializer.Serialize(result.Formats));
    return 0;
}
```

**Estimated Time:** 3 hours (29 commands × 6 min each)

---

## 📝 TODO: Phase 4 - Integration Tests

### Create Test Files

**Structure:**
```
tests/ExcelMcp.Core.Tests/Integration/Commands/Range/
    RangeCommandsTests.NumberFormatting.cs (5-7 tests)
    RangeCommandsTests.VisualFormatting.cs (8-10 tests)
    RangeCommandsTests.Validation.cs (10-12 tests)

tests/ExcelMcp.Core.Tests/Integration/Commands/Table/
    TableCommandsTests.Formatting.cs (8-10 tests)

tests/ExcelMcp.McpServer.Tests/Integration/Tools/
    ExcelRangeToolTests.Formatting.cs (10-12 tests)
    ExcelTableToolTests.Formatting.cs (4-6 tests)
```

### Test Pattern (Example)
```csharp
[Fact]
[Trait("Category", "Integration")]
[Trait("Feature", "RangeFormatting")]
public async Task SetNumberFormat_CurrencyFormat_AppliesCorrectly()
{
    // Arrange
    var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
        nameof(RangeCommandsTests), nameof(SetNumberFormat_CurrencyFormat_AppliesCorrectly), 
        _tempDir, ".xlsx");
    
    await using var batch = await ExcelSession.BeginBatchAsync(testFile);
    
    // Set some test values
    var values = new List<List<object?>> { new() { 1234.56 }, new() { 7890.12 } };
    await _rangeCommands.SetValuesAsync(batch, "Sheet1", "A1:A2", values);
    
    // Act - Apply currency format
    var result = await _rangeCommands.SetNumberFormatAsync(
        batch, "Sheet1", "A1:A2", NumberFormatPresets.Currency);
    
    // Assert - Verify success
    Assert.True(result.Success, $"SetNumberFormat failed: {result.ErrorMessage}");
    
    // Verify format was applied
    var getResult = await _rangeCommands.GetNumberFormatsAsync(batch, "Sheet1", "A1:A2");
    Assert.True(getResult.Success);
    Assert.Equal(2, getResult.Formats.Count);
    Assert.Equal(NumberFormatPresets.Currency, getResult.Formats[0][0]);
    Assert.Equal(NumberFormatPresets.Currency, getResult.Formats[1][0]);
    
    await batch.SaveAsync();
}
```

**Test Count:** ~35-42 tests
**Estimated Time:** 4-5 hours

---

## 📝 TODO: Phase 5 - Documentation

### Update COMMANDS.md

**Add 29 new commands with examples:**

```markdown
#### range-get-number-formats
Get number format codes from a range.

**Usage:**
```bash
excelcli range-get-number-formats --file data.xlsx --sheet Sheet1 --range A1:D10
```

**Output:** 2D array of format codes (JSON)

---

#### range-set-number-format
Apply uniform number format to range.

**Usage:**
```bash
excelcli range-set-number-format --file data.xlsx --sheet Sheet1 --range A1:A10 --format "$#,##0.00"
```

**Common Formats:**
- Currency: `$#,##0.00`
- Percentage: `0.00%`
- Date: `m/d/yyyy`
- Text: `@`

---

... (27 more commands)
```

### Update README.md

**Update tool counts:**
```markdown
| **excel_range** | 45 actions | Range operations + formatting + validation |
| **excel_table** | 30 actions | Table management + formatting |
```

**Add formatting capabilities section:**
```markdown
### Formatting & Validation Features

**Number Formats:** Apply currency, percentage, date formats to ranges
**Font Styles:** Bold, italic, underline, font family and size
**Colors:** Background colors with RGB control
**Borders:** Apply borders with multiple styles and weights
**Alignment:** Horizontal, vertical, text wrapping, rotation
**AutoFit:** Automatic column/row sizing
**Data Validation:** Dropdown lists, number ranges, date constraints, custom formulas
```

### Update Tool Prompts

**File:** `src/ExcelMcp.McpServer/Prompts/RangePrompts.cs`

**Add formatting examples:**
```csharp
FORMATTING (21 new actions):
- Number formats: get-number-formats, set-number-format (use NumberFormatPresets: Currency, Percentage, DateShort)
- Fonts: get-font, set-font (name, size, bold, italic, color, underline, strikethrough)
- Colors: get-background-color, set-background-color (RGB int), clear-background-color
- Borders: get-borders, set-borders (style, weight, color), clear-borders
- Alignment: get-alignment, set-alignment (horizontal, vertical, wrap, indent, orientation)
- AutoFit: autofit-columns, autofit-rows, set-column-width, set-row-height
- Validation: get-validation, add-validation, modify-validation, remove-validation

EXAMPLES:
1. Format column as currency: set-number-format, sheet="Sales", range="D:D", formatCode="$#,##0.00"
2. Bold header row: set-font, sheet="Data", range="1:1", fontOptions={ bold: true, size: 12 }
3. Apply dropdown list: add-validation, range="A2:A100", validation={ type: "List", formula1: "Option1,Option2,Option3" }
```

**Estimated Time:** 1-2 hours

---

## 📋 COMPLETION CHECKLIST

### MCP Server (Estimated: 3-4 hours)
- [ ] Update ExcelRangeTool.cs RegularExpression with 21 new actions
- [ ] Add parameters to ExcelRange method signature (formatCode, fontOptionsJson, borderOptionsJson, alignmentOptionsJson, width, height, validationRuleJson)
- [ ] Add 21 switch cases in ExcelRange method
- [ ] Implement 21 helper methods (GetNumberFormatsAsync, SetNumberFormatAsync, etc.)
- [ ] Update ExcelTableTool.cs with 8 new actions
- [ ] Implement 8 table helper methods

### CLI (Estimated: 3 hours)
- [ ] Add 29 new command routing cases in Program.cs
- [ ] Implement 29 CLI helper methods
- [ ] Add argument parsing for all new commands
- [ ] Test CLI commands manually

### Tests (Estimated: 4-5 hours)
- [ ] Create RangeCommandsTests.NumberFormatting.cs (5-7 tests)
- [ ] Create RangeCommandsTests.VisualFormatting.cs (8-10 tests)
- [ ] Create RangeCommandsTests.Validation.cs (10-12 tests)
- [ ] Create TableCommandsTests.Formatting.cs (8-10 tests)
- [ ] Create ExcelRangeToolTests.Formatting.cs (10-12 tests)
- [ ] Create ExcelTableToolTests.Formatting.cs (4-6 tests)
- [ ] Run all tests, fix failures

### Documentation (Estimated: 1-2 hours)
- [ ] Update COMMANDS.md with 29 new commands (syntax, examples, notes)
- [ ] Update README.md tool counts (excel_range: 45, excel_table: 30)
- [ ] Update README.md with formatting features section
- [ ] Update RangePrompts.cs with formatting examples
- [ ] Update TablePrompts.cs with formatting examples
- [ ] Update main README with capability highlights

### Final Validation (Estimated: 1 hour)
- [ ] Build entire solution (Release mode, TreatWarningsAsErrors=true after XML docs fixed)
- [ ] Run all unit tests (Category=Unit)
- [ ] Run all integration tests (Category=Integration)
- [ ] Manually test 5-10 key formatting scenarios via MCP
- [ ] Manually test 3-5 key formatting scenarios via CLI
- [ ] Check for TODO/FIXME/HACK markers, remove or resolve
- [ ] Git status clean (no uncommitted experimental code)

---

## 🎯 IMPLEMENTATION SEQUENCE (Recommended Order)

1. **MCP Server** (3-4 hours) - Enables testing via MCP protocol
2. **Tests** (4-5 hours) - Validates implementation correctness
3. **CLI** (3 hours) - Adds scripting capability
4. **Documentation** (1-2 hours) - Makes features discoverable
5. **Final Validation** (1 hour) - Ensures quality

**Total Time:** 12-15 hours

---

## 📊 QUICK REFERENCE: Actions by Category

### Number Formatting (3 actions)
1. get-number-formats
2. set-number-format
3. set-number-formats

### Font Formatting (2 actions)
4. get-font
5. set-font

### Color Formatting (3 actions)
6. get-background-color
7. set-background-color
8. clear-background-color

### Border Formatting (3 actions)
9. get-borders
10. set-borders
11. clear-borders

### Alignment (2 actions)
12. get-alignment
13. set-alignment

### AutoFit/Size (4 actions)
14. autofit-columns
15. autofit-rows
16. set-column-width
17. set-row-height

### Validation (4 actions)
18. get-validation
19. add-validation
20. modify-validation
21. remove-validation

### Table Formatting (8 actions)
22. set-column-number-format
23. set-header-font
24. set-data-font
25. set-header-color
26. set-banded-rows
27. autofit-columns (table)
28. set-column-validation
29. remove-column-validation

---

**Next Step:** Begin MCP Server integration with ExcelRangeTool.cs
