# MCP Tool ↔ Core Interface Delta Analysis

> Generated 2026-02-07 — Comprehensive comparison of hand-written MCP tool files vs Core interface files  
> Purpose: Identify what NEW attributes are needed on Core interfaces to auto-generate MCP tools.

---

## 1. Existing Core Attributes (src/ExcelMcp.Core/Attributes/)

| Attribute | Target | Properties | Purpose |
|-----------|--------|-----------|---------|
| `[ServiceCategory("cat", "Pascal")]` | Interface | `Category`, `PascalName` | Service routing category |
| `[McpTool("excel_xxx")]` | Interface, Method | `ToolName` | Maps interface → MCP tool name |
| `[NoSession]` | Interface | (none) | Marks as not requiring session |
| `[ServiceAction("name")]` | Method | `Action` | Overrides action name derivation |
| `[FromString("exposed")]` | Parameter | `ExposedName` | Exposes enum as string parameter |
| `[FileOrValue("File")]` | Parameter | `FileSuffix` | Creates dual param (value + file) |
| `[RequiredParameter]` | Parameter | (none) | Marks parameter as required |

**Generator Model (ServiceInfoExtractor → ServiceInfo/MethodInfo/ParameterInfo) already extracts:**
- Category, CategoryPascal, McpToolName, NoSession
- MethodName, ActionName, ReturnType, Parameters, XmlDocSummary, HasBatchParameter
- Parameter: Name, TypeName, HasDefault, DefaultValue, IsFileOrValue, FileSuffix, IsFromString, ExposedName, IsRequired, IsEnum, XmlDocDescription

---

## 2. MCP Tool Attributes NOT in Core (The Delta)

### 2.1 Tool-Level Metadata (on `[McpServerTool]` / `[McpMeta]`)

| MCP Attribute | Example | In Core? | Notes |
|---------------|---------|----------|-------|
| `Name = "excel_xxx"` | `powerquery` | ✅ YES | `[McpTool]` attribute already exists |
| `Title = "..."` | `"Excel Power Query Operations"` | ❌ **NO** | Human-readable title for MCP clients |
| `Destructive = true/false` | `true` (most), `false` (calc mode) | ❌ **NO** | Whether tool modifies data |
| `[McpMeta("category", "...")]` | `"data"`, `"analysis"`, `"query"` | ❌ **NO** | UI grouping category |
| `[McpMeta("requiresSession", bool)]` | `true` / `false` | ⚠️ PARTIAL | `[NoSession]` exists but is inverted; need value on most |
| `[McpMeta("fileFormat", ".xlsm")]` | VBA tool only | ❌ **NO** | Required file format constraint |

### 2.2 Method-Level Summary Text (LLM Guidance)

**Every MCP tool has a rich `<summary>` on the method** providing LLM-targeted guidance (best practices, workflows, related tools). The Core interface `<summary>` is usually shorter and more technical. These are **different** docs serving different audiences.

| Tool | MCP Summary Lines | Core Summary Lines | Same? |
|------|-------------------|-------------------|-------|
| range | 18 lines (best practices, data format, named ranges, copy ops, number formats) | 14 lines (similar but shorter) | ⚠️ SIMILAR but MCP has more LLM-specific tips |
| powerquery | 18 lines (test-first workflow, datetime columns, M-code formatting, destinations) | 12 lines (similar core) | ⚠️ SIMILAR |
| chart | 24 lines (overlapping data avoidance, positioning, chart types, create options) | Not on interface | ❌ NO - only on MCP |
| file | 10 lines (workflow, session reuse, timeout) | 3 lines (minimal) | ❌ VERY DIFFERENT |
| table | 14 lines (best practices, data model workflow, DAX-backed tables, CSV append) | Not examined | ❌ DIFFERENT |

**Verdict:** The MCP summary is LLM-specific guidance text. Core summaries describe the API technically. A new `[McpSummary("...")]` or `[LlmGuidance("...")]` attribute (or conventionally a separate doc block) would be needed if we want to auto-generate the MCP `<summary>`.

### 2.3 Parameter Description Differences

MCP `<param>` docs are often **different** from Core `<param>` docs — the MCP versions are LLM-optimized with examples and constraints:

| Parameter | MCP Description | Core Description |
|-----------|----------------|-----------------|
| `sessionId` | `"Session ID from file 'open'. Required for all actions."` | Not present (batch param instead) |
| `sheetName` | `"Name of worksheet - REQUIRED for cell addresses, use empty string for named ranges"` | `"Name of the worksheet containing the range"` |
| `rangeAddress` | `"Cell range address (e.g., 'A1', 'A1:D10') or named range name (e.g., 'SalesData')"` | `"Cell range address (e.g., 'A1', 'A1:D10')"` |
| `values` | `"2D array of values - rows are outer array, columns are inner array (e.g., [[1,2,3],[4,5,6]])"` | `"2D array of values to set"` (shorter) |
| `path` | `"Full Windows path to Excel file. ASK USER for the path - do not guess"` | `"Path to the Excel file to validate"` |

**Verdict:** MCP param docs need either a separate attribute or a convention for providing LLM-optimized descriptions. Could reuse the existing Core `<param>` XML docs if they're made richer.

---

## 3. Per-Tool Comparison Table

### Legend
- **PP** = Pre-processing logic in method body
- **DV** = Non-null DefaultValue (not just `[DefaultValue(null)]`)
- **Meta** = Extra `[McpMeta]` beyond category + requiresSession

| # | MCP Tool Name | Title | Destr. | Category | reqSession | Meta | Core Interface(s) | PP Logic | Notes |
|---|--------------|-------|--------|----------|-----------|------|-------------------|----------|-------|
| 1 | `file` | Excel File Operations | true | session | **false** | | `IFileCommands` (partial) | ✅ Extensive custom routing, path validation, timeout conversion | **Most custom** - manual Open/Close/Create/List, no RouteAction |
| 2 | `worksheet` | Excel Worksheet Operations | true | structure | true | | `ISheetCommands` | ✅ CopyToFile/MoveToFile use ForwardToServiceNoSession; param remapping (sheetName→oldName/sourceName) | Cross-file ops bypass session |
| 3 | `worksheet_style` | Excel Worksheet Style Operations | true | structure | true | | `ISheetStyleCommands` | Minimal (visibility?.ToString()) | Clean routing |
| 4 | `range` | Excel Range Operations | true | data | true | | `IRangeCommands` | ✅ `values as List<List<object>>` cast | Type cast |
| 5 | `range_edit` | Excel Range Edit Operations | true | data | true | | `IRangeEditCommands` | ✅ `BuildFindOptions()`, `BuildReplaceOptions()` | Object construction from flat params |
| 6 | `range_format` | Excel Range Format Operations | true | data | true | | `IRangeFormatCommands` | Minimal remapping (e.g., `backgroundColor`→`fillColor`) | Param name remapping |
| 7 | `range_link` | Excel Range Link Operations | true | data | true | | `IRangeLinkCommands` | Minimal (`isLocked`→`locked`) | Param name remapping |
| 8 | `table` | Excel Table Operations | true | data | true | | `ITableCommands` | ✅ `ParseCsvToRows(csvData)`, `rows as List<List<object>>`, multi-param reuse (hasHeaders→showTotals, styleName→totalFunction, newName→columnName) | Significant param overloading |
| 9 | `table_column` | Excel Table Column Operations | true | data | true | | `ITableColumnCommands` | ✅ `int.TryParse(columnPosition)`, `ParseJsonList()` / `DeserializeJson<List<TableSortColumn>>()` conditional on action | JSON parsing + type conversion |
| 10 | `powerquery` | Excel Power Query Operations | true | query | true | | `IPowerQueryCommands` | ✅ `TimeSpan.FromSeconds(refreshTimeoutSeconds)`, `loadDestination?.ToString()`, `oldName` mapping for rename | TimeSpan conversion, enum→string |
| 11 | `connection` | Excel Data Connection Operations | true | query | true | | `IConnectionCommands` | Minimal (timeout: null) | Clean routing |
| 12 | `namedrange` | Excel Named Range Operations | true | data | true | | `INamedRangeCommands` | ✅ `value` doubles as `reference` for create/update | Param reuse |
| 13 | `pivottable` | Excel PivotTable Operations | true | analysis | true | | `IPivotTableCommands` | ✅ `tableName: sourceTableName ?? dataModelTableName` | Param merging |
| 14 | `pivottable_field` | Excel PivotTable Field Operations | true | analysis | true | | `IPivotTableFieldCommands` | ✅ `ParseJsonList(filterValues)`, `aggregationFunction?.ToString()`, `sortDirection?.ToString()`, `dateGroupingInterval?.ToString()` | JSON parsing + enum→string |
| 15 | `pivottable_calc` | Excel PivotTable Calc Operations | true | analysis | true | | `IPivotTableCalcCommands` | ✅ `memberType?.ToString()`, `fieldName` → `memberName` remapping | Enum→string, param remapping |
| 16 | `chart` | Excel Chart Operations | true | analysis | true | | `IChartCommands` | ✅ targetRange pre-processing: sets left/top to 0 when targetRange provided | Positioning logic |
| 17 | `chart_config` | Excel Chart Configuration | true | analysis | true | | `IChartConfigCommands` | Minimal (trendlineType→type remapping) | Clean routing, many params |
| 18 | `slicer` | Excel Slicer Operations | true | analysis | true | | `ISlicerCommands` | ✅ Auto-generate slicerName from field/column, `ParseJsonListOrSingle(selectedItems)` | Name generation + JSON parsing |
| 19 | `vba` | Excel VBA Operations | true | automation | true | `fileFormat=.xlsm` | `IVbaCommands` | ✅ `SplitCsvParameters(parameters)`, `moduleName` → `procedureName` remapping | CSV split, param remapping |
| 20 | `datamodel` | Excel Data Model Operations | true | analysis | true | | `IDataModelCommands` | ✅ `tableName` → `oldName` remapping for rename | Param remapping |
| 21 | `datamodel_rel` | Excel Data Model Relationship Operations | true | analysis | true | | `IDataModelRelCommands` | Minimal | Clean routing |
| 22 | `conditionalformat` | Excel Conditional Formatting | true | structure | true | | `IConditionalFormattingCommands` | Minimal | Clean routing |
| 23 | `calculation_mode` | Excel Calculation Mode Control | **false** | settings | true | | `ICalculationCommands` | Minimal | **Only non-destructive tool** |

---

## 4. Pre-Processing Patterns (What Can't Be Auto-Generated)

These are the transformations in MCP tool method bodies that go beyond simple routing:

### 4.1 Type Conversions

| Pattern | Used In | What It Does |
|---------|---------|-------------|
| `TimeSpan.FromSeconds(int?)` | PowerQuery, File | Converts int seconds param → TimeSpan |
| `enum?.ToString()` | PowerQuery, PivotTableField, PivotTableCalc, WorksheetStyle | Converts nullable enum → string for Service routing |
| `int.TryParse(string)` | TableColumn | Converts string columnPosition → int? |
| `values as List<List<object>>` | Range, Table | Cast from `List<List<object?>>` to `List<List<object>>` |

### 4.2 JSON Parsing

| Pattern | Used In | What It Does |
|---------|---------|-------------|
| `ParseJsonList(json)` | PivotTableField, TableColumn | Parse `["a","b"]` → `List<string>` |
| `ParseJsonListOrSingle(json)` | Slicer | Parse JSON array or treat as single value |
| `DeserializeJson<T>(json)` | TableColumn | Parse `[{columnName, ascending}]` → `List<TableSortColumn>` |
| `ParseCsvToRows(csv)` | Table | Parse multi-line CSV → `List<List<object?>>` |
| `SplitCsvParameters(csv)` | VBA | Split comma-separated string → string[] |

### 4.3 Parameter Mapping / Overloading

| Pattern | Used In | What It Does |
|---------|---------|-------------|
| Single param → multiple Core params | Table (`newName`→columnName, `styleName`→totalFunction), NamedRange (`value`→reference), Chart (`rangeAddress`=multiple uses) | MCP has fewer params than Core methods; same param serves multiple roles |
| Conditional param derivation | Worksheet (`sheetName`→oldName/sourceName), PowerQuery (`queryName`→oldName for rename), DataModel (`tableName`→oldName) | One MCP param maps to different Core params based on action |
| Auto-generation | Slicer (auto-generate slicerName from fieldName/columnName if not provided) | Business logic in MCP layer |
| Pre-processing | Chart (set left/top=0 when targetRange provided) | Derived defaults |

### 4.4 Special Routing

| Pattern | Used In | What It Does |
|---------|---------|-------------|
| ForwardToServiceNoSession | Worksheet (CopyToFile, MoveToFile) | Some actions bypass session routing |
| Completely custom (no RouteAction) | File | Manual switch on all actions, custom JSON responses |
| Object construction | RangeEdit (`BuildFindOptions`, `BuildReplaceOptions`) | Constructs complex objects from flat params |

---

## 5. DefaultValue Summary

Most parameters use `[DefaultValue(null)]`. Non-null defaults:

| Tool | Parameter | Default | Core Default |
|------|-----------|---------|-------------|
| file | `save` | `false` | N/A (custom) |
| file | `show` | `false` | N/A (custom) |
| file | `timeoutSeconds` | `300` | N/A (custom) |
| table | `hasHeaders` | `true` | `true` (same) |
| table | `visibleOnly` | `false` | N/A |
| table_column | `ascending` | `true` | `true` (same) |
| slicer | `clearFirst` | `true` | `true` (same) |
| datamodel_rel | `active` | `true` | `true` (same) |
| calculation_mode | (Destructive) | `false` | N/A |

---

## 6. New Attributes Needed for Auto-Generation

Based on the delta analysis, these new attributes would be required on Core interfaces to fully auto-generate MCP tools:

### 6.1 Must-Have (No Workaround)

| New Attribute | Target | Example | Why Needed |
|---------------|--------|---------|-----------|
| `[McpTitle("...")]` | Method | `[McpTitle("Excel Power Query Operations")]` | Tool title for MCP clients. No current equivalent. |
| `[Destructive(bool)]` | Method or Interface | `[Destructive(true)]` | MCP SDK requires this on `[McpServerTool]`. Default true but calc_mode is false. |
| `[McpCategory("...")]` | Interface | `[McpCategory("analysis")]` | UI grouping. Different from ServiceCategory (routing). Values: data, analysis, query, structure, session, automation, settings |
| `[McpMeta("key", value)]` | Interface or Method | `[McpMeta("fileFormat", ".xlsm")]` | Arbitrary metadata (only VBA uses fileFormat currently) |

### 6.2 Should-Have (For Quality)

| New Attribute | Target | Example | Why Needed |
|---------------|--------|---------|-----------|
| `[McpSummary("...")]` or separate doc file | Interface/Method | Multi-line LLM guidance text | MCP method summary is different from Core interface summary. Could also use a convention (e.g., `/// <mcpsummary>`) or external .md files. |
| `[McpParamDescription("...")]` | Parameter | LLM-optimized param description | MCP param descriptions differ from Core. Alternative: make Core `<param>` docs richer. |
| `[McpParamName("...")]` | Parameter | `[McpParamName("sessionId")]` | MCP exposes `sessionId` but Core uses `IExcelBatch batch`. Need to know the MCP-facing name. Generator already handles batch→sessionId. |
| `[DefaultValue(val)]` on Core params | Parameter | Already exists via C# default values | Generator already reads these. ✅ |

### 6.3 Pre-Processing Attributes (For Complex Logic)

These patterns are harder to express as attributes and may need **code hooks** or **transform attributes**:

| Pattern | Proposed Solution | Complexity |
|---------|------------------|-----------|
| `TimeSpan.FromSeconds()` | `[TimeSpanFromSeconds]` on the Core param (already `TimeSpan timeout` in Core) | Low - generator can emit conversion |
| `enum?.ToString()` | Already handled by `[FromString]` attribute | ✅ Done |
| `ParseJsonList` | `[JsonList]` attribute on `List<string>` param | Medium - needs to know the MCP type is `string` but Core type is `List<string>` |
| `ParseJsonListOrSingle` | `[JsonListOrSingle]` attribute variant | Medium |
| `DeserializeJson<T>` | `[JsonDeserialize]` attribute | Medium |
| `ParseCsvToRows` | `[CsvToRows]` attribute | Medium |
| `SplitCsvParameters` | `[CsvSplit]` attribute | Low |
| `BuildFindOptions/BuildReplaceOptions` | `[ExpandToObject]` or flatten params in Core | Hard - would need to flatten Core method signatures |
| Param overloading (same MCP param → multiple Core params) | `[McpParamAlias("mcpName")]` on each Core param that uses same MCP input | Medium |
| Auto-generation logic (slicer name) | Custom hook / `[AutoGenerate]` | Hard |
| Conditional logic (chart targetRange → left/top=0) | Custom hook | Hard |
| No-session routing (CopyToFile) | Already signaled by no `IExcelBatch` param in Core | ✅ Done |
| Fully custom tools (file) | Cannot auto-generate. Keep as hand-written. | N/A |

---

## 7. Recommendations

### Tier 1: Add These Attributes Now (Easy wins, unblock generation for 15+ tools)

```csharp
// On interface or method - new attributes
[McpTitle("Excel Power Query Operations")]     // → McpServerTool.Title
[Destructive(true)]                             // → McpServerTool.Destructive
[McpCategory("query")]                          // → McpMeta("category", ...)

// These already exist and work:
[McpTool("powerquery")]                   // → McpServerTool.Name
[NoSession]                                     // → McpMeta("requiresSession", false)
```

### Tier 2: Enhance Param Transforms (Medium effort, covers 80% of pre-processing)

```csharp
// New parameter-level attributes
[TimeSpanSeconds]                               // int → TimeSpan conversion
[JsonList]                                      // string → List<string> via ParseJsonList
[JsonListOrSingle]                              // string → List<string> via ParseJsonListOrSingle  
[JsonDeserialize]                               // string → T via DeserializeJson<T>
[CsvRows]                                       // string → List<List<object?>> via ParseCsvToRows
[CsvSplit]                                      // string → string[] via SplitCsvParameters
[McpParamAlias("mcpParamName")]                 // Multiple Core params sharing one MCP param
```

### Tier 3: Keep Hand-Written (Complex custom logic)

These tools have too much custom logic to auto-generate and should remain hand-written:

1. **`file`** - Completely custom routing, no RouteAction, custom JSON responses
2. **`worksheet`** - CopyToFile/MoveToFile no-session routing  
3. **`table`** - Heavy param overloading (hasHeaders→showTotals, styleName→totalFunction)
4. **`chart`** - targetRange pre-processing logic
5. **`slicer`** - Auto-name generation

The remaining **~17 tools** could be auto-generated with Tier 1 + Tier 2 attributes.

---

## 8. Summary Statistics

| Metric | Count |
|--------|-------|
| Total MCP tools | 23 (including file) |
| Tools using `[McpServerToolType]` | 23 |
| Tools with `Destructive = true` | 22 |
| Tools with `Destructive = false` | 1 (calculation_mode) |
| Unique `[McpMeta("category")]` values | 7 (data, analysis, query, structure, session, automation, settings) |
| Tools with `requiresSession = true` | 21 |
| Tools with `requiresSession = false` | 2 (file, file has it explicit) |
| Tools with pre-processing logic | 15 |
| Tools with clean/minimal routing | 8 |
| Tools likely auto-generatable (Tier 1+2) | ~17-18 |
| Tools requiring hand-written code | ~5 |
| New Core attributes needed (Tier 1) | 3 (`McpTitle`, `Destructive`, `McpCategory`) |
| New Core attributes needed (Tier 2) | 6-7 (param transform attributes) |
| Existing attributes sufficient | 7 (`ServiceCategory`, `McpTool`, `NoSession`, `ServiceAction`, `FromString`, `FileOrValue`, `RequiredParameter`) |
