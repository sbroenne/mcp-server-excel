# Tool Documentation Cleanup Summary

## Problem Statement

Tool C# files contain extensive "LLM Usage Patterns" sections that duplicate guidance already present in the prompt files (`src/ExcelMcp.McpServer/Prompts/Content/*.md`).

According to `mcp-llm-guidance.instructions.md`, LLM guidance belongs in **prompt files**, not in C# code documentation.

## What Should Stay in C# Code

- **Technical API documentation** for C# developers
- **Parameter descriptions** for MCP schema generation (brief, parameter-focused)
- **Technical constraints** (e.g., "Requires .xlsm files", "Excel COM limitations")
- **Prerequisites** (e.g., "VBA trust must be enabled")
- **Architectural notes** (e.g., "Data operations moved to ExcelRangeTool")

## What Should Be Removed from C# Code

- ❌ "LLM Usage Patterns:" sections
- ❌ Workflow optimization tips ("Use batch mode for...", "75-90% faster")
- ❌ Decision guidance ("Use worksheet for visibility, data-model for DAX")
- ❌ Integration hints ("After loading to data model, use excel_datamodel")
- ❌ Usage examples ("Use 'list' to see...", "Use 'create' to add...")
- ❌ Common mistakes for LLMs
- ❌ Action disambiguation for LLMs

## Files Requiring Cleanup

### High Priority (Most LLM Guidance)

**1. ExcelPowerQueryTool.cs**
- **Class summary** (lines 14-48): ~35 lines of LLM guidance
  - "LLM Usage Patterns:" (11 bullet points)
  - "ATOMIC OPERATIONS:" (5 bullet points)
  - "IMPORTANT FOR DATA MODEL WORKFLOWS:" (4 bullet points)
  - "VALIDATION AND EXECUTION:" (3 bullet points)
- **Method Description** (lines 56-75): ~20 lines of workflow guidance
  - Performance tips
  - Numbered workflow steps
  - Load destination guide
- **Recommendation**: Reduce to 2-3 lines technical description

**2. ExcelWorksheetTool.cs**
- **Class summary** (lines 14-33): ~20 lines with "LLM Usage Patterns:"
- **Keep**: "Data operations have been moved to ExcelRangeTool" (architectural note)
- **Remove**: All usage pattern bullet points
- **Recommendation**: Reduce to 3-4 lines (technical description + architectural note)

**3. ExcelVbaTool.cs**
- **Class summary** (lines 13-29): ~17 lines
- **Keep**: "⚠️ IMPORTANT: Requires .xlsm files!" (technical constraint)
- **Keep**: "Setup Required: Run setup-vba-trust command once" (prerequisite)
- **Remove**: "LLM Usage Patterns:" section (8 bullet points)
- **Recommendation**: Keep technical constraints, remove usage patterns

**4. ExcelDataModelTool.cs**
- Has "LLM Usage Patterns:" section
- Also has reference to LLM patterns in error messages (line 79)
- **Recommendation**: Clean up both class summary and error messages

### Medium Priority

**5. ExcelTableTool.cs**
- Has "LLM Usage Patterns:" section
- **Recommendation**: Simplify to technical description

**6. ExcelRangeTool.cs**
- Has "LLM Usage Patterns:" section
- **Recommendation**: Simplify to technical description

**7. ExcelQueryTableTool.cs**
- Has "LLM Usage Patterns:" section
- **Recommendation**: Simplify to technical description

**8. ExcelNamedRangeTool.cs**
- Has "LLM Usage Patterns:" section
- **Recommendation**: Simplify to technical description

### Low Priority (Minimal LLM Guidance)

**9. ExcelFileTool.cs**
- **Class summary** (lines 10-17): ~7 lines with "LLM Usage Pattern:"
- **Recommendation**: Simplify to 2 lines technical description

### Already Clean (No Action Needed)

**10. BatchSessionTool.cs**
- ✅ Clean technical documentation
- Only mentions "Allows LLMs to control lifecycle" (acceptable architectural note)

**11. ExcelConnectionTool.cs**
- ✅ Clean technical documentation
- No LLM-specific guidance

**12. ExcelPivotTableTool.cs**
- ⏳ Need to check (just renamed)

## Refactoring Examples

### BEFORE (ExcelPowerQueryTool.cs)
```csharp
/// <summary>
/// Excel Power Query management tool for MCP server.
/// Handles M code operations, query management, and data loading configurations.
///
/// LLM Usage Patterns:
/// - Use "list" to see all Power Queries in a workbook
/// - Use "view" to examine M code for a specific query
/// - Use "create" to add new queries from .pq files (atomic: import + load in one call)
/// - Use "export" to save M code to files for version control
/// - Use "update-mcode" to modify M code only (no refresh)
/// - Use "update-and-refresh" to update M code and refresh data atomically
/// - Use "refresh" to refresh query data from source
/// - Use "unload" to convert query to connection-only (inverse of load-to)
/// - Use "refresh-all" to refresh all queries in workbook
/// - Use "delete" to remove queries
/// - Use "get-load-config" to check current loading configuration
///
/// ATOMIC OPERATIONS:
/// - create: Import + load in one atomic operation (replaces import + load-to)
/// - update-mcode: Update M code without refresh (for staging changes)
/// - update-and-refresh: Update M code + refresh in one atomic operation
/// - unload: Convert to connection-only (inverse of load-to)
/// - refresh-all: Refresh all queries in workbook
///
/// IMPORTANT FOR DATA MODEL WORKFLOWS:
/// - "create" with loadDestination='data-model' loads to Power Pivot Data Model (ready for DAX measures)
/// - "create" with loadDestination='worksheet' loads to worksheet (users see formatted table)
/// - "create" with loadDestination='both' loads to BOTH worksheet AND Power Pivot
/// - For Power Pivot operations beyond loading data (DAX measures, relationships), use excel_datamodel or excel_powerpivot tools
///
/// VALIDATION AND EXECUTION:
/// - Create DEFAULT behavior: Automatically loads to worksheet (validates M code by executing it)
/// - Validation = Execution: Power Query M code is only validated when data is actually loaded/refreshed
/// - Connection-only queries are NOT validated until first execution
/// </summary>
```

### AFTER (Proposed)
```csharp
/// <summary>
/// Excel Power Query management tool for MCP server.
/// Manages Power Query M code operations, query lifecycle, and data loading configurations.
/// Supports loading to worksheet, Power Pivot Data Model, or both destinations.
/// </summary>
```

### BEFORE (ExcelVbaTool.cs)
```csharp
/// <summary>
/// Excel VBA management tool for MCP server.
/// Handles VBA macro operations, code management, and script execution in macro-enabled workbooks.
///
/// ⚠️ IMPORTANT: Requires .xlsm files! Standard .xlsx files don't support VBA macros.
///
/// Setup Required: Run setup-vba-trust command once to enable VBA automation.
///
/// LLM Usage Patterns:
/// - Use "list" to see all VBA modules and procedures in a workbook
/// - Use "view" to examine VBA code for a specific module without exporting
/// - Use "export" to save VBA modules to .vba files for version control
/// - Use "import" to add VBA modules from files
/// - Use "update" to modify existing VBA modules
/// - Use "run" to execute VBA procedures with optional parameters
/// - Use "delete" to remove VBA modules from workbooks
/// - Check VBA trust status with trust-status action if automation fails
/// </summary>
```

### AFTER (Proposed)
```csharp
/// <summary>
/// Excel VBA management tool for MCP server.
/// Manages VBA macro operations, code import/export, and script execution in macro-enabled workbooks.
///
/// ⚠️ IMPORTANT: Requires .xlsm files! Standard .xlsx files don't support VBA macros.
///
/// Prerequisites: VBA trust must be enabled for automation. Use setup-vba-trust command to configure.
/// </summary>
```

## Implementation Plan

### Step 1: Commit Staged Changes First
Commit the audit script fix and file renames before starting documentation cleanup:
```bash
git commit -m "refactor: Fix audit false positives + rename tool files for consistency"
```

### Step 2: Refactor Tool Documentation (One File at a Time)
For each tool file:
1. Read current class summary
2. Identify what to keep (technical constraints, prerequisites, architectural notes)
3. Identify what to remove (LLM usage patterns, workflow guidance)
4. Simplify to 2-4 lines of technical description
5. Check method `[Description]` attributes (simplify if needed)
6. Verify file still compiles
7. Stage and review changes

**Order** (highest impact first):
1. ExcelPowerQueryTool.cs (~60 lines → ~4 lines)
2. ExcelWorksheetTool.cs (~20 lines → ~4 lines)
3. ExcelVbaTool.cs (~17 lines → ~5 lines)
4. ExcelDataModelTool.cs
5. ExcelTableTool.cs
6. ExcelRangeTool.cs
7. ExcelQueryTableTool.cs
8. ExcelNamedRangeTool.cs
9. ExcelFileTool.cs
10. ExcelPivotTableTool.cs (if needed)

### Step 3: Verify Consistency
After refactoring all files:
- All tool files should have concise 2-4 line summaries
- Technical constraints and prerequisites preserved
- No "LLM Usage Patterns" sections remaining
- Build passes with 0 warnings
- All LLM guidance is in prompt files (already exists)

### Step 4: Commit Documentation Changes
```bash
git add src/ExcelMcp.McpServer/Tools/*.cs
git commit -m "docs: Remove LLM guidance from tool code documentation

Cleaned up tool file documentation to separate technical implementation
details from LLM behavioral guidance:

- Removed 'LLM Usage Patterns' sections from class summaries (10 files)
- Simplified class summaries to 2-4 lines technical description
- Preserved technical constraints, prerequisites, and architectural notes
- All LLM guidance remains in prompt files (Prompts/Content/*.md)

Files refactored: ExcelPowerQueryTool, ExcelWorksheetTool, ExcelVbaTool,
ExcelDataModelTool, ExcelTableTool, ExcelRangeTool, ExcelQueryTableTool,
ExcelNamedRangeTool, ExcelFileTool, ExcelPivotTableTool

Guidance: LLM usage patterns belong in MCP prompts, not C# documentation
(see .github/instructions/mcp-llm-guidance.instructions.md)"
```

## Why This Matters

**For C# Developers**:
- Code documentation should describe the API for developers
- LLM-specific guidance is noise for developers reading the code

**For LLMs**:
- Prompt files are the authoritative source for LLM behavior
- C# XML comments are NOT sent to LLMs via MCP protocol
- Duplicating guidance creates maintenance burden (update in 2 places)

**For Project Maintainability**:
- Single source of truth for LLM guidance (prompt files)
- Clean separation of concerns (C# docs vs. LLM prompts)
- Easier to update LLM behavior without touching C# code

## Next Steps

1. **User Decision**: Commit staged changes (audit fix + renames) now?
2. **Start Refactoring**: Begin with ExcelPowerQueryTool.cs (biggest impact)
3. **Systematic Cleanup**: Process all 10 files one by one
4. **Final Verification**: Build passes, consistency check, commit
