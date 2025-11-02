---
applyTo: "src/ExcelMcp.McpServer/Prompts/**/*.md"
---

# MCP LLM Prompt Creation Guide

> **How to create effective guidance for LLMs consuming the MCP server**

## Core Principle

**You are writing for LLMs (GitHub Copilot, Claude) that:**
1. Already know Excel (ranges, formulas, 2D arrays, cell references)
2. Already receive tool schema from MCP (parameter names, types, descriptions)
3. Need to know how THIS MCP SERVER works, not how Excel works

## What LLMs Already Know (DON'T INCLUDE)

**Excel Domain Knowledge:**
- ✅ What ranges, cells, formulas, and worksheets are
- ✅ How 2D arrays represent tabular data
- ✅ That formulas start with `=` and use relative/absolute references
- ✅ Common Excel functions (SUM, VLOOKUP, IF, etc.)
- ✅ Number formatting, cell styling, data validation concepts

**Programming Concepts:**
- ✅ JSON syntax and data structures
- ✅ Array indexing, null values, type conversions
- ✅ Error handling, async operations, batch processing

**MCP Protocol:**
- ✅ How to call MCP tools (syntax from schema)
- ✅ Parameter structure (schema provides this)
- ✅ Required vs optional parameters (schema provides this)

## What LLMs Need to Know (MUST INCLUDE)

**1. Action Catalog:**
- Complete list of valid action values
- Example: "Actions: get-values, set-values, clear-all, clear-contents, clear-formats"

**2. Action Disambiguation:**
- When to use each action
- Differences between similar actions
- Example: "clear-all removes formatting, clear-contents preserves it"

**3. Tool Selection:**
- When to use this tool vs other tools
- Example: "Use excel_range for data, excel_worksheet for lifecycle"

**4. Server-Specific Behavior:**
- Quirks of THIS implementation
- Example: "Single cell returns [[value]] (2D array), not scalar"
- Example: "For named ranges, use sheetName='' (empty string)"

**5. Common Mistakes:**
- Pitfalls specific to this server
- Example: "Don't forget batch mode for multiple operations"

**6. Parameter Value Examples:**
- Actual values for string parameters
- Example: rangeAddress can be "A1", "A1:C10", or "SalesData"

## Prompt File Structure

```markdown
## [Tool Name] Tool

**Actions**: [comma-separated list of all action values]

**When to use [tool_name]**:
- [Scenario 1]
- Use [other_tool] for [different scenario]

**Server-specific behavior**:
- [Quirk 1]
- [Quirk 2]

**Action guide**:
- [action-name]: [What makes this action different from similar ones]
- [action-name]: [When to choose this over alternatives]

**Common mistakes**:
- [Mistake 1 specific to this server]
```

## Length Guidelines

**Keep it SHORT:**
- ✅ One markdown file per tool (all actions in one file)
- ✅ 50-150 lines total per tool
- ✅ Focus on disambiguation, not explanation
- ❌ Don't write tutorials about Excel concepts
- ❌ Don't explain what LLMs already know

## Examples

### ✅ GOOD - Server-specific, concise

```markdown
## excel_range Tool

**Actions**: get-values, set-values, clear-all, clear-contents, clear-formats

**When to use excel_range**:
- Data operations (read/write)
- Use excel_worksheet for sheet lifecycle
- Use excel_namedrange for range definitions

**Server behavior**:
- Single cell returns [[value]] (2D array)
- Named ranges: sheetName=""
- Batch mode recommended for multiple ops

**Action disambiguation**:
- clear-all: Content + formatting
- clear-contents: Content only
- clear-formats: Formatting only
```

### ❌ BAD - Teaching Excel to LLMs

```markdown
## What are Excel Ranges?

An Excel range is a group of cells. For example, A1:C10 represents
10 rows and 3 columns. Ranges are fundamental to Excel...

[200 lines of Excel tutorial]

**How to read values:**
```json
{
  "action": "get-values",
  "rangeAddress": "A1:C10"
}
```
```

**Why bad:**
- Explains Excel concepts LLMs already know
- Includes JSON syntax (schema already provides)
- Way too long for simple action catalog

## Format Guidelines

**Use markdown files (.md):**
- Store in `src/ExcelMcp.McpServer/Prompts/Content/`
- One file per tool
- Plain markdown (no C# code)

**Writing style:**
- Bullet points over paragraphs
- Action-oriented ("Use X for Y")
- Comparative ("X vs Y: choose X when...")
- Example values in quotes ("A1", "SalesData")

**What to emphasize:**
- ⭐ Action catalog (most important)
- ⭐ Tool selection (when to use this vs others)
- ⭐ Server quirks (non-obvious behavior)
- ⚠️ Common mistakes (server-specific pitfalls)

**What to minimize:**
- Domain knowledge (LLMs know Excel)
- Syntax examples (schema provides)
- Long explanations (keep concise)

## Testing Your Prompts

**Ask yourself as an LLM:**
1. Do I know which action values are valid? ✅ Must include
2. Do I know when to use action A vs action B? ✅ Must include
3. Do I know when to use this tool vs another tool? ✅ Must include
4. Am I learning Excel concepts? ❌ Remove this
5. Am I seeing JSON syntax examples? ❌ Remove this
6. Is this longer than 150 lines? ❌ Make it shorter

## Anti-Patterns to Avoid

**❌ The Tutorial:**
```markdown
# Excel Ranges Explained
Ranges are groups of cells. They use A1 notation...
[300 lines teaching Excel]
```

**❌ The Syntax Guide:**
```markdown
# How to Call excel_range
```json
{
  "action": "get-values",
  "excelPath": "file.xlsx"
}
```
[100 lines of JSON examples]
```

**❌ The Encyclopedia:**
```markdown
# Complete Reference
Every possible parameter combination...
[500 lines of exhaustive documentation]
```

## Success Criteria

A good prompt:
- ✅ Lists all valid action values
- ✅ Disambiguates similar actions
- ✅ Explains server-specific quirks
- ✅ Helps choose between tools
- ✅ Under 150 lines
- ✅ Pure markdown in Content/ directory
- ❌ Doesn't teach Excel concepts
- ❌ Doesn't show JSON syntax
- ❌ Doesn't duplicate schema info

## Completions (Autocomplete) - NOT YET IMPLEMENTED

**Purpose**: Provide autocomplete suggestions for prompt arguments and resource URIs

**FUTURE**: Completions will be stored as `.md` files in `Content/Completions/` directory

**Current State**: 
- Completions currently in `ExcelCompletionHandler.cs` (C# code)
- **TODO**: Migrate to `.md` files for easier maintenance
- **TODO**: Create loader to read completion markdown files

**What completions do**:
- Suggest valid action values when user types `action=`
- Suggest format codes when user types `formatString=`
- Suggest file paths when user types Excel file URIs
- Suggest common parameter values (privacy levels, alignment, colors, etc.)

**Future .md file structure**:
```markdown
# Completions for [parameter-name]

value1
value2
value3
```

**Completion Guidelines**:
- ✅ Include only valid values (not examples)
- ✅ Most common values first
- ✅ Limit to 10-15 suggestions per parameter
- ✅ Use lowercase for consistency (except Excel-specific like "Private")
- ❌ Don't include every possible value (overwhelming)
- ❌ Don't duplicate MCP schema info

**When to add completions**:
- New parameter with enum-like values (actions, types, modes)
- Common string patterns (format codes, colors, alignments)
- File/path parameters (suggest existing files)

## Elicitations (Pre-flight Checklists)
**Purpose**: Guide users to provide ALL needed information before calling tools (prevents back-and-forth)

**FUTURE**: Elicitations will be stored as `.md` files in `Content/Elicitations/` directory

**Current State**:
- Elicitations currently in `ExcelElicitationPrompts.cs` (C# code with embedded strings)
- **TODO**: Migrate to `.md` files for easier maintenance
- **TODO**: Create loader to read elicitation markdown files

**What elicitations do**:
- Checklist of REQUIRED information
- Checklist of RECOMMENDED information (avoid second call)
- Workflow optimization hints (batch mode detection)
- Ask user for missing info BEFORE tool invocation

**Future .md file structure**:
```markdown
# BEFORE [OPERATION] - GATHER THIS INFO

REQUIRED:
☐ Parameter 1 (description)
☐ Parameter 2 (description)

RECOMMENDED (avoid second call):
☐ Optional param that improves workflow
☐ Common follow-up parameter

WORKFLOW OPTIMIZATION:
☐ Batch mode? (detect keywords: numbers, plurals, lists)
☐ Prerequisites? (check dependencies first)

ASK USER FOR MISSING INFO before calling [tool_name].
```

**When to create elicitations**:
- Complex operations with many optional parameters
- Operations that commonly require follow-up calls
- Multi-step workflows (import → configure → refresh)
- Batch-friendly operations (detect plural requests)

**Elicitation Guidelines**:
- ✅ Start with REQUIRED (must-have info)
- ✅ Add RECOMMENDED (nice-to-have, avoids round-trips)
- ✅ Include workflow hints (batch mode detection)
- ✅ Keep checklist format (☐ bullets)
- ✅ Explain WHY each parameter matters
- ❌ Don't duplicate tool schema (LLMs already have it)
- ❌ Don't include obvious parameters (excelPath always needed)

**Example - Detecting batch opportunities**:
- "import 5 queries" → use begin_excel_batch
- "create measures for Sales, Customers, Orders" → batch mode
- "add Total Sales, Avg Price, Customer Count" → batch mode

## Workflow Guidance (SuggestedNextActions & WorkflowHint) - C# IMPLEMENTATION

**Purpose**: Guide LLM workflow after each operation (next logical steps)

**IMPLEMENTATION: C# Static Methods** (NOT .md files)

**Why C# instead of .md files:**
1. **Runtime Context Required**: Workflow guidance depends on runtime state (success/failure, batch mode, operation count, error types)
2. **Conditional Logic**: Different messages for different scenarios - requires if/else branching
3. **Already Reusable**: Static methods shared between CLI and MCP Server
4. **Type Safety**: Parameters are strongly typed (bool success, int count)
5. **Different from Prompts**: Prompts are static content read once; workflow guidance is dynamic per-operation

**Implementation**:
- Location: `src/ExcelMcp.McpServer/Tools/*Tool.cs`
- Pattern: Ad-hoc JSON properties in tool responses
- Example:
     ```csharp
     return JsonSerializer.Serialize(new
     {
         success = true,
         workflowHint = "File is ready for Excel operations.",
         suggestedNextActions = new[]
         {
             "Use excel_worksheet to manage worksheets",
             "Use excel_powerquery to manage Power Query connections"
         }
     }, JsonOptions);
     ```

**Characteristics**:
- **Success scenarios**: "Use begin_excel_batch", "Use worksheet 'create'", "Verify results"
- **Failure scenarios**: "Check directory exists", "Try different path", "Review error messages"
- **Contextual**: Batch mode detection ("Creating multiple? Use batch mode")
- **Workflow chains**: Import → Configure → Refresh → Verify

**When to Add Workflow Guidance:**
- After CREATE operations: Suggest next steps (configure, populate, verify)
- After LIST operations: Suggest actions based on count (create if empty, inspect if populated)
- After UPDATE operations: Suggest verification steps
- After FAILURE: Suggest troubleshooting steps specific to error type
- Batch mode hints: "Creating multiple? Use begin_excel_batch for better performance"

## Remember

**LLMs using your MCP server already know:**
- Excel (ranges, formulas, cells)
- JSON (syntax, structure)
- Programming (arrays, null, types)

**LLMs need to know:**
- Which actions exist → **Prompts** - .md files
- How to choose between actions → **Prompts** - .md files
- Server-specific behavior → **Prompts** - .md files
- When to use this tool vs others → **Prompts** - .md files
- Valid parameter values → **Completions** - .md files (TODO: migrate from C#)
- What info to gather first → **Elicitations** - .md files (TODO: migrate from C#)
- What to do next → **Workflow Guidance** - C# static methods (runtime context required)

**Architecture Summary:**

| Guidance Type | Format | Why | Status |
|---------------|--------|-----|--------|
| **Prompts** | .md files | Static content, read once | ✅ Implemented |
| **Completions** | .md files | Static value lists | TODO: Migrate from C# |
| **Elicitations** | .md files | Static checklists | TODO: Migrate from C# |
| **Workflow Guidance** | C# static methods | Dynamic, runtime context | ✅ Keep as C# |

**Keep it short. Keep it specific. Keep it server-focused.**
