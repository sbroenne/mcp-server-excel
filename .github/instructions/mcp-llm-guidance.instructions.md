---
applyTo: "src/ExcelMcp.McpServer/Prompts/**/*.md"
---

# MCP LLM Guidance Creation Guide

> **How to create effective guidance for LLMs consuming the MCP server**

## Core Principle

**Write FOR expert LLMs (GitHub Copilot, Claude), not ABOUT the system.**

LLMs already know Excel, JSON, and MCP protocol. They need server-specific patterns only.

## What to Include

**1. Action Catalog:**
- Complete list of valid action values
- Example: "Actions: get-values, set-values, clear-all"

**2. Action Disambiguation:**
- When to use each action
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

## What to Exclude

**DON'T explain:**
- Excel concepts (ranges, formulas, cells)
- JSON syntax
- Programming basics (arrays, null, types)
- MCP protocol syntax
- Parameter types/requirements (schema provides)

## Prompt File Structure

```markdown
## [Tool Name] Tool

**Actions**: [comma-separated list]

**When to use [tool_name]**:
- [Scenario 1]
- Use [other_tool] for [different scenario]

**Server-specific behavior**:
- [Quirk 1]
- [Quirk 2]

**Action guide**:
- [action-name]: [What makes this different]
- [action-name]: [When to choose this]

**Common mistakes**:
- [Mistake 1 specific to this server]
```

## Length Guidelines

- One markdown file per tool
- 50-150 lines total per tool
- Focus on disambiguation, not explanation
- Don't write Excel tutorials
- Don't explain JSON syntax

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
- Action catalog (most important)
- Tool selection (when to use this vs others)
- Server quirks (non-obvious behavior)
- Common mistakes (server-specific pitfalls)

## Completions (Autocomplete) - NOT IMPLEMENTED

**Status**: The MCP SDK supports completions but this feature is not currently implemented.

**Alternative**: The MCP SDK auto-generates enum values in the tool schema, so LLMs already see valid action values. For freeform parameters like format codes or color values, document suggestions in tool XML documentation instead.

## Elicitations (Pre-flight Checklists) - IMPLEMENTED

**Purpose**: Guide users to provide ALL needed information before calling tools

**Current State**:
- Elicitations stored as `.md` files in `Content/Elicitations/` directory
- Loader implemented: `MarkdownLoader.LoadElicitation()`
- Used in `ExcelElicitationPrompts.cs`

**File structure**:
```markdown
# BEFORE [OPERATION] - GATHER THIS INFO

REQUIRED:
☐ Parameter 1 (description)
☐ Parameter 2 (description)

RECOMMENDED (avoid second call):
☐ Optional param that improves workflow

WORKFLOW OPTIMIZATION:
☐ Batch mode? (detect keywords: numbers, plurals, lists)
☐ Prerequisites? (check dependencies first)

ASK USER FOR MISSING INFO before calling [tool_name].
```

**When to create elicitations**:
- Complex operations with many optional parameters
- Operations that commonly require follow-up calls
- Multi-step workflows
- Batch-friendly operations

## Workflow Guidance (SuggestedNextActions & WorkflowHint) - C# IMPLEMENTATION

**Purpose**: Guide LLM workflow after each operation

**IMPLEMENTATION: C# Static Methods** (NOT .md files)

**Why C# instead of .md:**
- Runtime context required (success/failure, batch mode, operation count)
- Conditional logic needed
- Already reusable between CLI and MCP Server

**Implementation**:
- Location: `src/ExcelMcp.McpServer/Tools/*Tool.cs`
- Pattern: Ad-hoc JSON properties in tool responses

**When to Add:**
- After CREATE operations: Suggest next steps
- After LIST operations: Suggest actions based on count
- After UPDATE operations: Suggest verification
- After FAILURE: Suggest troubleshooting
- Batch mode hints: "Creating multiple? Use begin_excel_batch"

## Success Criteria

A good prompt:
- Lists all valid action values
- Disambiguates similar actions
- Explains server-specific quirks
- Helps choose between tools
- Under 150 lines
- Doesn't teach Excel concepts
- Doesn't show JSON syntax
- Doesn't duplicate schema info

## Architecture Summary

| Guidance Type | Format | Why | Status |
|---------------|--------|-----|--------|
| **Prompts** | .md files | Static content, read once | Implemented |
| **Completions** | N/A | SDK auto-generates enum values | Not implemented |
| **Elicitations** | .md files | Static checklists | Implemented |
| **Workflow Guidance** | C# static methods | Dynamic, runtime context | Keep as C# |

**Keep it short. Keep it specific. Keep it server-focused.**
