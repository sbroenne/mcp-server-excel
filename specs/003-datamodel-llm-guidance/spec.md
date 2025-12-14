# Feature Specification: Data Model LLM Guidance Enhancement

**Feature Branch**: `003-datamodel-llm-guidance`  
**Created**: December 14, 2025  
**Status**: Draft

## Problem

LLMs get stuck when Data Model operations fail. Their recovery instinct is to "start fresh" by deleting and recreating tables. **But deleting a table cascades to delete ALL its measures** - potentially hours of user work lost.

Secondary issue: LLMs "forget" the MCP tools exist in long conversations and suggest manual Power Pivot steps instead.

## Solution

Two simple changes:

1. **Update `excel_datamodel` tool description** - Add clear warnings about destructive operations and a quick recovery guide
2. **Create `excel_datamodel.md` prompt file** - Troubleshooting reference LLMs can consult when stuck

## Requirements

- **FR-001**: Tool description warns that delete-table also deletes all associated measures
- **FR-002**: Tool description includes quick recovery tips (use update-measure, check list-measures first)
- **FR-003**: Create prompt file with common error scenarios and non-destructive fixes

## Success Criteria

- **SC-001**: Tool description contains "DESTRUCTIVE" warning for delete-table
- **SC-002**: Prompt file exists with at least 5 common error recovery patterns
- **SC-003**: No delete-table suggestion in any recovery guidance

## Out of Scope

- Changing Excel COM behavior (cascade delete is baked in)
- Blocking delete operations (users may legitimately need them)
- Complex error response modifications (keep it simple)
