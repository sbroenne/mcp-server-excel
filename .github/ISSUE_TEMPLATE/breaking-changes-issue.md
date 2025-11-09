# Breaking Changes for Pre-1.0 Release

## Overview

Implement breaking changes from `MCP-BREAKING-CHANGES-PROPOSAL.md` before the 1.0 release. These changes will improve API clarity, consistency, and developer experience but require updates across ~30-40 files.

**Estimated Effort**: 5-7 days  
**Status**: Scoped and documented, ready for implementation  
**Related PR**: #[PR_NUMBER] (prompts + completions)

---

## Objectives

Since ExcelMcp MCP Server hasn't been released yet, we can make breaking changes without affecting users. This is a **golden opportunity** to improve the API before 1.0.

### Key Changes

1. **Better Terminology**: `batchId` → `sessionId` (clearer intent)
2. **Consistent Naming**: `excelPath` → `filePath`, `sheetName` → `worksheetName`
3. **Standardized Errors**: Error codes and structured error responses
4. **Cleaner Code**: Remove redundant validation attributes
5. **Richer Responses**: Add metadata to all tool outputs
6. **Modern SDK**: Investigate structured tool output support

---

## Detailed Scope

### Phase 1: Critical Renaming (2 days)

#### 1.1 batchId → sessionId
**Affected files**: 17 C# files, 6 markdown files

- [ ] Rename `BatchSessionTool.cs` → `SessionTool.cs`
- [ ] Rename tools:
  - `begin_excel_batch` → `begin_excel_session`
  - `commit_excel_batch` → `end_excel_session`
  - `list_excel_batches` → `list_excel_sessions`
- [ ] Update all `batchId` parameters to `sessionId` in:
  - All 9 tool files in `src/ExcelMcp.McpServer/Tools/`
  - `ExcelToolsBase.cs`
  - All prompt files (4 files)
- [ ] Update documentation:
  - `BATCH-SESSION-GUIDE.md` → `SESSION-GUIDE.md`
  - Update all prompt content
  - Update README.md
  - Update MCP-IMPLEMENTATION-SUMMARY.md
- [ ] Update tests (all files referencing batchId)
- [ ] Update Program.cs cleanup handler

#### 1.2 excelPath → filePath
**Affected files**: 16 C# files

- [ ] Update all tool files:
  - `ExcelPowerQueryTool.cs`
  - `ExcelWorksheetTool.cs`
  - `ExcelParameterTool.cs`
  - `ExcelCellTool.cs`
  - `ExcelVbaTool.cs`
  - `ExcelConnectionTool.cs`
  - `ExcelDataModelTool.cs`
  - `ExcelFileTool.cs`
  - `HyperlinkTool.cs`
  - `TableTool.cs`
- [ ] Update all Core command interfaces
- [ ] Update all Core command implementations
- [ ] Update all tests
- [ ] Update all prompt content and documentation

#### 1.3 sheetName → worksheetName
**Affected files**: ~5 files

- [ ] `ExcelWorksheetTool.cs`
- [ ] Worksheet Core commands
- [ ] Related tests
- [ ] Prompt content
- [ ] Documentation

### Phase 2: Error Response Standardization (1-2 days)

#### 2.1 Define Error Codes
- [ ] Create `src/ExcelMcp.Core/Models/ErrorCodes.cs`
- [ ] Define standard error codes:
  ```csharp
  FILE_NOT_FOUND
  QUERY_NOT_FOUND
  WORKSHEET_NOT_FOUND
  INVALID_M_CODE
  PRIVACY_LEVEL_REQUIRED
  VBA_TRUST_REQUIRED
  EXCEL_BUSY
  SESSION_NOT_FOUND
  SESSION_FILE_MISMATCH
  ```

#### 2.2 Standardize Error Format
- [ ] Update all Core commands to return structured errors:
  ```json
  {
    "success": false,
    "error": {
      "code": "QUERY_NOT_FOUND",
      "message": "Power Query 'SalesData' not found",
      "details": {
        "queryName": "SalesData",
        "availableQueries": ["Data1", "Data2"]
      }
    }
  }
  ```
- [ ] Update all MCP tools to handle new error format
- [ ] Update tests for new error format
- [ ] Document error codes in README.md

### Phase 3: Cleanup & Enhancement (2-3 days)

#### 3.1 Remove Redundant Validation Attributes
- [ ] Remove `[RegularExpression]` from MCP tool parameters
- [ ] Remove `[StringLength]` from MCP tools
- [ ] Remove `[FileExtensions]` from MCP tools
- [ ] Keep only `[Description]` attributes in MCP layer
- [ ] Ensure all validation happens in Core layer
- [ ] Update tests

#### 3.2 Add Rich Metadata to Responses
- [ ] Add to all tool responses:
  - Timestamp
  - Server version
  - Operation duration (when relevant)
  - Related operations/suggestions
  - Performance metrics
- [ ] Define standard metadata structure
- [ ] Update all tools to include metadata
- [ ] Update tests

### Phase 4: Structured Tool Output (1-2 days)

#### 4.1 Research MCP C# SDK Support
- [ ] Review latest MCP C# SDK documentation
- [ ] Check if SDK supports typed responses (not JSON strings)
- [ ] Review [Microsoft's blog article](https://devblogs.microsoft.com/dotnet/mcp-csharp-sdk-2025-06-18-update/)
- [ ] Test with sample typed response

#### 4.2 Implementation (if supported)
- [ ] Define result types for all operations
- [ ] Update all tools to return typed objects
- [ ] Verify MCP protocol compliance
- [ ] Update tests
- [ ] Document approach

#### 4.3 Documentation (if not supported)
- [ ] Document SDK limitation
- [ ] Keep current JSON string approach
- [ ] Note for future when SDK adds support

### Phase 5: Final Integration (1 day)

#### 5.1 Update All Documentation
- [ ] README.md (all examples with new names)
- [ ] SESSION-GUIDE.md (renamed from BATCH-SESSION-GUIDE.md)
- [ ] All prompt content files
- [ ] MCP-IMPLEMENTATION-SUMMARY.md
- [ ] Architecture diagrams (if any)

#### 5.2 Testing
- [ ] Build solution (0 warnings, 0 errors)
- [ ] Run all unit tests
- [ ] Run all integration tests
- [ ] Test MCP server startup
- [ ] Manual testing with VS Code
- [ ] Verify all tools work with new parameter names
- [ ] Verify error responses are correct

#### 5.3 Migration Guide
- [ ] Create MIGRATION-GUIDE.md for internal reference
- [ ] Document all changes
- [ ] Provide before/after examples
- [ ] Note that this is pre-1.0 (no external migration needed)

---

## Files Affected

**C# Files**: ~30 files
- 9 tool files in `src/ExcelMcp.McpServer/Tools/`
- 4 prompt files in `src/ExcelMcp.McpServer/Prompts/`
- 1 Program.cs
- ~10 Core command files
- ~10 test files

**Documentation**: ~10 files
- README.md
- BATCH-SESSION-GUIDE.md → SESSION-GUIDE.md
- MCP-IMPLEMENTATION-SUMMARY.md
- New: ErrorCodes.md
- New: MIGRATION-GUIDE.md

---

## Success Criteria

- [ ] All `batchId` references changed to `sessionId`
- [ ] All `excelPath` references changed to `filePath`
- [ ] All `sheetName` references changed to `worksheetName`
- [ ] Error response format standardized with error codes
- [ ] Validation attributes cleaned up
- [ ] Rich metadata added to all responses
- [ ] Structured output investigated and implemented (if SDK supports)
- [ ] All tests passing
- [ ] All documentation updated
- [ ] Build successful (0 warnings, 0 errors)
- [ ] Manual testing completed
- [ ] MCP server starts and works correctly

---

## Implementation Strategy

### Recommended Approach

1. **Create feature branch**: `feature/breaking-changes-pre-1.0`
2. **Implement phase by phase**: Commit after each phase
3. **Test incrementally**: Run tests after each major change
4. **Review at checkpoints**: Pause after each phase for review
5. **Merge before 1.0**: Ensure this is in before first NuGet publish

### Git Workflow

```bash
git checkout -b feature/breaking-changes-pre-1.0
# Implement Phase 1.1
git commit -m "Phase 1.1: Rename batchId to sessionId"
# Test
# Implement Phase 1.2
git commit -m "Phase 1.2: Rename excelPath to filePath"
# Continue...
```

---

## References

- **Proposal Document**: `MCP-BREAKING-CHANGES-PROPOSAL.md`
- **Implementation Plan**: `BREAKING-CHANGES-IMPLEMENTATION-PLAN.md`
- **Status Document**: `BREAKING-CHANGES-STATUS.md`
- **Microsoft Blog**: https://devblogs.microsoft.com/dotnet/mcp-csharp-sdk-2025-06-18-update/
- **Related PR**: Prompts + Completions implementation

---

## Risk Assessment

**LOW RISK** - This is pre-1.0:
- No external users to break
- Can iterate freely
- Can revert if needed

**MEDIUM EFFORT** - Large scope:
- 30-40 files affected
- 5-7 days estimated
- Requires careful testing

**HIGH VALUE** - Long-term benefit:
- Clearer, more consistent API
- Better developer experience
- Less technical debt
- Professional API design

---

## Questions / Discussion

- Should we do all phases at once or incrementally?
- Any other breaking changes to include while we're at it?
- Priority order of phases - is batch→session most critical?
- Should structured output investigation happen first or last?

---

**Label**: `enhancement`, `breaking-change`, `pre-release`  
**Milestone**: `1.0.0`  
**Assignee**: TBD  
**Estimated time**: 5-7 days
