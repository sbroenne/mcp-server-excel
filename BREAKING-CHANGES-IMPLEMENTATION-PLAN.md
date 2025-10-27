# Breaking Changes Implementation Plan

**Date**: October 27, 2025
**Status**: Ready for Implementation
**Estimated Effort**: 5-7 days (per MCP-BREAKING-CHANGES-PROPOSAL.md)

---

## Scope Summary

Per user request, implementing all breaking changes from MCP-BREAKING-CHANGES-PROPOSAL.md plus investigating structured tool output from Microsoft blog article.

**Files Affected**:
- 17 C# files with `batchId` references
- 16 C# files with `excelPath` references  
- 6 markdown documentation files
- All tool files for parameter standardization
- All Core command files
- All test files

**Total Estimated Changes**: ~30-40 files

---

## Implementation Phases

### Phase 1: Critical Renaming (Day 1-2)

#### 1.1 batchId → sessionId
**Files to update**:
- `src/ExcelMcp.McpServer/Tools/BatchSessionTool.cs` → Rename to `SessionTool.cs`
- All tools with `batchId` parameter (9 tool files)
- `src/ExcelMcp.Core/Session/ExcelBatch.cs` (parameter names)
- All test files with batch references
- Documentation: BATCH-SESSION-GUIDE.md → SESSION-GUIDE.md

**Tool Renames**:
- `begin_excel_batch` → `begin_excel_session`
- `commit_excel_batch` → `end_excel_session` 
- `list_excel_batches` → `list_excel_sessions`

#### 1.2 excelPath → filePath
**Files to update**:
- All 9 tool files in `src/ExcelMcp.McpServer/Tools/`
- All Core command interfaces
- All Core command implementations
- All test files

#### 1.3 sheetName → worksheetName
**Files to update**:
- `ExcelWorksheetTool.cs`
- Worksheet command files
- Related tests

### Phase 2: Error Response Standardization (Day 3)

#### 2.1 Define Error Codes
Create `src/ExcelMcp.Core/Models/ErrorCodes.cs`:
```csharp
public static class ErrorCodes
{
    public const string FILE_NOT_FOUND = "FILE_NOT_FOUND";
    public const string QUERY_NOT_FOUND = "QUERY_NOT_FOUND";
    public const string WORKSHEET_NOT_FOUND = "WORKSHEET_NOT_FOUND";
    public const string INVALID_M_CODE = "INVALID_M_CODE";
    public const string PRIVACY_LEVEL_REQUIRED = "PRIVACY_LEVEL_REQUIRED";
    public const string VBA_TRUST_REQUIRED = "VBA_TRUST_REQUIRED";
    public const string EXCEL_BUSY = "EXCEL_BUSY";
    public const string SESSION_NOT_FOUND = "SESSION_NOT_FOUND";
    public const string SESSION_FILE_MISMATCH = "SESSION_FILE_MISMATCH";
}
```

#### 2.2 Standardize Error Response Format
Update all Core commands to return:
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

### Phase 3: Cleanup & Enhancement (Day 4-5)

#### 3.1 Remove Redundant Validation Attributes
- Remove `[RegularExpression]` from MCP tool parameters
- Move validation to Core layer
- Keep only `[Description]` attributes in MCP layer

#### 3.2 Add Rich Metadata to Responses
Enhance all tool responses with:
- Timestamp
- Version info
- Related operations suggestions
- Performance metrics (when relevant)

### Phase 4: Structured Tool Output Investigation (Day 5-6)

#### 4.1 Research MCP C# SDK Support
Investigate if current SDK supports:
- Returning typed objects instead of JSON strings
- Schema validation
- Automatic serialization

#### 4.2 Implementation (if SDK supports)
If SDK supports typed responses:
- Define result types for all operations
- Update all tools to return typed objects
- Verify MCP protocol compliance

If SDK doesn't support:
- Document limitation
- Keep current JSON string approach

### Phase 5: Documentation & Testing (Day 6-7)

#### 5.1 Update All Documentation
- README.md (all parameter names, tool names)
- BATCH-SESSION-GUIDE.md → SESSION-GUIDE.md
- All prompt content (update terminology)
- MCP-IMPLEMENTATION-SUMMARY.md
- Architecture diagrams

#### 5.2 Update All Tests
- Rename test methods
- Update test data
- Fix parameter names
- Verify all tests pass

#### 5.3 Final Validation
- Build solution (0 warnings, 0 errors)
- Run all tests
- Test MCP server startup
- Manual testing with VS Code

---

## Risk Assessment

**HIGH RISK**:
- Breaking all existing integrations (but OK - pre-1.0)
- Test suite may have failures requiring fixes
- Documentation may be incomplete in some areas

**MITIGATION**:
- Systematic approach with git commits after each phase
- Comprehensive testing after each major change
- Can revert individual phases if needed

---

## Decision Points

Before starting, confirm:
1. **Scope**: Should this be done in THIS PR or separate PR?
2. **Timeline**: Is 5-7 day timeline acceptable?
3. **Testing**: Can we defer some test fixes to follow-up?
4. **Structured Output**: Proceed even if SDK doesn't support yet?

---

## Current Status

✅ Plan created
⏳ Awaiting confirmation to proceed with implementation
⏳ Starting with Phase 1.1 (batchId → sessionId) as highest priority

---

**Note**: This is a MASSIVE change affecting the entire codebase. Recommend careful review at each phase checkpoint.
