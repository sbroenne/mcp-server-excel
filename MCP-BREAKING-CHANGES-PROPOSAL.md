# MCP Server Breaking Changes Proposal (Pre-1.0)

**Date**: October 27, 2025  
**Status**: Proposal  
**Context**: MCP Server has NOT been released yet - we can break compatibility freely before 1.0 launch

---

## Opportunity: Clean Slate Before 1.0 🎯

Since the MCP server hasn't been released to production, we have a **unique opportunity** to make architectural improvements that would be breaking changes after 1.0. Let's take advantage of this window!

**Current Version**: `1.0.0` (unreleased)  
**Target Version**: `1.0.0` (with breaking changes incorporated before first release)

---

## Recommended Breaking Changes

### 1. ✅ CONSOLIDATE: Merge `action` Parameter Pattern ⚡

**Current Problem**: Tools use `action` string parameter with validation

```csharp
[RegularExpression("^(list|view|import|export|...)$")]
string action
```

**Issues**:
- Actions are strings, prone to typos
- No compile-time safety
- Completion support requires custom logic
- Tool parameter validation is verbose

**BETTER APPROACH**: Separate tools per major operation type

#### Option A: Keep Current Pattern ✅ RECOMMENDED

**Rationale**: 
- MCP best practice is resource-based tools (not operation-based)
- Current design follows MCP spec guidance
- Each tool represents a **resource** (excel_powerquery, excel_worksheet)
- Actions are **methods** on that resource (CRUD operations)
- Easier for LLMs to discover capabilities
- Follows RESTful design principles

**Validation**: Current pattern is **already optimal** for MCP

#### Option B: Split into Many Tools ❌ NOT RECOMMENDED

```csharp
// Don't do this - creates 60+ tools
excel_powerquery_list
excel_powerquery_view
excel_powerquery_import
// ... 11 PowerQuery tools
// ... 10 Worksheet tools
// ... 11 Connection tools
// = 60+ total tools
```

**Problems**:
- Tool explosion (poor discoverability)
- Goes against MCP best practices
- Harder for LLMs to navigate
- More complex server configuration

**DECISION**: ✅ **KEEP current action-based pattern** - it's already correct!

---

### 2. ✅ SIMPLIFY: Rename `batchId` to `sessionId` 🎯

**Current**: `batchId` parameter for batch sessions

**Issue**: "Batch" implies bulk processing, but we're doing session management

**BETTER**: `sessionId` - clearer intent

```csharp
// Before
[Description("Optional batch session ID from begin_excel_batch")]
string? batchId = null

// After
[Description("Optional session ID from begin_excel_session")]
string? sessionId = null
```

**Impact**:
- Rename tools:
  - `begin_excel_batch` → `begin_excel_session`
  - `commit_excel_batch` → `commit_excel_session` or `end_excel_session`
  - `list_excel_batches` → `list_excel_sessions`
- Update all tool parameters
- Update prompts
- Update documentation

**Benefits**:
- Clearer terminology
- "Session" is standard term for stateful workflows
- Aligns with database/web terminology users know

---

### 3. ✅ STANDARDIZE: Tool Return Values - Typed Responses 📦

**Current Problem**: All tools return `Task<string>` with JSON serialization

```csharp
public static async Task<string> ExcelPowerQuery(...)
{
    return JsonSerializer.Serialize(result);
}
```

**Issues**:
- No type safety
- LLMs must parse JSON strings
- No schema validation
- Inconsistent response formats

**BETTER APPROACH**: Return strongly-typed objects (if SDK supports)

```csharp
// If MCP C# SDK supports typed responses
public static async Task<PowerQueryListResult> ListPowerQueries(...)
{
    return new PowerQueryListResult
    {
        Success = true,
        Queries = queries,
        Count = queries.Length
    };
}
```

**Research Needed**: Does MCP C# SDK support typed tool responses?

**If NO**: Keep current string/JSON approach ✅  
**If YES**: Switch to typed responses for Phase 2

---

### 4. ✅ IMPROVE: Resource URIs - Use Standard Patterns 🔗

**For Resources Implementation** (from main proposal):

**Pattern to Use**: Follow MCP URI conventions

```
excel://metadata/{encodedFilePath}
excel://query/{encodedFilePath}/{queryName}
excel://worksheet/{encodedFilePath}/{sheetName}
excel://datamodel/{encodedFilePath}
excel://vba/{encodedFilePath}/{moduleName}
```

**Important**: Encode file paths properly (URI encoding)

```csharp
var encodedPath = Uri.EscapeDataString(filePath);
var resourceUri = $"excel://query/{encodedPath}/{queryName}";
```

This prevents issues with special characters in file paths.

---

### 5. ✅ REFINE: Parameter Naming Consistency 📝

**Current Inconsistencies**:

| Tool | File Path Param | Query Name Param |
|------|----------------|------------------|
| excel_powerquery | `excelPath` | `queryName` |
| excel_worksheet | `excelPath` | `sheetName` |
| excel_connection | `excelPath` | `connectionName` |

**Issue**: `excelPath` is clear, but some tools use different conventions

**Standardization**:

```csharp
// Standard pattern for ALL tools
[Description("Excel file path (.xlsx or .xlsm)")]
string filePath,  // NOT excelPath

[Description("Power Query name")]
string queryName,  // Good

[Description("Worksheet name")]
string worksheetName,  // NOT sheetName (clearer)

[Description("Connection name")]
string connectionName,  // Good
```

**Changes**:
- `excelPath` → `filePath` (all tools)
- `sheetName` → `worksheetName` (clearer, more explicit)
- `sourcePath` / `targetPath` - keep as-is (clear context)

**Benefits**:
- Consistent parameter naming
- `filePath` is more generic (could support .xls later)
- `worksheetName` is more explicit than `sheetName`

---

### 6. ✅ ENHANCE: Error Response Format 🚨

**Current**: Tools throw `McpException` or return error in JSON

**Inconsistency**: Some return `{ success: false, error: "..." }`, others throw

**BETTER**: Standardize error handling

```csharp
// Standard error response format
{
  "success": false,
  "error": {
    "code": "QUERY_NOT_FOUND",
    "message": "Power Query 'SalesData' not found in workbook",
    "details": {
      "queryName": "SalesData",
      "availableQueries": ["CustomerData", "ProductData"]
    }
  }
}
```

**Error Codes**: Define standard error codes
- `FILE_NOT_FOUND`
- `QUERY_NOT_FOUND`
- `WORKSHEET_NOT_FOUND`
- `INVALID_M_CODE`
- `PRIVACY_LEVEL_REQUIRED`
- `VBA_TRUST_REQUIRED`
- `EXCEL_BUSY`
- `BATCH_NOT_FOUND`
- `BATCH_FILE_MISMATCH`

**Benefits**:
- LLMs can handle errors programmatically
- Consistent error format
- Better error recovery
- Actionable error messages

---

### 7. ✅ OPTIMIZE: Remove Redundant Validation Attributes 🧹

**Current**: Heavy use of validation attributes

```csharp
[Required]
[RegularExpression("^(list|view|...)$")]
[StringLength(255, MinimumLength = 1)]
[FileExtensions(Extensions = "xlsx,xlsm")]
```

**Issues**:
- Over-validation at MCP layer
- Core layer already validates
- Double validation adds overhead
- Attributes clutter code

**BETTER**: Minimal MCP validation, comprehensive Core validation

```csharp
// MCP layer - minimal validation (required params only)
[Description("Action to perform")]
string action,

[Description("Excel file path")]
string filePath,

// Core layer - comprehensive validation (in Core commands)
public async Task<OperationResult> Import(string filePath, string queryName, ...)
{
    // Validate parameters
    filePath = PathValidator.ValidateAndNormalizePath(filePath);
    
    if (!filePath.EndsWith(".xlsx") && !filePath.EndsWith(".xlsm"))
        return OperationResult.Error("FILE_INVALID_EXTENSION", ...);
        
    // ... more validation
}
```

**Benefits**:
- Cleaner MCP tool definitions
- Single source of truth for validation (Core)
- Easier to maintain
- Better error messages from Core

---

### 8. ✅ CLARIFY: Privacy Level Parameter 🔒

**Current**: `privacyLevel` is optional string with regex validation

**Issue**: LLMs don't know when it's required

**BETTER**: Return `PrivacyLevelRequired` result type when needed

```csharp
// Import returns specific result type
public class PowerQueryPrivacyErrorResult : ResultBase
{
    public bool PrivacyLevelRequired { get; set; } = true;
    public string[] DetectedPrivacyLevels { get; set; }
    public string RecommendedPrivacyLevel { get; set; }
    public string Explanation { get; set; }
    public string[] Options { get; set; } // ["None", "Private", "Organizational", "Public"]
}
```

**LLM Flow**:
1. Call import without privacyLevel
2. Get `PrivacyLevelRequired` response with options
3. Ask user to choose
4. Call import again with user's choice

**Benefits**:
- Explicit privacy level requirement
- Educational for users
- Security-first design
- No ambiguity

---

### 9. ✅ ADD: Metadata to Tool Responses 📊

**Current**: Tools return minimal info

```json
{
  "success": true,
  "message": "Query imported"
}
```

**BETTER**: Include rich metadata

```json
{
  "success": true,
  "data": {
    "queryName": "SalesData",
    "sourceFile": "sales.pq",
    "loadedToWorksheet": "Sheet1",
    "rowsLoaded": 1000,
    "columnsLoaded": 8
  },
  "metadata": {
    "operation": "import",
    "duration": "2.3s",
    "excelVersion": "16.0",
    "privacyLevel": "Private"
  },
  "suggestions": [
    "Refresh the query to load latest data",
    "Use worksheet 'read' to view imported data",
    "Use powerquery 'view' to inspect M code"
  ]
}
```

**Benefits**:
- LLMs can provide better feedback to users
- Helps users understand what happened
- Enables better workflows
- Actionable next steps

---

## Implementation Priority

### Phase 0: Pre-Release Changes (BEFORE 1.0)

**Timeline**: Before first NuGet publish

#### Breaking Changes (High Priority)

1. **✅ CRITICAL** - `batchId` → `sessionId` rename
   - Effort: 1-2 days
   - Impact: Better terminology
   - Changes: ~30 files
   - **BREAKING**: Tool names and parameters change
   
2. **✅ CRITICAL** - Parameter naming standardization
   - Effort: 1 day
   - Impact: Consistency
   - Changes: All tool files
   - **BREAKING**: `excelPath` → `filePath`, `sheetName` → `worksheetName`

#### Non-Breaking Enhancements (Can be done pre-1.0)

3. **✅ HIGH PRIORITY** - Implement Resources with URI patterns
   - Effort: 2-3 days
   - Impact: VS Code integration, better UX
   - Changes: New resource provider classes
   - **ADDITIVE**: Not breaking (new feature)
   - Reference: See MCP-ENHANCEMENT-PROPOSAL.md for details
   
4. **✅ HIGH PRIORITY** - Implement Completions
   - Effort: 1-2 days
   - Impact: Auto-suggest for actions, parameters
   - Changes: New completion handler
   - **ADDITIVE**: Not breaking (new feature)
   - Reference: See PROMPTS-AND-COMPLETIONS-IMPLEMENTATION-GUIDE.md
   
5. **✅ HIGH PRIORITY** - Add Prompts (7 additional prompts)
   - Effort: 2-3 hours
   - Impact: LLM education, better workflows
   - Changes: New prompt files
   - **ADDITIVE**: Not breaking (new feature)
   - Reference: See PROMPTS-AND-COMPLETIONS-IMPLEMENTATION-GUIDE.md
   
6. **✅ MEDIUM PRIORITY** - Error response format standardization
   - Effort: 2-3 days
   - Impact: Better error handling
   - Changes: Core + MCP layers
   - **BREAKING**: Error format changes
   
7. **✅ MEDIUM PRIORITY** - Remove redundant validation attributes
   - Effort: 1 day
   - Impact: Cleaner code
   - Changes: Tool files
   - **NON-BREAKING**: Internal cleanup
   
8. **✅ LOW PRIORITY** - Rich metadata in responses
   - Effort: 2-3 days
   - Impact: Better UX
   - Changes: All tools
   - **NON-BREAKING**: Adds fields to existing responses

### Phase 1: Post-1.0 Enhancements (AFTER 1.0 Launch)

**These are future enhancements that can be added later:**
- Sampling support (when SDK supports it)
- Additional resource types (if needed)
- Enhanced completion logic (file system scanning)
- Icons for prompts/resources

---

## Breaking Changes Checklist

Before releasing 1.0, ensure:

- [ ] All `batchId` references changed to `sessionId`
- [ ] All tools renamed: `begin_excel_batch` → `begin_excel_session`, etc.
- [ ] All `excelPath` → `filePath`
- [ ] All `sheetName` → `worksheetName`
- [ ] Error response format standardized
- [ ] Error codes defined and documented
- [ ] Validation attributes cleaned up
- [ ] Rich metadata added to all tool responses
- [ ] All prompts updated with new terminology
- [ ] All documentation updated
- [ ] All tests updated
- [ ] BATCH-SESSION-GUIDE.md → SESSION-GUIDE.md
- [ ] README.md updated with new terminology

---

## Migration Guide (For Internal Use)

### Find/Replace Changes

```bash
# Terminal commands for bulk rename
rg "batchId" -l | xargs sed -i 's/batchId/sessionId/g'
rg "excelPath" -l | xargs sed -i 's/excelPath/filePath/g'
rg "sheetName" -l | xargs sed -i 's/sheetName/worksheetName/g'
rg "begin_excel_batch" -l | xargs sed -i 's/begin_excel_batch/begin_excel_session/g'
rg "commit_excel_batch" -l | xargs sed -i 's/commit_excel_batch/end_excel_session/g'
```

**Important**: Review each change manually - some contexts might need different handling.

---

## What NOT to Change

### ✅ Keep These Patterns (Already Good)

1. **Action-based tools** - Correct MCP pattern
2. **Resource-based organization** - Already optimal
3. **`excel_` tool prefix** - Essential for multi-server environments (see below)
4. **Batch/session architecture** - Core concept is sound (just rename)
5. **Tool count (9 tools)** - Right balance
6. **Description attributes** - Very helpful for LLMs
7. **JSON responses** - Standard for MCP C# SDK
8. **Stdio transport** - Correct for local use case

### 📌 Why Keep `excel_` Prefix? (Important!)

**Question**: Why not just `worksheet` instead of `excel_worksheet`?

**Answer**: The `excel_` prefix is **critical** for these reasons:

#### 1. Multi-Server Namespace Collision Prevention
Users commonly run **multiple MCP servers** simultaneously in VS Code:
- `mcp-excel` (your server)
- `mcp-google-sheets` (Google Sheets integration)
- `mcp-notion` (Notion database integration)
- `mcp-postgres` (PostgreSQL integration)

**Without prefix (CONFLICTS):**
```typescript
// Which server handles "worksheet"?
worksheet({ action: "read", ... })  // Excel? Sheets? Notion? ❌ AMBIGUOUS
table({ action: "query", ... })     // Excel? Database? ❌ AMBIGUOUS
```

**With prefix (CLEAR):**
```typescript
excel_worksheet({ action: "read", ... })   // ✅ Excel server
sheets_worksheet({ action: "read", ... })  // ✅ Google Sheets server
notion_database({ action: "query", ... })  // ✅ Notion server
postgres_table({ action: "query", ... })   // ✅ Database server
```

#### 2. MCP Specification Best Practice
From MCP spec:
> "Tool names should be unique across all servers a client might connect to. Prefixing with the server's domain is recommended."

**Examples from MCP ecosystem:**
- `github_create_issue`, `github_create_pr` (not `create_issue`)
- `filesystem_read`, `filesystem_write` (not `read`, `write`)
- `slack_send_message` (not `send_message`)

#### 3. IDE Autocomplete Grouping
When user types `excel` in VS Code Copilot, all tools appear grouped:
```
excel_cell
excel_connection
excel_datamodel
excel_file
excel_parameter
excel_powerquery
excel_vba
excel_worksheet
```

Without prefix, tools would be scattered alphabetically (poor UX).

#### 4. Package/Command Naming Consistency
- Package: `Sbroenne.ExcelMcp.McpServer`
- Command: `mcp-excel`
- Tools: `excel_*` (matches package domain)

#### 5. Future-Proofing
Generic names like `worksheet`, `cell`, `file` would conflict with:
- Future Google Sheets MCP server
- Future LibreOffice Calc MCP server
- Any spreadsheet/database MCP integration

**Conclusion**: The `excel_` prefix adds **zero** functional overhead but provides **critical** namespace isolation and follows **MCP best practices**. Keep it!

---

## Risk Assessment

### Low Risk Changes ✅
- Parameter renaming (sessionId, filePath, worksheetName)
- Error format standardization
- Metadata additions
- Validation cleanup

### Medium Risk Changes ⚠️
- Tool renames (begin_excel_session, etc.)
- URI pattern for resources (new feature)

### High Risk Changes 🚨
- None! All changes are straightforward

**Overall Risk**: **LOW** - since nothing is released yet, we can change freely

---

## Recommendation

**DO THESE BEFORE 1.0 RELEASE**:

1. ✅ `batchId` → `sessionId` (clearer terminology)
2. ✅ `excelPath` → `filePath` (more generic)
3. ✅ `sheetName` → `worksheetName` (more explicit)
4. ✅ Standardize error responses with error codes
5. ✅ Clean up validation attributes
6. ✅ Add rich metadata to responses

**TOTAL EFFORT**: ~5-7 days  
**BENEFIT**: Cleaner API, better UX, less technical debt

**DON'T DO**:
- ❌ Split tools into 60+ separate tools (current pattern is correct)
- ❌ Change fundamental architecture (it's already good)
- ❌ Remove backward compatibility concerns (there is none yet!)

---

## Next Steps

1. **Approve this proposal** ✅
2. **Create feature branch**: `feature/pre-1.0-breaking-changes`
3. **Implement changes in priority order**
4. **Update all tests**
5. **Update all documentation**
6. **Test with VS Code GitHub Copilot**
7. **Release 1.0.0 with clean API**

---

## Conclusion

Since the MCP server hasn't been released, we have a **golden opportunity** to make these improvements **without breaking anything**. The changes are relatively small but will result in:

- 🎯 **Clearer terminology** (session vs batch)
- 📝 **More consistent API** (parameter naming)
- 🚨 **Better error handling** (standard error codes)
- 📊 **Richer responses** (metadata + suggestions)
- 🧹 **Cleaner code** (less validation clutter)

**Total effort**: ~1 week of focused work  
**Payoff**: Years of better API design

**Recommendation**: ✅ **DO IT NOW** before first release!

---

**Author**: GitHub Copilot  
**Context**: Pre-1.0 breaking changes opportunity  
**Date**: October 27, 2025
