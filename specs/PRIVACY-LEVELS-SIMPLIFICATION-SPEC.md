# Privacy Levels Simplification Specification

> **Status**: Proposal  
> **Created**: 2025-10-30  
> **Author**: LLM Analysis based on Microsoft Docs research

## Executive Summary

Privacy levels in Excel Power Query are an **edge case security feature** that most users never encounter. Based on research and LLM perspective, the current implementation is overcomplicated for ExcelMcp's use case (AI-assisted Excel development, primarily creating NEW files). This spec proposes radical simplification.

---

## Background Research

### What Privacy Levels Actually Are

**Purpose**: Prevent unintentional data leakage between sources during query folding (Formula.Firewall).

**Four Levels** (per [Microsoft Docs](https://learn.microsoft.com/en-us/power-query/privacy-levels)):
- **None** - Ignores privacy (least secure, testing only)
- **Private** - No data sharing with ANY other source (most secure)
- **Organizational** - Data shared within organization, not public
- **Public** - Data can be shared with anyone

**When They Matter**:
1. Combining data from **multiple different sources** (e.g., SQL + Web)
2. Query folding scenarios where Excel pushes operations to source
3. Prevents passing data from Private source to Public API unintentionally

**When They DON'T Matter**:
1. Single data source queries (90% of Power Query usage)
2. Simple table transformations
3. Files created from scratch (no existing privacy settings)

### How Excel Actually Handles This

**UI Behavior** (per Microsoft Docs):
1. **Global Setting** (File → Options → Privacy):
   - "Combine data according to Privacy Level settings" (DEFAULT)
   - "Ignore Privacy Levels and potentially improve performance" (Fast Combine)

2. **Per-File Setting**:
   - Same options as global
   - Overrides global when set

3. **Per-Data-Source Setting**:
   - Each connection/query can have its own privacy level
   - Stored in connection metadata

4. **Dialog Prompts**:
   - Excel shows dialog when privacy mismatch detected
   - User must explicitly choose level OR enable "Ignore Privacy Levels"

### What ExcelMcp Currently Does (Broken)

**Current Implementation** (`PowerQueryCommands.cs:160-240`):
```csharp
private static void ApplyPrivacyLevel(dynamic workbook, PowerQueryPrivacyLevel privacyLevel)
{
    // 1. Try to set via CustomDocumentProperties (doesn't work - not official API)
    customProps.Add("PowerQueryPrivacyLevel", false, 4, privacyValue);
    
    // 2. Try to disable DisplayAlerts (not related to privacy)
    application.DisplayAlerts = false;
    
    // 3. Catch all exceptions and silently fail
    catch (Exception) { /* best-effort */ }
}
```

**Problems**:
1. ❌ Uses custom properties (not official Excel COM API)
2. ❌ No actual COM API for setting privacy levels programmatically
3. ❌ Settings don't persist or affect Excel behavior
4. ❌ Complex detection/recommendation logic (`DetectPrivacyLevelsAndRecommend`)
5. ❌ CLI parameter `--privacy-level` doesn't actually work
6. ❌ MCP Server parameter `privacyLevel` is placebo

---

## LLM Perspective: What We Actually Need

**As an LLM using ExcelMcp, I care about**:
1. ✅ Creating queries that work immediately
2. ✅ Not getting Formula.Firewall errors that block my workflow
3. ✅ Simple, predictable behavior
4. ❌ I DON'T care about enterprise data governance (that's user's responsibility)
5. ❌ I DON'T need to configure security settings programmatically

**Typical LLM Workflow**:
```
1. Create new Excel file
2. Import Power Query (usually single source: CSV, JSON, Web)
3. Transform data
4. Load to worksheet
5. Done
```

**Privacy level relevance**: ZERO in 95% of cases.

**When it DOES matter**:
- User explicitly combines SQL + Web data
- User hits Formula.Firewall error
- **User manually fixes in Excel UI** (that's the correct workflow)

---

## Official COM API Research

### ❌ No Programmatic Privacy Level API Exists

**Searched**: Microsoft Office VBA API, Excel Object Model, Power Query COM

**Findings**:
1. Privacy levels are **UI-only setting** in Excel
2. No `Workbook.PrivacyLevel` property
3. No `Connection.PrivacyLevel` property
4. No `Query.PrivacyLevel` property
5. Only option: **User must set manually in Excel UI**

**Microsoft's Official Guidance** ([Privacy Levels Docs](https://learn.microsoft.com/en-us/power-query/privacy-levels#setting-the-privacy-level-options)):
> "Privacy levels are set through Excel's Options dialog or per-connection in Data Source Settings"

**No programmatic API mentioned anywhere.**

### ✅ What DOES Work (Sort of)

**Fast Combine Option** (Global Setting):
- Can be enabled via registry/group policy for testing
- ExcelMcp has NO business modifying user's global Excel settings
- Security violation to auto-enable

**Per-Query M Code**:
```powerquery
// M code can wrap data sources with privacy level functions
let
    Source = Privacy.Private(Web.Contents("https://api.example.com"))
in
    Source
```

**BUT**: This is embedded in M code, not a workbook/connection property.

---

## Proposed Solution: Radical Simplification

### Phase 1: Deprecate Privacy Level Parameters (IMMEDIATE)

**Remove/Ignore**:
1. ❌ `--privacy-level` CLI parameter
2. ❌ `privacyLevel` parameter in Core commands
3. ❌ `privacyLevel` parameter in MCP tools
4. ❌ `PowerQueryPrivacyLevel` enum (keep for backward compat, mark obsolete)
5. ❌ `ApplyPrivacyLevel()` method (delete)
6. ❌ `DetectPrivacyLevelsAndRecommend()` method (delete)
7. ❌ `PowerQueryPrivacyErrorResult` class (delete)

**Keep** (for documentation only):
- `PowerQueryPrivacyLevel` enum (mark `[Obsolete]`)
- Documentation explaining why we don't support this

### Phase 2: Update Documentation (IMMEDIATE)

**Add to COMMANDS.md**:
```markdown
## Power Query Privacy Levels - NOT SUPPORTED

**Why ExcelMcp doesn't manage privacy levels:**

1. **No COM API**: Excel does not expose privacy levels via COM interop
2. **Edge case**: 95% of queries use single data source (privacy irrelevant)
3. **User responsibility**: Security settings should be user-controlled, not automated
4. **Simple workaround**: If you get Formula.Firewall error:
   - Open file in Excel
   - File → Options → Privacy → "Ignore Privacy Levels" (for testing)
   - OR set privacy level per connection in Excel UI

**For production workbooks with privacy requirements:**
- Configure privacy levels manually in Excel UI (one-time setup)
- ExcelMcp respects existing privacy settings
- Settings persist in .xlsx file
```

**Add to MCP Prompts**:
```markdown
Privacy levels are user-managed in Excel UI. If you encounter Formula.Firewall 
errors, instruct user to configure privacy levels in Excel manually.
```

### Phase 3: Minimal Error Handling (NEW)

**Detect Formula.Firewall Errors**:
```csharp
catch (COMException ex) when (ex.Message.Contains("Formula.Firewall"))
{
    return new OperationResult
    {
        Success = false,
        ErrorMessage = "Privacy level error detected. This query combines data from " +
                      "multiple sources. Open the file in Excel and configure privacy " +
                      "levels: File → Options → Privacy. See COMMANDS.md for details.",
        WorkflowHint = "Privacy levels must be configured manually in Excel UI"
    };
}
```

**No detection, no recommendation, no parameters - just clear error message.**

---

## Migration Plan

### Breaking Changes (Acceptable)

**CLI**:
- `--privacy-level` parameter **REMOVED** (no longer accepted)
- Old commands: `pq-import file.xlsx Query1 query.pq --privacy-level Private`
- New commands: `pq-import file.xlsx Query1 query.pq`
- **Impact**: Tests using `--privacy-level` will fail

**MCP Server**:
- `privacyLevel` parameter **REMOVED** from all tools
- Old: `excel_powerquery({ action: "import", privacyLevel: "Private" })`
- New: `excel_powerquery({ action: "import" })`
- **Impact**: LLM prompts mentioning privacy levels need update

**Core**:
- `PowerQueryPrivacyLevel?` parameter **REMOVED** from all methods
- Old: `ImportAsync(batch, name, file, privacyLevel)`
- New: `ImportAsync(batch, name, file)`
- **Impact**: No external users of Core library

### Code Removal (Files to Modify)

**1. Core Layer** (`src/ExcelMcp.Core/Commands/PowerQueryCommands.cs`):
```csharp
// DELETE METHODS (lines 29-240):
- DetectPrivacyLevelFromMCode()
- DetermineRecommendedPrivacyLevel()
- GeneratePrivacyExplanation()
- DetectPrivacyLevelsAndRecommend()
- ApplyPrivacyLevel()

// UPDATE METHOD SIGNATURES (remove privacyLevel parameter):
- ImportAsync(IExcelBatch batch, string queryName, string mCodeFile)  // was: ..., PowerQueryPrivacyLevel? privacyLevel = null
- UpdateAsync(...)  // same
- SetLoadToTableAsync(...)  // same
- SetLoadToDataModelAsync(...)  // same
- SetLoadToBothAsync(...)  // same

// DELETE CATCH BLOCKS for privacy errors (lines vary)
```

**2. Core Models** (`src/ExcelMcp.Core/Models/ResultTypes.cs`):
```csharp
// MARK OBSOLETE (line 744):
[Obsolete("Privacy levels are not supported by ExcelMcp. Manage via Excel UI.")]
public enum PowerQueryPrivacyLevel { None, Private, Organizational, Public }

// DELETE CLASS (lines 775-795):
public class PowerQueryPrivacyErrorResult : OperationResult { ... }

// DELETE CLASS (after PowerQueryPrivacyErrorResult):
public class QueryPrivacyInfo { ... }
```

**3. CLI Layer** (`src/ExcelMcp.CLI/Commands/PowerQueryCommands.cs`):
```csharp
// DELETE METHOD (lines 25-54):
private static PowerQueryPrivacyLevel? ParsePrivacyLevel(string[] args)

// DELETE METHOD (lines 64-85):
private static void DisplayPrivacyConsentPrompt(PowerQueryPrivacyErrorResult error)

// UPDATE ALL COMMAND METHODS - Remove privacy parsing:
public async Task<int> Import(string[] args)
{
    // DELETE:
    var privacyLevel = ParsePrivacyLevel(args);
    
    // UPDATE call:
    var result = await _coreCommands.ImportAsync(batch, queryName, mCodeFile);  // no privacyLevel
    
    // DELETE privacy error handling
}
```

**4. MCP Server** (`src/ExcelMcp.McpServer/Tools/ExcelPowerQueryTool.cs`):
```csharp
// UPDATE TOOL SIGNATURE - Remove privacyLevel parameter from:
- ExcelPowerQuery() method signature
- All action handlers (import, update, load-to-table, etc.)

// DELETE parameter documentation from server.json
```

**5. Tests** (`tests/ExcelMcp.Core.Tests/Integration/Commands/PowerQueryPrivacyLevelTests.cs`):
```csharp
// DELETE ENTIRE FILE (300+ lines)
// This file tests functionality we're removing
```

**6. Test Data** (various PowerShell scripts):
```powershell
# UPDATE: Remove --privacy-level flags from:
- tests/ExcelMcp.Core.Tests/TestData/create-simple-datamodel.ps1
- tests/ExcelMcp.Core.Tests/TestData/create-datamodel-testfile.ps1
```

**7. Documentation** (`docs/COMMANDS.md`, `docs/INSTALLATION.md`):
```markdown
# DELETE sections about privacy levels
# ADD new "Privacy Levels - NOT SUPPORTED" section (see Phase 2 above)
```

---

## Implementation Checklist

- [ ] **Core**: Remove privacy methods and parameters
- [ ] **Core**: Mark `PowerQueryPrivacyLevel` enum as `[Obsolete]`
- [ ] **Core**: Delete `PowerQueryPrivacyErrorResult` class
- [ ] **Core**: Add Formula.Firewall error detection with helpful message
- [ ] **CLI**: Remove `--privacy-level` parsing
- [ ] **CLI**: Remove privacy error prompt logic
- [ ] **CLI**: Update all command methods
- [ ] **MCP**: Remove `privacyLevel` parameters from tools
- [ ] **MCP**: Update server.json schema
- [ ] **MCP**: Update prompts to remove privacy references
- [ ] **Tests**: Delete `PowerQueryPrivacyLevelTests.cs`
- [ ] **Tests**: Update test scripts (remove --privacy-level flags)
- [ ] **Tests**: Verify all Core tests pass without privacy parameters
- [ ] **Tests**: Verify all CLI tests pass
- [ ] **Tests**: Verify all MCP tests pass
- [ ] **Docs**: Add "Privacy Levels - NOT SUPPORTED" section to COMMANDS.md
- [ ] **Docs**: Add workaround instructions for Formula.Firewall errors
- [ ] **Docs**: Update INSTALLATION.md (remove privacy env var references)
- [ ] **Copilot Instructions**: Update to reflect simplified approach

---

## Rationale Summary

| Aspect | Current (Broken) | Proposed (Simplified) |
|--------|------------------|----------------------|
| **Parameters** | `--privacy-level`, `privacyLevel` | None |
| **Code Complexity** | ~300 lines detection/recommendation | ~10 lines error detection |
| **User Experience** | Fake parameter gives false confidence | Honest: "Use Excel UI" |
| **Security** | Pretends to set privacy (doesn't work) | Defers to Excel's real settings |
| **Edge Cases** | Complex detection logic (unreliable) | Clear error message + docs |
| **Maintenance** | High (broken code needs fixes) | Zero (no code to maintain) |
| **LLM Utility** | Confusing parameter that doesn't work | Clear: handle via Excel UI |

---

## Expected Outcomes

**After Implementation**:
1. ✅ **Simpler API**: No privacy parameters anywhere
2. ✅ **Honest Documentation**: Explains Excel UI is the correct approach
3. ✅ **Less Code**: ~300 lines deleted
4. ✅ **Fewer Tests**: Delete entire test file
5. ✅ **Clear Errors**: Formula.Firewall errors have helpful message
6. ✅ **Better LLM Experience**: No false impression that privacy is automated

**What We Lose**:
- ❌ Fake parameter that didn't work anyway
- ❌ Complex detection logic that was unreliable
- ❌ False sense of security automation

**What We Gain**:
- ✅ Simpler, more maintainable codebase
- ✅ Honest communication about capabilities
- ✅ Correct guidance to use Excel UI for security settings
- ✅ Alignment with Microsoft's actual API capabilities

---

## Alternative Considered: M Code Wrapper Approach

**Could we embed privacy in M code?**

```powerquery
let
    Source = Privacy.Private(Web.Contents("https://api.com"))
in
    Source
```

**Why NOT**:
1. Modifying user's M code is invasive
2. Only works for new queries, not updates
3. Privacy level is data source property, not query property
4. Still requires user to understand privacy implications
5. Doesn't solve existing query privacy mismatches

**Verdict**: More complexity for marginal benefit. Better to delegate to Excel UI.

---

## Compliance Considerations

**Q: Does removing privacy support violate security best practices?**

**A: No.** ExcelMcp is a **development tool**, not a runtime data processing engine. Privacy levels are:
- ✅ Still enforced by Excel (Formula.Firewall)
- ✅ Still configurable by users (Excel UI)
- ✅ Still persisted in .xlsx files
- ✅ Not bypassable via COM API

**What changes**: We stop **pretending** to programmatically set privacy levels when we actually can't.

---

## References

- [Microsoft Docs: Privacy Levels](https://learn.microsoft.com/en-us/power-query/privacy-levels)
- [Microsoft Docs: Data Privacy Firewall](https://learn.microsoft.com/en-us/power-query/data-privacy-firewall)
- [Microsoft Docs: Security Best Practices](https://learn.microsoft.com/en-us/power-query/security-best-practices-power-query)
- [Excel VBA API Reference](https://learn.microsoft.com/en-us/office/vba/api/overview/excel) (no privacy level API)

---

## Conclusion

Privacy levels are an **edge case security feature** with **no programmatic COM API**. The current ExcelMcp implementation is **broken placebo code** that provides false confidence. 

**Recommended Action**: DELETE privacy level parameters and logic entirely. Replace with honest documentation and clear error messages directing users to Excel UI.

**User Impact**: POSITIVE - clearer expectations, simpler API, correct security guidance.

**LLM Impact**: POSITIVE - no confusing parameters, honest about limitations, better error messages.

**Maintenance Impact**: POSITIVE - 300 fewer lines of broken code to maintain.
