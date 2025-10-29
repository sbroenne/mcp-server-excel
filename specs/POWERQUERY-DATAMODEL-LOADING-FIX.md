# Power Query Data Model Loading Fix - Issues #42 and #64

> **Status:** BROKEN - Feature Never Worked  
> **User Reports:** Issue #42 (original), Issue #64 (2025-10-29 production error)  
> **Test File:** `DataModelLoadingIssueTests.cs` **EXPECTS FAILURE**

## Latest Production Error (2025-10-29)

```json
{
  "action": "set-load-to-data-model",
  "QueryName": "Milestones",
  "ConfigurationApplied": false,
  "DataLoadedToModel": false,
  "ErrorMessage": "Failed to configure query for Data Model loading",
  "Success": false
}
```

**Test Evidence:** Test passes with `Assert.False(setDataModelResult.Success)` - **EXPECTS FAILURE!**

## Problem Statement

When using `set-load-to-data-model` action with Power Query, the configuration appears to succeed but:
1. Queries remain connection-only (`IsConnectionOnly: true`)
2. Data doesn't actually load into the Power Pivot Data Model
3. Requires undocumented additional steps (like `excel_datamodel refresh`)
4. Workflow is non-intuitive for AI agents and users

## Root Cause Analysis

### Current Implementation Issues

1. **`SetLoadToDataModelAsync` doesn't actually load data**
   - Sets configuration flags but doesn't trigger data load
   - Uses fallback to named range markers when COM methods fail
   - No actual refresh happens after configuration change

2. **Multiple approaches but all unreliable**
   ```csharp
   // Approach 1: LoadToWorksheetModel property (may not exist)
   query.LoadToWorksheetModel = true;
   
   // Approach 2: Connection settings (doesn't load data)
   connection.RefreshOnFileOpen = false;
   connection.BackgroundQuery = false;
   
   // Approach 3: Named range marker fallback (doesn't load data)
   names.Add("DataModel_Query_{queryName}", "=Sheet1!$A$1");
   ```

3. **Workflow guidance doesn't mention the gap**
   - Current hint: "WORKFLOW: Configure → Refresh → Data available in PowerPivot"
   - But doesn't clarify that configuration ≠ data loading
   - Doesn't mention Data Model refresh may be required

### Existing Commands Review - No Duplication Found ✅

**PowerQueryCommands.RefreshAsync:**
- Refreshes query via `connection.Refresh()` or `queryTable.Refresh()`
- Only refreshes the QUERY data (loads to worksheet if configured)
- Does NOT refresh Data Model if query is set to load-to-data-model
- Returns `PowerQueryRefreshResult` with error handling

**DataModelCommands.RefreshAsync:**
- Refreshes Data Model via `model.Refresh()` or `table.Refresh()`
- Refreshes ALL tables in Data Model OR specific table
- Reloads data from Power Query sources into Data Model
- Returns `OperationResult` with suggested actions

**Key Distinction:**
- `PowerQueryCommands.RefreshAsync` = Refresh query definition/connection
- `DataModelCommands.RefreshAsync` = Refresh data IN the Data Model
- BOTH may be needed: Query refresh updates M code, Model refresh loads data

**No Duplication - Commands are Complementary:**
- PowerQuery refresh: Updates query execution, loads to worksheet
- DataModel refresh: Loads query results into Power Pivot Model
- SetLoadToDataModel: Configures WHERE data goes (worksheet, model, both)
- Our fix will call BOTH when needed for atomic operation

### What Actually Needs to Happen

For data to appear in Power Pivot Data Model:
1. Query must be configured to load to data model (`query.LoadToWorksheetModel = true`)
2. Query must be refreshed via `connection.Refresh()` (PowerQueryCommands)
3. Data Model must be refreshed via `model.Refresh()` or `table.Refresh()` (DataModelCommands)
4. Workbook must save changes (or changes take effect on reopen)

**Critical Discovery:**
- Setting `LoadToWorksheetModel` only configures the DESTINATION
- Query refresh loads data but may not commit to Data Model immediately
- Data Model refresh is what actually imports query results into Power Pivot
- This requires BOTH PowerQueryCommands.RefreshAsync AND DataModelCommands.RefreshAsync

## Proposed Solutions

### Solution 1: Make `SetLoadToDataModelAsync` Atomic (RECOMMENDED) ⚠️ BREAKING CHANGE

**🔥 Breaking Change: Always auto-refresh, remove parameter**

**New Signature:**
```csharp
public async Task<PowerQueryLoadToDataModelResult> SetLoadToDataModelAsync(
    IExcelBatch batch, 
    string queryName, 
    PowerQueryPrivacyLevel? privacyLevel = null)
    // REMOVED autoRefresh parameter - ALWAYS refreshes atomically
{
    // 1. Configure load mode (LoadToWorksheetModel = true)
    // 2. ALWAYS refresh (no conditional):
    //    a. Call DataModelCommands.RefreshAsync to load to model
    // 3. Verify data actually loaded to Data Model
    // 4. Return detailed result with DataLoadedToModel status
}
```

**Why Remove autoRefresh Parameter:**
- Simpler API - one action always does the complete job
- Matches user expectations - "set-load-to-data-model" should LOAD the data
- No confusion about when to use true vs false
- Batch operations can use `SetConnectionOnlyAsync` then manual refresh if needed

**Dependencies:**
- Will use existing `DataModelCommands.RefreshAsync` (no duplication)
- Leverages `DataModelHelpers.HasDataModel` for verification

**New Result Type (BREAKING):**
```csharp
public class PowerQueryLoadToDataModelResult : OperationResult
{
    public bool ConfigurationApplied { get; set; }
    public bool DataLoadedToModel { get; set; }
    public int RowsLoaded { get; set; }  // NEW - actual row count
    public int TablesInDataModel { get; set; }
    public string WorkflowStatus { get; set; } // "Complete" | "Failed" | "Partial"
    // REMOVED AutoRefreshTriggered - always true now
}
```

### Solution 2: Simplify All Load Mode Actions ⚠️ BREAKING CHANGES

**Make ALL load-mode setters atomic (consistent API):**

```csharp
// All return PowerQueryLoadConfigResult with actual verification

public async Task<PowerQueryLoadConfigResult> SetLoadToTableAsync(
    IExcelBatch batch,
    string queryName,
    string sheetName,
    PowerQueryPrivacyLevel? privacyLevel = null)
    // ALWAYS refreshes and loads to worksheet atomically

public async Task<PowerQueryLoadToDataModelResult> SetLoadToDataModelAsync(
    IExcelBatch batch,
    string queryName,
    PowerQueryPrivacyLevel? privacyLevel = null)
    // ALWAYS refreshes and loads to Data Model atomically

public async Task<PowerQueryLoadConfigResult> SetLoadToBothAsync(
    IExcelBatch batch,
    string queryName,
    string sheetName,
    PowerQueryPrivacyLevel? privacyLevel = null)
    // ALWAYS refreshes and loads to both destinations atomically

public async Task<OperationResult> SetConnectionOnlyAsync(
    IExcelBatch batch,
    string queryName)
    // Just configures, no refresh needed (no data loading)
```

**Consistency Benefits:**
- Every "set-load-to-X" actually loads the data
- Users never confused about next steps
- "set-connection-only" is the ONLY one that doesn't load
- Clear naming: "SetLoadTo" = configures AND loads

### Solution 3: Enhanced Workflow Guidance

**Update `PowerQueryWorkflowGuidance.cs`:**

```csharp
public static List<string> GetNextStepsAfterLoadConfig(string loadMode, bool dataLoaded)
{
    if (!dataLoaded)
    {
        return new List<string>
        {
            $"Query configured to load as: {loadMode}",
            "⚠️ Configuration set but data not yet loaded",
            "Use 'refresh' to load data to configured destination",
            "Or use 'excel_datamodel refresh' to load all queries to Data Model",
            "Or reopen workbook in Excel to apply changes",
            "Use 'excel_datamodel list-tables' to verify data loaded"
        };
    }
    
    return new List<string>
    {
        $"Query configured and data loaded to: {loadMode}",
        "Data Model now contains this query's data",
        "Use 'excel_datamodel list-tables' to verify",
        "Use 'excel_datamodel list-measures' to create DAX calculations"
    };
}

public static string GetWorkflowHint(string operation, bool success, bool dataLoaded = false)
{
    return operation switch
    {
        "pq-set-load-to-data-model" when dataLoaded => 
            "COMPLETE: Configuration applied and data loaded to Power Pivot",
        "pq-set-load-to-data-model" when !dataLoaded => 
            "PARTIAL: Configuration set but data not loaded. Use 'refresh' or 'excel_datamodel refresh'",
        // ... other cases
    };
}
```

## Implementation Plan

### Phase 1: Core Logic Improvements (This PR)

**1. Add DataModelCommands dependency to PowerQueryCommands**
- [ ] Update PowerQueryCommands constructor to accept IDataModelCommands
- [ ] Store as private field `_dataModelCommands`
- [ ] Update all instantiations (CLI, MCP Server, Tests)

**2. Enhance `TrySetQueryLoadToDataModel`**
- [ ] Add actual refresh after configuration
- [ ] Verify data loaded to Data Model using DataModelHelpers
- [ ] Return detailed status

**3. Create `PowerQueryLoadToDataModelResult`**
- [ ] Add new result type with detailed fields
- [ ] Update `SetLoadToDataModelAsync` to return this type

**4. Add `autoRefresh` parameter**
- [ ] ~~Default to `true` for better UX~~
- [ ] ~~Allow disabling for batch operations~~
- [x] **DECISION: Don't add parameter - always refresh atomically**
- [x] **Simpler API, clearer expectations**

**5. Implement atomic refresh logic**
- [ ] ~~If autoRefresh=true~~, call `_dataModelCommands.RefreshAsync(batch)`
- [ ] ALWAYS call `_dataModelCommands.RefreshAsync(batch)` after configuration
- [ ] Verify query appears in Data Model tables
- [ ] Count rows loaded
- [ ] Confirm data actually present

**6. Improve verification logic**
- [ ] Use `DataModelHelpers.HasDataModel(workbook)`
- [ ] Check if query appears in `model.ModelTables`
- [ ] Get row count from table
- [ ] Return in result object

### Phase 2: MCP Server Updates

**5. Update `ExcelPowerQueryTool.cs`**
- [ ] Handle new `PowerQueryLoadToDataModelResult` type
- [ ] ~~Add `autoRefresh` parameter to MCP action~~
- [x] **NO parameter changes - always atomic**
- [ ] Update JSON serialization for new result fields
- [ ] Add `rowsLoaded` to response

**6. Update workflow guidance**
- [ ] Remove "use refresh" from SuggestedNextActions (no longer needed)
- [ ] Update to "Data loaded successfully" messaging
- [ ] Clarify atomic operation in WorkflowHint
- [ ] Update all load-mode actions consistently

**7. Update MCP prompts/completions**
- [ ] Remove suggestions to call refresh after set-load-to-data-model
- [ ] Update workflow examples to show one-step operation

### Phase 3: CLI Updates

**7. Update `PowerQueryCommands.cs` (CLI)**
- [ ] Display enhanced result information (rowsLoaded, workflowStatus)
- [ ] ~~Add `--auto-refresh` flag~~
- [x] **NO flag needed - always atomic**
- [ ] Update help text to reflect atomic operation
- [ ] Update output formatting for new result type

**8. Update documentation**
- [ ] Update COMMANDS.md with breaking changes
- [ ] Add migration examples (old → new)
- [ ] Document batch operation strategies
- [ ] ~~Document when to use `--auto-refresh=false`~~
- [x] **Document batch operations use set-connection-only + manual refresh**

### Phase 4: Testing

**9. Add comprehensive tests**
- [ ] Test ~~`autoRefresh=true`~~ atomic operation actually loads data
- [ ] ~~Test `autoRefresh=false` only sets configuration~~
- [x] **All operations are atomic - no conditional testing needed**
- [ ] Test verification logic (rowsLoaded, tablesInDataModel)
- [ ] Test Data Model table counting
- [ ] Test workflow guidance updates
- [ ] Test breaking changes don't affect other operations

**10. Integration test scenarios**
- [ ] Import → SetLoadToDataModel → Verify in Data Model (one call)
- [ ] ~~Import → SetLoadToDataModel (manual) → Refresh → Verify~~
- [x] **No manual refresh scenarios - all atomic**
- [ ] Multiple queries → SetConnectionOnly → Manual refresh (batch optimization)
- [ ] Existing query → SetLoadToDataModel → Auto-upgrade to LoadToBoth
- [ ] Connection-only query → SetLoadToDataModel → Verify upgrade

## Verification Methods

### Method 1: Check Model.ModelTables (PREFERRED)
```csharp
dynamic model = workbook.Model;
dynamic tables = model.ModelTables;
bool queryExistsInModel = false;
int rowCount = 0;

for (int i = 1; i <= tables.Count; i++)
{
    dynamic table = tables.Item(i);
    if (table.Name == queryName)
    {
        queryExistsInModel = true;
        rowCount = table.RowCount;
        break;
    }
}

// Return in result:
result.DataLoadedToModel = queryExistsInModel && rowCount > 0;
result.TablesInDataModel = tables.Count;
```

### Method 2: Check Query.LoadedToDataModel Property
```csharp
dynamic query = FindQuery(workbook, queryName);
bool isLoadedToModel = query.LoadedToDataModel;  // May not exist in all Excel versions

// CAUTION: This property may return true even if data isn't actually loaded yet
// Use Method 1 for actual verification
```

### Method 3: Query Data Model Table List (FALLBACK)
```csharp
var dataModelCommands = new DataModelCommands();
var tablesResult = await dataModelCommands.ListTablesAsync(batch);
bool queryInDataModel = tablesResult.Tables.Any(t => t.Name == queryName);

// CAUTION: This creates a separate batch - may have timing issues
```

## Edge Cases & Error Handling

### Edge Case 1: Data Model Not Available
```csharp
// Excel version doesn't support Data Model (pre-2013)
if (!DataModelHelpers.HasDataModel(workbook))
{
    result.Success = false;
    result.ErrorMessage = "Data Model not available in this Excel version. Requires Excel 2013+ with Power Pivot.";
    result.SuggestedNextActions = new List<string>
    {
        "Use 'set-load-to-table' to load to worksheet instead",
        "Upgrade to Excel 2013+ for Data Model support",
        "Check if Power Pivot add-in is enabled"
    };
    return result;
}
```

### Edge Case 2: LoadToWorksheetModel Property Doesn't Exist
```csharp
// Older Excel versions may not have this property
try
{
    query.LoadToWorksheetModel = true;
}
catch (RuntimeBinderException)
{
    // Fall back to alternative method
    result.ConfigurationApplied = false;
    result.ErrorMessage = "Excel version doesn't support LoadToWorksheetModel property.";
    result.WorkflowStatus = "Unsupported - Use Excel 2016+ for this feature";
    return result;
}
```

### Edge Case 3: Query Already Loaded to Worksheet
```csharp
// If query is already loading to a worksheet, what happens?
// Options:
// A) Remove worksheet loading, switch to Data Model only
// B) Set to LoadToBoth mode (worksheet + Data Model)
// C) Error - user must explicitly choose

// RECOMMENDED: Auto-detect and use LoadToBoth
var loadConfig = await GetLoadConfigAsync(batch, queryName);
if (loadConfig.LoadMode == PowerQueryLoadMode.LoadToTable)
{
    // Already loading to worksheet - use LoadToBoth instead
    await SetLoadToBothAsync(batch, queryName, loadConfig.TargetSheet, privacyLevel);
    result.WorkflowHint = "Query already loads to worksheet - configured to load to BOTH worksheet and Data Model";
}
```

### Edge Case 4: Refresh Fails but Configuration Succeeds
```csharp
// Configuration applied, but refresh failed
if (configApplied && !refreshSucceeded)
{
    result.Success = true;  // Configuration succeeded
    result.ConfigurationApplied = true;
    result.DataLoadedToModel = false;
    result.AutoRefreshTriggered = true;
    result.WorkflowStatus = "Configuration Only - Refresh Failed";
    result.ErrorMessage = $"Load mode configured but refresh failed: {refreshError}";
    result.SuggestedNextActions = new List<string>
    {
        "Configuration saved - data will load on workbook reopen",
        "Fix refresh error and call 'excel_datamodel refresh'",
        "Use 'get-load-config' to verify configuration persisted"
    };
}
```

### Edge Case 5: Verification Timing (Excel Processing Delay)
```csharp
// Excel may need time to process Data Model changes
if (autoRefresh)
{
    // Wait briefly for Excel to process
    await Task.Delay(500);
    
    // Verify data actually loaded
    bool dataLoaded = VerifyDataInModel(workbook, queryName);
    
    if (!dataLoaded)
    {
        // Try waiting a bit longer
        await Task.Delay(1500);
        dataLoaded = VerifyDataInModel(workbook, queryName);
    }
    
    result.DataLoadedToModel = dataLoaded;
    
    if (!dataLoaded)
    {
        result.WorkflowHint = "Configuration applied, refresh triggered, but data not yet visible in Data Model. " +
                              "Changes may take effect after workbook save/reopen.";
    }
}
```

## Breaking Changes

### ⚠️ BREAKING CHANGES - Backwards Compatibility NOT Required

**1. SetLoadToDataModelAsync - Always Refreshes**
- **Before:** Only configured load mode, data not loaded
- **After:** Configures AND loads data atomically
- **Migration:** Remove manual `excel_datamodel refresh` calls after this action

**2. Return Type Changed**
- **Before:** Returns `OperationResult`
- **After:** Returns `PowerQueryLoadToDataModelResult` with verification
- **Impact:** MCP Server JSON schema changes, CLI output changes

**3. SetLoadToTableAsync - Also Made Atomic (Consistency)**
- **Before:** May not have refreshed automatically
- **After:** ALWAYS refreshes and loads to worksheet
- **Migration:** Remove manual `refresh` calls after this action

**4. SetLoadToBothAsync - Also Made Atomic**
- **Before:** May not have refreshed automatically
- **After:** ALWAYS refreshes and loads to both destinations
- **Migration:** Remove manual `refresh` calls after this action

**5. Removed autoRefresh Parameter (Not Added)**
- **Decision:** Don't add autoRefresh parameter at all
- **Rationale:** Simpler API, clear expectations, no user confusion
- **For Batch Operations:** Use SetConnectionOnlyAsync then manual refresh

### Migration Strategy

**Old Code:**
```typescript
// Configure (doesn't load data)
excel_powerquery({ action: "set-load-to-data-model", excelPath: "file.xlsx", queryName: "Sales" })
// Manually refresh
excel_datamodel({ action: "refresh", excelPath: "file.xlsx" })
```

**New Code:**
```typescript
// One step - configures AND loads
excel_powerquery({ action: "set-load-to-data-model", excelPath: "file.xlsx", queryName: "Sales" })
// Done! Data is in Data Model
```

**Batch Operations - Old Code:**
```typescript
for (query of queries) {
  excel_powerquery({ action: "set-load-to-data-model", excelPath: "file.xlsx", queryName: query })
  excel_datamodel({ action: "refresh", excelPath: "file.xlsx" })  // Inefficient!
}
```

**Batch Operations - New Code:**
```typescript
// Option A: Let each action refresh (simple but slower)
for (query of queries) {
  excel_powerquery({ action: "set-load-to-data-model", excelPath: "file.xlsx", queryName: query })
  // Each loads its own data
}

// Option B: Batch configure then single refresh (efficient)
for (query of queries) {
  excel_powerquery({ action: "set-connection-only", excelPath: "file.xlsx", queryName: query })
  // Then manually update connections to load to data model (TBD - may need new action)
}
excel_datamodel({ action: "refresh", excelPath: "file.xlsx" })  // One refresh for all
```

**⚠️ Note:** Batch optimization may need a new action like `configure-load-mode-no-refresh` - TBD

## Success Criteria

### Functional Requirements
- [ ] `SetLoadToDataModelAsync` with `autoRefresh=true` actually loads data to Data Model
- [ ] Works on NEWLY imported queries (import → set-load-to-data-model)
- [ ] Works on EXISTING queries (query already exists, just change load mode)
- [ ] Works on CONNECTION-ONLY queries (upgrade to data model loading)
- [ ] Result shows clear distinction between configuration vs. data loading
- [ ] SuggestedNextActions guide users correctly
- [ ] AI agents can complete workflow without trial and error
- [ ] Graceful fallback when Data Model not available
- [ ] Clear error when LoadToWorksheetModel property doesn't exist

### User Experience
- [ ] Single action loads data (no manual refresh needed by default)
- [ ] Clear feedback on what happened vs. what's pending
- [ ] Workflow hints explain next steps accurately
- [ ] Error messages guide users to resolution
- [ ] Batch operation guidance (autoRefresh=false → manual refresh all)

### Testing
- [ ] All existing tests pass
- [ ] New tests verify data actually loads
- [ ] Test on EXISTING query (change load mode)
- [ ] Test on CONNECTION-ONLY query (upgrade to data model)
- [ ] Test batch operations (multiple queries with autoRefresh=false)
- [ ] Integration tests cover full workflow
- [ ] Tests verify Data Model contains data
- [ ] Test verification timing (immediate vs. post-save)

## Migration Guide for Users

### Scenario 1: Import New Query with Immediate Loading (RECOMMENDED)
```typescript
// NEW: One-step import and load to Data Model
excel_powerquery({ 
  action: "import", 
  excelPath: "data.xlsx", 
  queryName: "Sales", 
  mCodeFile: "sales.pq",
  loadToDataModel: true,  // NEW PARAMETER (if we add it)
  autoRefresh: true       // NEW PARAMETER (if we add it)
})
// ✅ Query imported, configured, AND data loaded to Data Model!
```

### Scenario 2: Change Load Mode on Existing Query
```typescript
// Query already exists, just change where it loads
excel_powerquery({ 
  action: "set-load-to-data-model", 
  excelPath: "data.xlsx", 
  queryName: "Sales"
  // autoRefresh defaults to true
})
// ✅ Load mode changed AND data loaded to Data Model!
```

### Scenario 3: Batch Operations (Multiple Queries)
```typescript
// Configure multiple queries WITHOUT loading yet
const queries = ["Sales", "Products", "Customers"];

for (const query of queries) {
  await excel_powerquery({ 
    action: "set-load-to-data-model", 
    excelPath: "data.xlsx", 
    queryName: query,
    autoRefresh: false  // Don't refresh yet
  });
}

// Then load ALL queries at once (more efficient)
await excel_datamodel({ 
  action: "refresh", 
  excelPath: "data.xlsx" 
});
// ✅ All queries loaded to Data Model in one operation!
```

### Scenario 4: Upgrade Connection-Only Query
```typescript
// Query imported with --connection-only originally
excel_powerquery({ 
  action: "set-load-to-data-model", 
  excelPath: "data.xlsx", 
  queryName: "Sales"
})
// ✅ Upgraded from connection-only to Data Model loading!
```

### Before (Current Behavior)
```typescript
// Step 1: Import query
excel_powerquery({ action: "import", excelPath: "data.xlsx", queryName: "Sales", mCodeFile: "sales.pq" })

// Step 2: Configure load mode
excel_powerquery({ action: "set-load-to-data-model", excelPath: "data.xlsx", queryName: "Sales" })
// ⚠️ Data NOT in Data Model yet!

// Step 3: Manually refresh (not documented clearly)
excel_datamodel({ action: "refresh", excelPath: "data.xlsx" })
// Now data is in Data Model
```

### After (New Behavior)
```typescript
// Step 1: Import query
excel_powerquery({ action: "import", excelPath: "data.xlsx", queryName: "Sales", mCodeFile: "sales.pq" })

// Step 2: Configure and load in one step
excel_powerquery({ action: "set-load-to-data-model", excelPath: "data.xlsx", queryName: "Sales" })
// ✅ Data automatically loaded to Data Model!

// Optional: Disable auto-refresh for batch operations
excel_powerquery({ 
  action: "set-load-to-data-model", 
  excelPath: "data.xlsx", 
  queryName: "Sales",
  autoRefresh: false  // Just configure, don't load yet
})
```

## Timeline Estimate

- Phase 1 (Core Logic): 4-6 hours
- Phase 2 (MCP Server): 2-3 hours
- Phase 3 (CLI): 1-2 hours
- Phase 4 (Testing): 3-4 hours
- **Total: 10-15 hours** (1-2 days)

## Related Issues

- None currently, but this improves overall Power Query UX

## Future Enhancements (Out of Scope for This PR)

### Enhancement 1: Add autoRefresh to ImportAsync
```csharp
public async Task<OperationResult> ImportAsync(
    IExcelBatch batch,
    string queryName,
    string mCodeFile,
    PowerQueryPrivacyLevel? privacyLevel = null,
    PowerQueryLoadMode loadMode = PowerQueryLoadMode.ConnectionOnly,  // NEW
    bool autoRefresh = false)  // NEW - default false for backwards compat
{
    // If loadMode = LoadToDataModel AND autoRefresh = true
    //   → Import → Configure → Refresh → Load to Model
    // This would be the ULTIMATE one-step solution
}
```

**Rationale:** Currently users must call `import` then `set-load-to-data-model`. A single `import` with flags would be more intuitive.

**Decision:** NOT in this PR - would require significant API changes. Consider for future enhancement.

### Enhancement 2: Batch Operation Helper
```csharp
public async Task<OperationResult> SetLoadToDataModelBatchAsync(
    IExcelBatch batch,
    List<string> queryNames,
    PowerQueryPrivacyLevel? privacyLevel = null)
{
    // Configure all queries
    // Then refresh Data Model ONCE
    // More efficient than individual refreshes
}
```

**Decision:** NOT in this PR - users can achieve this with `autoRefresh=false` + manual `excel_datamodel refresh`.

### Enhancement 3: Smart Load Mode Detection
When calling `set-load-to-data-model` on a query already loading to worksheet:
- **Option A:** Error - require explicit user choice
- **Option B:** Auto-upgrade to LoadToBoth (RECOMMENDED for UX)
- **Option C:** Add parameter `allowUpgradeToB oth: bool`

**Decision for This PR:** Implement Option B (auto-upgrade to LoadToBoth with clear messaging).

## References

- Issue #42: https://github.com/sbroenne/mcp-server-excel/issues/42
- Excel COM API: `Workbook.Model.ModelTables`
- Power Query COM: `WorkbookQuery.LoadToWorksheetModel`
