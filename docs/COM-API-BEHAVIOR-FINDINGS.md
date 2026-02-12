# Excel COM API Behavior Findings

> **Summary of diagnostic test findings from empirical testing of raw Excel COM API behavior**

This document captures the key discoveries made through diagnostic tests that use raw COM API calls to understand Excel's actual behavior, without our abstractions.

## Test Suites

- **PowerQueryComApiBehaviorTests**: 12 scenarios testing Power Query M code operations
- **DataModelComApiBehaviorTests**: 11 scenarios testing Data Model (Power Pivot) operations

## Critical Finding: Query Deletion Orphans Tables

### Power Query (Scenario 4)

**FINDING: Deleting a Power Query does NOT delete the associated table (ListObject)**

```csharp
// Delete the query
query.Delete();

// Table SURVIVES - query deletion does NOT cascade to table
// Tables after delete: 1  (same as before)
// Orphaned table name: TestQuery
// Data rows still accessible: 3
```

**Implication:** Cleanup code that removes orphaned tables after query deletion is **JUSTIFIED** and necessary. This is not "cargo cult" code - it addresses real orphan behavior.

### Data Model (Scenario 8)

**FINDING: Deleting a Power Query does NOT delete the associated Data Model table**

```csharp
// Delete query that loaded to Data Model
query.Delete();

// Model table SURVIVES!
// Queries after delete: 0
// Model tables after delete: 1  (orphaned)
```

**Implication:** Same as Power Query - Data Model tables become orphaned when their source query is deleted. Cleanup code is required.

---

## Power Query API Findings

### Scenario 1: Query Creation and Loading

- `Queries.Add(name, formula)` creates query definition only
- Query alone does NOT create a table
- To load to worksheet, must use `QueryTables.Add()` with connection string:

  ```
  OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={QueryName}
  ```

- `QueryTable.Refresh(false)` required for synchronous data load

### Scenario 2-3: Formula Updates

**FINDING: Updating formula does NOT automatically update loaded data**

```csharp
query.Formula = ModifiedQuery;  // Updates M code
// Table still shows OLD data
// Must call qt.Refresh(false) to update
```

**Implication:** Refresh operation is required after formula updates.

### Scenario 6: Connection-Only Mode

**FINDING: There is NO "connection-only" flag on WorkbookQuery**

- Excel UI has this option, but COM API does not expose it
- `WorkbookQuery` only has these members:
  - **Properties:** `Application`, `Creator`, `Description`, `Formula`, `Name`, `Parent`
  - **Methods:** `Delete()`, `Refresh()`
  - **NO `IsConnectionOnly`, `LoadDestination`, or similar property exists**
- "Connection-only" is the **ABSENCE** of load destinations (no ListObject, no Data Model connection)
- To convert loaded query to connection-only:
  1. Find and remove worksheet tables (ListObjects with matching QueryTable)
  2. Find and remove Data Model connections (connections with "Query - {name}" pattern)
  3. Keep the query in `Workbook.Queries` collection
  
**Implication:** Connection-only must be implemented by removing ALL load destinations.

### Scenario 12-14: Unload to Connection-Only (Bug Discovery)

**CRITICAL BUG FOUND: Unload only removes worksheet tables, NOT Data Model connections**

```csharp
// Current Unload implementation (INCOMPLETE):
foreach (ListObject in worksheet.ListObjects)
{
    if (QueryTable.Connection.Contains(queryName))
        listObject.Unlist();
}
// BUG: Never checks/removes Data Model connections!
```

**Scenarios Tested:**

| Scenario | Initial State | After Current Unload | Correct Result |
|----------|--------------|---------------------|----------------|
| 12: Data Model Only | Query → Data Model | Data Model connection REMAINS | Should remove connection |
| 13: Both Destinations | Query → Worksheet + Data Model | Worksheet removed, Data Model REMAINS | Should remove both |
| 14: Proper Implementation | Query → Both | BOTH removed | ✅ Correct |

**Proper Connection-Only Implementation:**

```csharp
// Step 1: Remove worksheet tables (existing behavior)
foreach (ListObject in worksheet.ListObjects)
{
    if (QueryTable.Connection.Contains(queryName))
        listObject.Unlist();
}

// Step 2: Remove Data Model connections (MISSING!)
foreach (Connection in workbook.Connections)
{
    if (connection.Name == $"Query - {queryName}")
        connection.Delete();
}

// Query remains in Workbook.Queries = connection-only
```

**Implication:** `Unload` method in `PowerQueryCommands.Lifecycle.cs` needs to also remove Data Model connections.

---

### Scenario 15: Data Model State Blocks Formula Updates (UNRESOLVED)

**FINDING: WorkbookQuery.Formula becomes read-only in certain workbook states**

**Error Code:** `0x800A03EC` (-2146827284) - "Application-defined or object-defined error"

**Symptoms:**
- `query.Formula = mCode` fails with 0x800A03EC on some workbooks
- Error affects ALL queries in the workbook, not just Data Model queries
- Error is workbook-specific - same queries work fine in fresh workbooks

**Attempted Solutions (Did NOT Work):**
1. **Save-and-retry**: Saving workbook before retry did NOT clear the error state
2. **Polly retry with backoff**: Error is NOT transient - retries don't help
3. **File location**: Copying file from OneDrive to local drive did NOT help

**Current Status:** Root cause unknown. The error appears to be related to internal Excel state that cannot be cleared programmatically via COM automation. Workaround: manually save and reopen the workbook in Excel UI.

**See:** GitHub Issue #323

---

### Scenario 16: Query Renaming

**FINDING: Renaming a query breaks the connection to the table**

```csharp
query.Name = "RenamedQuery";
// Table's QueryTable.Connection still references old name!
// Refresh FAILS because connection can't find "OldName"
```

**Implication:** Renaming requires fixing the connection string or recreating the table.

---

## Data Model API Findings

### ModelMeasures.Add() Signature

**CORRECT Signature:**

```csharp
measures.Add(
    "MeasureName",           // Required: String
    modelTable,              // Required: ModelTable object (NOT string!)
    "DAX Formula",           // Required: String
    model.ModelFormatGeneral // Required: ModelFormat* object (NOT null!)
    // Description            // Optional: String
);
```

**Key Points:**

- `AssociatedTable` parameter is a **ModelTable object**, not a table name string
- `FormatInformation` is **REQUIRED** - use `model.ModelFormatGeneral` for default format
- Available format types: ModelFormatGeneral, ModelFormatCurrency, ModelFormatDate, ModelFormatDecimalNumber, ModelFormatWholeNumber, ModelFormatPercentageNumber, ModelFormatScientificNumber, ModelFormatBoolean

### Scenario 3-5: Measure Lifecycle

- Measures can be added, updated (by changing Formula property), and deleted
- `measure.Delete()` removes the measure from the model
- Measure can reference any table in the model, not just the associated table

### Scenario 6-7: Relationships

- Relationships require two tables with compatible columns
- `ModelRelationships.Add()` creates relationship between columns
- Relationships can be deleted individually

### Scenario 9: Orphaned Measures

**FINDING: Measures referencing deleted tables become invalid**

When a table is removed from the Data Model, measures that reference it remain but become #ERROR.

---

## COM API Quirks

### 1-Based Indexing

All Excel COM collections use 1-based indexing:

```csharp
collection.Item(1)  // First item, NOT collection.Item(0)
```

### Numeric Property Types

All Excel COM numeric properties return `double`, not `int`:

```csharp
// WRONG: Runtime error
int position = field.Position;

// CORRECT: Explicit conversion required
int position = Convert.ToInt32(field.Position);
```

### Error Handling

COM exceptions provide HResult codes for error identification:

```csharp
catch (COMException ex) when (ex.HResult == -2147417851)  // RPC_E_SERVERCALL_RETRYLATER
```

---

## Implications for Production Code

### 1. Cleanup Code is Necessary

The discovery that `query.Delete()` leaves tables orphaned validates our cleanup code:

- Power Query delete should also delete/unlist associated table
- Data Model cleanup should remove orphaned model tables

### 2. Refresh is Required After Updates

Formula changes do NOT auto-refresh:

- After updating M code, must call refresh
- Refresh should be synchronous (`Refresh(false)`)

### 3. Connection Strings Must Be Managed

When renaming queries, connection strings break:

- Must update connection string or recreate table
- Consider using query name as table name for consistency

### 4. ModelMeasures.Add Requires Proper Parameters

The API signature differs from what might be assumed:

- TableName → ModelTable object
- FormatInformation → Required ModelFormat* object (not null)

### 5. Unload Must Remove ALL Load Destinations (BUG FIX NEEDED)

Current `Unload` method only removes worksheet tables (ListObjects):

- **Missing:** Removal of Data Model connections
- Queries loaded only to Data Model are NOT affected by current Unload
- Queries loaded to BOTH destinations only have worksheet table removed

**Fix Required in `PowerQueryCommands.Lifecycle.cs`:**

```csharp
// After unlisting worksheet tables, also remove Data Model connections:
foreach (Connection in workbook.Connections)
{
    if (connection.Name == $"Query - {queryName}")
        connection.Delete();
}
```

### 6. CUBEVALUE Formula Limitation in Automation Mode (Issue #313)

**CRITICAL FINDING: CUBEVALUE formulas do NOT work in Excel COM automation mode**

**Symptoms:**
- CUBEVALUE returns #N/A when Excel is hidden (automation mode)
- CUBEVALUE returns #VALUE! when Excel is visible
- All syntax variations fail, including:
  - `=CUBEVALUE("ThisWorkbookDataModel","[Measures].[TotalAmount]")`
  - `=CUBEVALUE("ThisWorkbookDataModel","Query[TotalAmount]")`
  - `=CUBEVALUE("ThisWorkbookDataModel","[Query].[Measures].[TotalAmount]")`

**What Works (via COM):**
- `Workbook.Model` - full access to Data Model object
- `Model.ModelTables` - listing/accessing tables
- `Model.ModelMeasures.Add()` - creating DAX measures
- `Model.ModelRelationships` - creating/managing relationships
- `Model.Refresh()` - refreshing the model
- `Model.DataModelConnection` - returns "ThisWorkbookDataModel" (Type 7 connection)
- `Model.CreateModelWorkbookConnection()` - creates additional connections
- Power Query loading to Data Model via `CreateModelConnection=true`

**What Fails (via COM):**
- **CUBEVALUE worksheet function** - cannot resolve measures even with correct connection name
- **CUBEMEMBER worksheet function** - also fails with #N/A
- Calculate methods succeed but don't resolve CUBEVALUE
- Error codes:
  - -2146826245 = #N/A (member doesn't exist in cube or syntax incorrect)
  - -2146826246 = #VALUE! (invalid tuple element)
  - 0x800AC472 = Excel busy (Calculate blocked in hidden mode)

**Root Cause (Confirmed by Microsoft Documentation):**
Per Microsoft's [Client Architecture Requirements for Analysis Services Development](https://learn.microsoft.com/en-us/analysis-services/multidimensional-models/olap-physical/client-architecture-requirements-for-analysis-services-development):

> **"Power Pivot for Excel and SQL Server Data Tools are the only client environments that are supported for creating and querying in-memory databases that use SharePoint or Tabular mode."**

The embedded VertiPaq/Analysis Services engine uses **INPROC transport** for in-process communication only. COM automation (external process) is NOT a supported client environment for querying the Data Model via CUBE functions. The Model object and its collections (ModelTables, ModelMeasures, ModelRelationships) work because they use Excel's internal object model, not the OLAP query layer.

**Workaround - Use PivotTables:**
To evaluate DAX measures programmatically, use PivotTable-based approaches:
1. Create a PivotTable connected to the Data Model (`pivottable action: 'CreateFromDataModel'`)
2. Add the measure as a value field (`pivottable action: 'AddValueField'`)
3. Read the PivotTable values (`pivottable action: 'GetData'`)

**Microsoft Docs References:**
- CUBEVALUE returns #N/A if "member doesn't exist in the cube"
- CUBEVALUE returns #VALUE! if "at least one element within the tuple is invalid"
- [CUBEVALUE function documentation](https://learn.microsoft.com/en-us/office/client-developer/excel/cubevalue-function)

**Test Evidence:** Scenario12 and Scenario13 in `DataModelComApiBehaviorTests.cs`

---

### Scenario 14-15: DAX EVALUATE Query Execution (Issue #356)

**CRITICAL FINDING: DAX EVALUATE queries CAN be executed via COM automation**

Despite CUBEVALUE/CUBEMEMBER worksheet functions failing (see above), DAX EVALUATE queries **DO WORK** through alternative COM APIs:

**What FAILED (Scenario 14):**
- `ListObjects.Add(xlSrcModel, ...)` - Returns "Value does not fall within expected range"
- `Connections.Add2(..., xlCmdDAX, ...)` - Same error

**What WORKS (Scenario 15):**

1. **Model.CreateModelWorkbookConnection + xlCmdDAX:**
   ```csharp
   // Create a model connection for a table
   dynamic modelWbConn = model.CreateModelWorkbookConnection("TableName");
   dynamic modelConnection = modelWbConn.ModelConnection;
   
   // Change command type to xlCmdDAX (8)
   modelConnection.CommandType = 8;  // xlCmdDAX
   modelConnection.CommandText = "EVALUATE 'TableName'";
   
   // Refresh executes the DAX query
   modelWbConn.Refresh();  // ✅ SUCCESS!
   ```

2. **ModelConnection.ADOConnection.Execute (BEST APPROACH):**
   ```csharp
   // Get DataModelConnection and its ModelConnection
   dynamic dataModelConn = model.DataModelConnection;
   dynamic modelConn = dataModelConn.ModelConnection;
   
   // Get ADO connection - this is a live MSOLAP connection!
   dynamic adoConnection = modelConn.ADOConnection;
   // ConnectionString: Provider=MSOLAP.8;...Data Source=$Embedded$...
   
   // Execute DAX EVALUATE query directly
   dynamic recordset = adoConnection.Execute("EVALUATE 'TableName'");
   
   // Read results from recordset
   while (!recordset.EOF)
   {
       // recordset.Fields.Item(0).Value, etc.
       recordset.MoveNext();
   }
   ```

**ADOConnection Details:**
- Provider: `MSOLAP.8` (Analysis Services OLE DB Provider)
- Data Source: `$Embedded$` (in-process connection to Excel's Data Model)
- Returns: Standard ADO Recordset with DAX query results
- Fields include fully qualified column names: `TableName[ColumnName]`

**This is DIFFERENT from CUBEVALUE:** 
- CUBEVALUE uses the worksheet function evaluation layer → blocked by INPROC transport limitation
- ADOConnection.Execute uses the MSOLAP provider directly → works!

**Implication for Issue #356:**
An `evaluate` action can be added to `datamodel` tool using the ADOConnection approach:
1. Get `Workbook.Model.DataModelConnection.ModelConnection.ADOConnection`
2. Execute DAX EVALUATE query via `adoConnection.Execute(daxQuery)`
3. Convert ADO Recordset to JSON result

**Test Evidence:** Scenario14 and Scenario15 in `DataModelComApiBehaviorTests.cs`

---

### DAX-Backed Excel Tables (Scenario 16)

**FINDING: Excel Tables (ListObjects) can be backed by DAX queries!**

This extends the Scenario 15 discovery - not only can DAX queries be executed, but the results can be materialized as Excel Tables that automatically update when the Data Model changes.

**What WORKS (Scenario 16):**

```csharp
// 1. Create a model workbook connection for a table
dynamic modelWbConn = model.CreateModelWorkbookConnection("TableName");
dynamic modelConnection = modelWbConn.ModelConnection;

// 2. Configure for DAX EVALUATE query
modelConnection.CommandType = 8;  // xlCmdDAX
modelConnection.CommandText = @"
    EVALUATE 
    SUMMARIZECOLUMNS(
        'Query'[Region],
        ""TotalAmount"", SUM('Query'[Amount]),
        ""TotalQty"", SUM('Query'[Qty])
    )";

// 3. Refresh to execute the DAX query
modelWbConn.Refresh();

// 4. Create Excel Table (ListObject) backed by the DAX query!
dynamic listObjects = targetSheet.ListObjects;
dynamic listObject = listObjects.Add(
    4,              // xlSrcModel = 4 (PowerPivot Model source)
    modelWbConn,    // The ModelWorkbookConnection with DAX
    true,           // HasHeaders
    1,              // xlYes = 1
    destRange       // Target range
);

// 5. Refresh the table to populate data
listObject.Refresh();

// Result: Excel Table with DAX-aggregated data!
// Headers: Region, TotalAmount, TotalQty
// Data: Aggregated rows from the DAX SUMMARIZECOLUMNS query
```

**Key Constants:**
- `xlSrcModel = 4` - ListObject source type for PowerPivot/Data Model
- `xlCmdDAX = 8` - Command type for DAX queries

**Capabilities Unlocked:**
1. **DAX Aggregation Tables** - Create summary tables with SUMMARIZE, SUMMARIZECOLUMNS, TOPN, etc.
2. **Filtered Data Tables** - Use FILTER, CALCULATETABLE to create subsets
3. **Cross-Table Analysis** - Join/aggregate data across multiple Data Model tables
4. **Auto-Refreshable** - Tables update when underlying Data Model refreshes

**Potential New Features for Issue #356:**
- Add `create-table-from-dax` action to `table` tool
- Allow users to create Excel Tables populated by arbitrary DAX EVALUATE queries
- Tables stay linked to Data Model and can be refreshed

**Test Evidence:** Scenario16 in `DataModelComApiBehaviorTests.cs`

---

## Test Execution

Run diagnostic tests on demand:

```powershell
# Power Query diagnostics
dotnet test --filter "FullyQualifiedName~PowerQueryComApiBehaviorTests"

# Data Model diagnostics  
dotnet test --filter "FullyQualifiedName~DataModelComApiBehaviorTests"
```

These tests are marked `[Trait("RunType", "OnDemand")]` and excluded from regular test runs due to their longer execution time.

---

## References

- Microsoft VBA Documentation: <https://learn.microsoft.com/en-us/office/vba/api/overview/excel>
- NetOffice (C# COM wrappers): <https://github.com/NetOfficeFw/NetOffice>
- Test Files (in `tests/ExcelMcp.Diagnostics.Tests/`):
  - `Integration/Diagnostics/PowerQueryComApiBehaviorTests.cs`
  - `Integration/Diagnostics/DataModelComApiBehaviorTests.cs`
  - `Integration/Diagnostics/PivotTableRefreshBehaviorTests.cs`

**NOTE: Diagnostics tests are excluded from CI. Run manually with:**
```powershell
dotnet test tests/ExcelMcp.Diagnostics.Tests/ --filter "RunType=OnDemand&Layer=Diagnostics"
```
