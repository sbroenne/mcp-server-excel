# Excel Tables (ListObjects) Implementation Guide

## Security and Robustness Requirements

This guide documents the **CRITICAL** security and robustness requirements for implementing Excel Tables (ListObjects) functionality. These requirements were identified during senior architect review and **MUST** be followed.

**Related Issue:** #24 (Excel Tables Implementation)  
**Severity:** 🔴 **CRITICAL**  
**Priority:** **P0 - Blocker**

---

## Table of Contents
1. [Security Requirements](#security-requirements)
2. [Robustness Requirements](#robustness-requirements)
3. [Implementation Examples](#implementation-examples)
4. [Testing Requirements](#testing-requirements)
5. [Security Checklist](#security-checklist)

---

## Security Requirements

### 1. Path Traversal Prevention (CRITICAL)

**✅ ALWAYS use `PathValidator.ValidateExistingFile()` for file paths**

```csharp
using Sbroenne.ExcelMcp.Core.Security;

public TableListResult List(string filePath)
{
    // CRITICAL: Validate path to prevent traversal attacks
    // This MUST be the FIRST operation before any file access
    filePath = PathValidator.ValidateExistingFile(filePath, nameof(filePath));
    
    var result = new TableListResult 
    { 
        FilePath = filePath, 
        Action = "list-tables" 
    };
    
    // ... rest of implementation
}
```

**Why:** Prevents attacks like `../../etc/passwd`, `C:\Windows\System32\config\SAM`

**Affected Methods:** ALL methods that accept `filePath` parameter

---

### 2. Table Name Validation (CRITICAL)

**✅ ALWAYS use `TableNameValidator.ValidateTableName()` for table names**

```csharp
using Sbroenne.ExcelMcp.Core.Security;

public OperationResult Create(string filePath, string sheetName, string tableName, string range)
{
    // Validate file path (FIRST!)
    filePath = PathValidator.ValidateExistingFile(filePath, nameof(filePath));
    
    // Validate table name (BEFORE using in Excel)
    tableName = TableNameValidator.ValidateTableName(tableName, nameof(tableName));
    
    var result = new OperationResult 
    { 
        FilePath = filePath, 
        Action = "create-table" 
    };
    
    // ... rest of implementation
}
```

**ValidationRules Enforced:**
- ❌ No null/empty/whitespace
- ❌ No more than 255 characters
- ❌ No spaces (use underscores instead)
- ❌ Must start with letter or underscore
- ❌ Only letters, digits, underscores, periods allowed
- ❌ No reserved names (Print_Area, Print_Titles, _FilterDatabase, etc.)
- ❌ No cell reference patterns (A1, R1C1, etc.)

**Prevents:** Name injection, formula injection, Excel formula execution

**Affected Methods:** `Create()`, `Rename()`

---

### 3. Range Validation (RECOMMENDED)

**✅ Use `RangeValidator.ValidateRange()` to prevent DoS attacks**

```csharp
using Sbroenne.ExcelMcp.Core.Security;

public OperationResult Create(string filePath, string sheetName, string tableName, string range)
{
    filePath = PathValidator.ValidateExistingFile(filePath, nameof(filePath));
    tableName = TableNameValidator.ValidateTableName(tableName, nameof(tableName));
    
    // Validate range address format
    range = RangeValidator.ValidateRangeAddress(range, nameof(range));
    
    var result = new OperationResult { FilePath = filePath, Action = "create-table" };
    
    return ExcelHelper.WithExcel(filePath, save: true, (excel, workbook) =>
    {
        dynamic? sheet = null;
        dynamic? rangeObj = null;
        try
        {
            sheet = ComUtilities.FindSheet(workbook, sheetName);
            if (sheet == null)
            {
                result.Success = false;
                result.ErrorMessage = $"Sheet '{sheetName}' not found";
                return 1;
            }
            
            // Get range object
            try
            {
                rangeObj = sheet.Range[range];
            }
            catch (COMException ex) when (ex.HResult == -2147352567) // VBA_E_ILLEGALFUNCALL
            {
                result.Success = false;
                result.ErrorMessage = $"Invalid range '{range}': Range does not exist or is malformed";
                return 1;
            }
            
            // CRITICAL: Validate range size to prevent DoS
            RangeValidator.ValidateRange(rangeObj, parameterName: nameof(range));
            
            // Create table
            dynamic? listObjects = null;
            dynamic? newTable = null;
            try
            {
                listObjects = sheet.ListObjects;
                
                // Use Add() method with correct parameters
                // Note: HasHeaders = 1 means Yes, 0 means No
                newTable = listObjects.Add(
                    SourceType: 1,  // xlSrcRange
                    Source: rangeObj,
                    LinkSource: false,
                    XlListObjectHasHeaders: 1,  // xlYes - has headers
                    Destination: Type.Missing
                );
                
                newTable.Name = tableName;
                
                result.Success = true;
                result.ErrorMessage = null;
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref newTable);
                ComUtilities.Release(ref listObjects);
            }
        }
        finally
        {
            ComUtilities.Release(ref rangeObj);
            ComUtilities.Release(ref sheet);
        }
    });
}
```

**Prevents:** DoS attacks from creating tables with millions/billions of cells

**Default Limit:** 1,000,000 cells (1000 rows × 1000 columns)

---

## Robustness Requirements

### 4. Null Reference Prevention - HeaderRowRange/DataBodyRange (CRITICAL)

**✅ ALWAYS check for null before accessing `HeaderRowRange` and `DataBodyRange`**

```csharp
public TableListResult List(string filePath)
{
    filePath = PathValidator.ValidateExistingFile(filePath, nameof(filePath));
    
    var result = new TableListResult { FilePath = filePath, Action = "list-tables" };
    
    return ExcelHelper.WithExcel(filePath, save: false, (excel, workbook) =>
    {
        dynamic? sheets = null;
        try
        {
            sheets = workbook.Worksheets;
            
            for (int i = 1; i <= sheets.Count; i++)
            {
                dynamic? sheet = null;
                dynamic? listObjects = null;
                try
                {
                    sheet = sheets.Item(i);
                    listObjects = sheet.ListObjects;
                    
                    for (int j = 1; j <= listObjects.Count; j++)
                    {
                        dynamic? listObject = null;
                        dynamic? headerRowRange = null;
                        dynamic? dataBodyRange = null;
                        try
                        {
                            listObject = listObjects.Item(j);
                            
                            // CRITICAL: HeaderRowRange can be NULL
                            headerRowRange = listObject.ShowHeaders ? listObject.HeaderRowRange : null;
                            
                            // CRITICAL: DataBodyRange can be NULL
                            dataBodyRange = listObject.DataBodyRange;
                            
                            var tableInfo = new TableInfo
                            {
                                Name = listObject.Name,
                                SheetName = sheet.Name,
                                
                                // Safe access with null-conditional operator
                                DataRowCount = dataBodyRange?.Rows.Count ?? 0,
                                ColumnCount = headerRowRange?.Columns.Count ?? 0,
                                
                                ShowHeaders = listObject.ShowHeaders,
                                ShowTotals = listObject.ShowTotals
                            };
                            
                            // Get column names only if headers exist
                            if (headerRowRange != null && listObject.ShowHeaders)
                            {
                                dynamic? columns = null;
                                try
                                {
                                    columns = listObject.ListColumns;
                                    
                                    for (int k = 1; k <= columns.Count; k++)
                                    {
                                        dynamic? column = null;
                                        try
                                        {
                                            column = columns.Item(k);
                                            tableInfo.ColumnNames.Add(column.Name);
                                        }
                                        finally
                                        {
                                            ComUtilities.Release(ref column);
                                        }
                                    }
                                }
                                finally
                                {
                                    ComUtilities.Release(ref columns);
                                }
                            }
                            
                            result.Tables.Add(tableInfo);
                        }
                        finally
                        {
                            ComUtilities.Release(ref dataBodyRange);
                            ComUtilities.Release(ref headerRowRange);
                            ComUtilities.Release(ref listObject);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref listObjects);
                    ComUtilities.Release(ref sheet);
                }
            }
            
            result.Success = true;
            return 0;
        }
        finally
        {
            ComUtilities.Release(ref sheets);
        }
    });
}
```

**Why:** Tables with no data rows have `DataBodyRange = null`. Tables with `ShowHeaders = false` have `HeaderRowRange = null`.

**Failure Mode:** `NullReferenceException` crashes the application

**Affected Methods:** `List()`, `GetInfo()`

---

### 5. COM Cleanup After Unlist() (CRITICAL)

**✅ ALWAYS release COM objects after calling `Unlist()`**

```csharp
public OperationResult Delete(string filePath, string tableName)
{
    filePath = PathValidator.ValidateExistingFile(filePath, nameof(filePath));
    tableName = TableNameValidator.ValidateTableName(tableName, nameof(tableName));
    
    var result = new OperationResult { FilePath = filePath, Action = "delete-table" };
    
    return ExcelHelper.WithExcel(filePath, save: true, (excel, workbook) =>
    {
        dynamic? sheets = null;
        try
        {
            sheets = workbook.Worksheets;
            bool found = false;
            
            for (int i = 1; i <= sheets.Count; i++)
            {
                dynamic? sheet = null;
                dynamic? listObjects = null;
                try
                {
                    sheet = sheets.Item(i);
                    listObjects = sheet.ListObjects;
                    
                    for (int j = 1; j <= listObjects.Count; j++)
                    {
                        dynamic? listObject = null;
                        try
                        {
                            listObject = listObjects.Item(j);
                            
                            if (listObject.Name == tableName)
                            {
                                // Step 1: Unlist (converts table to range, removes table formatting)
                                listObject.Unlist();
                                
                                // Step 2: CRITICAL - Release COM reference immediately
                                // The COM object is now invalid after Unlist()
                                ComUtilities.Release(ref listObject);
                                
                                // Step 3: Explicit null to prevent use-after-free
                                listObject = null;
                                
                                found = true;
                                break;
                            }
                        }
                        finally
                        {
                            // Final cleanup (handles case where Unlist() not called)
                            if (listObject != null)
                            {
                                ComUtilities.Release(ref listObject);
                            }
                        }
                    }
                    
                    if (found) break;
                }
                finally
                {
                    ComUtilities.Release(ref listObjects);
                    ComUtilities.Release(ref sheet);
                }
            }
            
            if (found)
            {
                result.Success = true;
                result.ErrorMessage = null;
            }
            else
            {
                result.Success = false;
                result.ErrorMessage = $"Table '{tableName}' not found";
            }
            
            return found ? 0 : 1;
        }
        finally
        {
            ComUtilities.Release(ref sheets);
        }
    });
}
```

**Why:** `Unlist()` removes the table but the COM reference still points to the deleted object. Not releasing causes memory leaks and potential use-after-free bugs.

**Failure Mode:** Excel.exe process leaks, memory leaks

**Affected Methods:** `Delete()`

---

## Testing Requirements

### Unit Tests (Security Validation)

```csharp
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
public class TableCommandsSecurityTests
{
    [Fact]
    public void List_WithPathTraversal_ThrowsArgumentException()
    {
        var commands = new TableCommands();
        
        // Path traversal attempts should be blocked by PathValidator
        Assert.Throws<ArgumentException>(() => 
            commands.List("../../etc/passwd"));
    }
    
    [Theory]
    [InlineData("Invalid Name")]         // Space
    [InlineData("123Start")]              // Starts with number
    [InlineData("Print_Area")]            // Reserved
    [InlineData("A1")]                    // Cell reference
    [InlineData("My@Table")]              // Invalid character
    public void Create_WithInvalidTableName_ThrowsArgumentException(string invalidName)
    {
        var commands = new TableCommands();
        var testFile = "test.xlsx";
        
        Assert.Throws<ArgumentException>(() => 
            commands.Create(testFile, "Sheet1", invalidName, "A1:B2"));
    }
}
```

### Integration Tests (Null Handling)

```csharp
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("RequiresExcel", "true")]
public class TableCommandsRobustnessTests
{
    [Fact]
    public void GetInfo_HeaderOnlyTable_ReturnsZeroDataRows()
    {
        // Create table with header but no data
        var commands = new TableCommands();
        var testFile = CreateTestFileWithHeaderOnlyTable();
        
        var result = commands.GetInfo(testFile, "HeaderOnlyTable");
        
        Assert.True(result.Success);
        Assert.Equal(0, result.Table.DataRowCount);  // Should not crash
    }
    
    [Fact]
    public void List_TableWithNoHeaders_ReturnsCorrectInfo()
    {
        var commands = new TableCommands();
        var testFile = CreateTestFileWithNoHeaderTable();
        
        var result = commands.List(testFile);
        
        Assert.True(result.Success);
        var table = result.Tables.First(t => t.Name == "NoHeaderTable");
        Assert.False(table.ShowHeaders);
        Assert.Empty(table.ColumnNames);  // Should not crash
    }
}
```

### OnDemand Tests (Process Cleanup)

```csharp
[Trait("RunType", "OnDemand")]
[Trait("Speed", "Slow")]
public class TableCommandsProcessCleanupTests
{
    [Fact]
    public void Delete_Table_NoProcessLeak()
    {
        // CRITICAL: Run this test before committing pool-related changes
        
        var commands = new TableCommands();
        var testFile = CreateTestFileWithTable();
        
        // Get initial Excel process count
        int initialCount = Process.GetProcessesByName("EXCEL").Length;
        
        // Create and delete table
        commands.Create(testFile, "Sheet1", "TempTable", "A1:B10");
        commands.Delete(testFile, "TempTable");
        
        // Force garbage collection
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        
        // Wait for COM cleanup
        Thread.Sleep(1000);
        
        // Verify no Excel.exe leak
        int finalCount = Process.GetProcessesByName("EXCEL").Length;
        Assert.Equal(initialCount, finalCount);
    }
}
```

---

## Security Checklist

Before merging TableCommands implementation, verify:

### Path Security
- [ ] All methods use `PathValidator.ValidateExistingFile()` as **FIRST** operation
- [ ] PathValidator is called before any `File.Exists()`, `ExcelHelper.WithExcel()`, or file operations
- [ ] Unit tests verify path traversal attempts are blocked

### Table Name Security
- [ ] All methods use `TableNameValidator.ValidateTableName()` before using names
- [ ] Validation happens before passing names to Excel COM
- [ ] Unit tests cover all validation rules (spaces, reserved names, cell references, etc.)

### Range Security
- [ ] `RangeValidator.ValidateRange()` is used before processing ranges
- [ ] `RangeValidator.ValidateRangeAddress()` validates address strings
- [ ] Unit tests verify DoS prevention for oversized ranges

### Null Handling
- [ ] All `HeaderRowRange` accesses check for null
- [ ] All `DataBodyRange` accesses check for null
- [ ] Integration tests cover header-only and no-header tables

### COM Cleanup
- [ ] `ComUtilities.Release()` called after `Unlist()`
- [ ] Explicit null assignment after release
- [ ] OnDemand tests verify no Excel.exe process leaks

### Error Handling
- [ ] `COMException` wrapped with clear error messages
- [ ] Invalid range errors caught and reported clearly
- [ ] All error paths release COM objects in `finally` blocks

---

## Additional Resources

- **PathValidator Source:** `src/ExcelMcp.Core/Security/PathValidator.cs`
- **TableNameValidator Source:** `src/ExcelMcp.Core/Security/TableNameValidator.cs`
- **RangeValidator Source:** `src/ExcelMcp.Core/Security/RangeValidator.cs`
- **ComUtilities Source:** `src/ExcelMcp.Core/ComInterop/ComUtilities.cs`
- **Existing Commands Pattern:** `src/ExcelMcp.Core/Commands/SheetCommands.cs`

---

**⚠️ CRITICAL REMINDER:** All 4 critical issues MUST be addressed before Issue #24 implementation begins. No exceptions.
