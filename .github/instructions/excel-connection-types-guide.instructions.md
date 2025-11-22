---
applyTo: "src/ExcelMcp.Core/Commands/ConnectionCommands.cs,src/ExcelMcp.Core/Connections/**/*.cs,tests/**/ConnectionCommands*.cs,tests/**/ConnectionTestHelper.cs"
---

# Excel Connection Types - Essential Patterns

> **Critical patterns for working with Excel connections**

## Connection Type Enum

[Official docs](https://learn.microsoft.com/en-us/office/vba/api/excel.xlconnectiontype): Types 1-9 (OLEDB, ODBC, TEXT, WEB, XMLMAP, DATAFEED, MODEL, WORKSHEET, NOSOURCE)

## COM API Behavior (CORRECTED 2025-01-30)

### ✅ Connections.Add2() Works for OLEDB/ODBC

**Previous Documentation (INCORRECT):** Claimed OLEDB/ODBC connection creation failed via COM API.

**Current Status (VERIFIED):** OLEDB and ODBC connections **can be created** using `Connections.Add2()` method.

**Test Evidence:**
- ✅ OLEDB with SQL Server LocalDB: **SUCCESS**
- ✅ OLEDB with Microsoft Access: **SUCCESS**  
- ✅ ODBC connections: **SUCCESS**
- ✅ TEXT connections: **SUCCESS** (as before)

**Key Requirement:** Must use `Connections.Add2()` (current method), not deprecated `Connections.Add()`.

**Implementation:**
```csharp
dynamic connections = workbook.Connections;
dynamic newConn = connections.Add2(
    Name: connectionName,
    Description: description ?? "",
    ConnectionString: connectionString,
    CommandText: "",
    lCmdtype: Type.Missing,            // Let Excel auto-detect
    CreateModelConnection: false,       // Don't create Data Model connection
    ImportRelationships: false          // Don't import relationships
);
```

### ✅ Use TEXT Connections for Testing

```csharp
// TEXT connections work reliably
string connectionString = $"TEXT;{csvFilePath}";
dynamic conn = connections.Add(
    Name: "TestText",
    Description: "Test",
    ConnectionString: connectionString,
    CommandText: ""
);  // ✅ WORKS
```

## Type 3/4 Handling Pattern

**Issue:** TEXT connections created with `"TEXT;path"` return type 4 (WEB) instead of 3 (TEXT).

**Solution:** Handle both types interchangeably:


## Connection String Formats

```csharp
// OLEDB
"Provider=SQLOLEDB;Data Source=server;Initial Catalog=db;"

// ODBC
"DSN=MyDataSource;UID=username;PWD=password;"

// TEXT
"TEXT;C:\\path\\to\\file.csv"

// WEB
"URL;https://example.com/data.xml"

// Power Query
"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=QueryName"
```

## LoadTo Operation Patterns

**OLEDB connections are the PRIMARY use case for LoadTo:**

```csharp
// ✅ CORRECT: OLEDB connections support QueryTables.Add() pattern
string connectionString = "Provider=SQLOLEDB;Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=tempdb;Integrated Security=SSPI;";
ConnectionTestHelper.CreateOleDbConnection(testFile, connectionName, connectionString);
var result = _commands.LoadTo(batch, connectionName, "Sheet1");  // ✅ WORKS

// ❌ WRONG: TEXT connections cannot be loaded via QueryTables.Add()
ConnectionTestHelper.CreateTextFileConnection(testFile, connectionName, csvFile);
var result = _commands.LoadTo(batch, connectionName, "Sheet1");  // ❌ FAILS with E_INVALIDARG
```

**Why:** TEXT connections in the Connections collection cannot be loaded via `QueryTables.Add()` using their connection string. TEXT files must be imported directly as QueryTables, not as separate Connection objects first.

## Testing Strategy

**Connection Type Selection for Tests:**
- **OLEDB** - Use for LoadTo, Refresh, and QueryTable operations (primary use case)
- **TEXT** - Use for connection lifecycle tests (List, View, Delete) without LoadTo
- **ODBC** - Use for validation of multiple connection types

## Key Takeaways

1. **OLEDB is primary for LoadTo** - QueryTable pattern requires OLEDB/ODBC connections
2. **TEXT connections work for lifecycle** - List, View, Delete, but NOT LoadTo
3. **Type 3/4 ambiguity** - Handle both interchangeably for TEXT connection type detection
4. **Always sanitize** - Never expose passwords in connection strings
5. **Test with appropriate type** - OLEDB for data loading, TEXT for lifecycle operations
