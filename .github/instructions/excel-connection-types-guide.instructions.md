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

## Testing Strategy

**Use TEXT connections for all automated tests:**

## Key Takeaways

1. **OLEDB/ODBC creation unreliable** - manage existing connections only
2. **TEXT connections work** - use for testing
3. **Type 3/4 ambiguity** - handle both interchangeably
4. **Always sanitize** - never expose passwords
5. **Test with TEXT** - no DB dependencies
