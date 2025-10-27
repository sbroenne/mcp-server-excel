---
applyTo: "src/ExcelMcp.Core/Commands/ConnectionCommands.cs,src/ExcelMcp.Core/Connections/**/*.cs,tests/**/ConnectionCommands*.cs,tests/**/ConnectionTestHelper.cs"
---

# Excel Connection Types - Complete Reference Guide

> **Comprehensive guide to Excel connection types, COM API behaviors, and testing strategies**

## Excel XlConnectionType Enum

Per [Microsoft official documentation](https://learn.microsoft.com/en-us/office/vba/api/excel.xlconnectiontype):

| Value | Constant | Type Name | Description | Creation Support | Testing Status |
|-------|----------|-----------|-------------|------------------|----------------|
| 1 | xlConnectionTypeOLEDB | OLEDB | OLE DB data sources (SQL Server, Access, etc.) | ❌ **UNRELIABLE** | Not recommended |
| 2 | xlConnectionTypeODBC | ODBC | ODBC data sources | ❌ **UNRELIABLE** | Not recommended |
| 3 | xlConnectionTypeTEXT | TEXT | Text/CSV file imports | ✅ **WORKS** | ✅ Recommended |
| 4 | xlConnectionTypeWEB | WEB | Web queries/URLs | ⚠️ **UNTESTED** | Potential alternative |
| 5 | xlConnectionTypeXMLMAP | XMLMAP | XML data imports | ⚠️ **UNTESTED** | Unknown |
| 6 | xlConnectionTypeDATAFEED | DATAFEED | Data feed connections | ⚠️ **UNTESTED** | Unknown |
| 7 | xlConnectionTypeMODEL | MODEL | Data model connections | ⚠️ **UNTESTED** | Unknown |
| 8 | xlConnectionTypeWORKSHEET | WORKSHEET | Worksheet connections | ⚠️ **UNTESTED** | Unknown |
| 9 | xlConnectionTypeNOSOURCE | NOSOURCE | No source connections | ⚠️ **UNTESTED** | Unknown |

---

## Excel COM API Limitations - CRITICAL

### Connections.Add() Method Issues

**Problem:** Excel's `Connections.Add()` method is **UNRELIABLE** for creating certain connection types programmatically.

#### ❌ OLEDB Connections - DO NOT USE IN TESTS

```csharp
// THIS FAILS WITH "Value does not fall within the expected range"
string connectionString = "Provider=SQLOLEDB;Data Source=server;Initial Catalog=db;";
dynamic conn = connections.Add(
    Name: "TestOleDb",
    Description: "Test connection",
    ConnectionString: connectionString,
    CommandText: ""
);
// Runtime error: "Value does not fall within the expected range"
```

**Excel COM Behavior:**
- `Connections.Add()` throws COMException for OLEDB
- **Named parameters** don't fix this (we tried)
- **Positional parameters** don't fix this (we tried)
- This is a **known Excel COM API limitation**

**Workaround for Production:**
- Users create OLEDB connections via **Excel UI** (Data → Get Data → From Database)
- Users import connections from **.odc files** (Office Data Connection files)
- ConnectionCommands manages **existing** connections, not creation from scratch

**Testing Strategy:**
- ❌ Don't test OLEDB connection creation in automated tests
- ✅ Test connection management (List, View, Export, Import from .odc, Update, Delete, Properties)

#### ❌ ODBC Connections - DO NOT USE IN TESTS

Same issues as OLEDB - `Connections.Add()` fails with "Value does not fall within the expected range".

#### ✅ TEXT Connections - RELIABLE FOR TESTING

```csharp
// THIS WORKS RELIABLY
string csvFilePath = "C:\\test\\data.csv";
string connectionString = $"TEXT;{csvFilePath}";

dynamic conn = connections.Add(
    Name: "TestText",
    Description: "Test text connection",
    ConnectionString: connectionString,
    CommandText: ""
);
// Success! Connection created with Type = 3 (TEXT)
```

**Why TEXT Connections Work:**
- Excel's `Connections.Add()` **successfully creates** TEXT connections
- CSV files are simple and always available in test environments
- No database or network dependencies
- Perfect for testing connection CRUD operations

**Current Testing Standard (as of 2025-10-27):**
All ConnectionCommands integration tests use TEXT connections created via `ConnectionTestHelper.CreateTextFileConnectionAsync()`.

#### ⚠️ Type Mapping Discovery Issue (2025-10-27)

**CRITICAL BUG FOUND:** When TEXT connections are created with `connectionString = "TEXT;{filePath}"`, Excel may return **type 4 (WEB)** instead of type 3 (TEXT) when reading `conn.Type`.

**Symptoms:**
```csharp
// Create TEXT connection
connections.Add(Name: "Test", ConnectionString: "TEXT;file.csv", ...);

// Later read it back
dynamic conn = FindConnection(book, "Test");
int type = conn.Type; // Returns 4 (WEB) instead of 3 (TEXT)!
```

**Impact:**
- `ConnectionHelpers.GetConnectionTypeName(4)` returns "WEB"
- Code tries to access `conn.WebConnection` instead of `conn.TextConnection`
- Runtime error: "'System.__ComObject' does not contain a definition for 'WebConnection'"

**Current Investigation:**
- ConnectionString format may be ambiguous: "TEXT;path" vs "URL;path"
- Excel may be interpreting our TEXT connections as WEB connections
- Need to verify correct connection string format for TEXT type

---

## Connection Type-Specific APIs

Each connection type has its own COM object with specific properties:

### OLEDB Connection (Type 1)
```csharp
if (conn.Type == 1) {
    dynamic oledb = conn.OLEDBConnection;
    oledb.BackgroundQuery = true;
    oledb.RefreshOnFileOpen = false;
    oledb.CommandText = "SELECT * FROM Table";
    oledb.Connection = "Provider=SQLOLEDB;...";
}
```

### ODBC Connection (Type 2)
```csharp
if (conn.Type == 2) {
    dynamic odbc = conn.ODBCConnection;
    odbc.BackgroundQuery = true;
    odbc.RefreshOnFileOpen = false;
    odbc.CommandText = "SELECT * FROM Table";
    odbc.Connection = "DSN=MyDSN;UID=user;PWD=pass;";
}
```

### TEXT Connection (Type 3)
```csharp
if (conn.Type == 3) {
    dynamic text = conn.TextConnection;
    text.TextFilePlatform = 65001; // UTF-8
    text.TextFileCommaDelimiter = true;
    text.TextFileParseType = 1; // Delimited
}
```

### WEB Connection (Type 4)
```csharp
if (conn.Type == 4) {
    dynamic web = conn.WebConnection;
    web.URL = "https://example.com/data.xml";
    web.RefreshOnFileOpen = false;
}
```

### Other Types (5-9)
Similar pattern - each has its own connection object property.

---

## Connection String Formats

### OLEDB
```
Provider=SQLOLEDB;Data Source=server;Initial Catalog=database;User ID=user;Password=pass;
```

### ODBC
```
DSN=MyDataSource;UID=username;PWD=password;
```

### TEXT (CSV Files)
```
TEXT;C:\path\to\file.csv
```
**NOTE:** May be interpreted as WEB type - investigation ongoing.

### WEB (URLs)
```
URL;https://example.com/data.xml
```

### Power Query Connections
```
OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=QueryName
```
**Note:** Power Query uses special OLEDB provider, managed via `pq-*` commands, not `conn-*` commands.

---

## ConnectionCommands Design Philosophy

### Purpose: Manage Existing Connections

ConnectionCommands is designed to **MANAGE** connections that already exist in workbooks:
- List connections
- View connection details
- Export connection definitions to .odc files
- Import connections from .odc files (created by Excel or external tools)
- Update existing connection properties
- Delete connections
- Refresh connection data
- Test connection validity
- Load connection-only connections to worksheets

### NOT for Creating Connections from Scratch

Users create connections via:
1. **Excel UI** - Data → Get Data → From Database/Web/File
2. **.odc Files** - Office Data Connection files (XML format)
3. **Power Query** - For M code-based connections (use `pq-*` commands)

ConnectionCommands **imports** from .odc files but doesn't generate raw connection strings.

---

## Testing Strategy - Current Best Practices (2025-10-27)

### ✅ What to Test

1. **List** - Enumerate all connections in workbook
2. **View** - Display connection details (type, string, properties)
3. **Export** - Save connection definition to JSON
4. **Import** - Load connection from JSON (for connections that Excel API supports)
5. **Update** - Modify existing connection properties
6. **Delete** - Remove connection from workbook
7. **GetProperties** - Read connection settings (BackgroundQuery, RefreshOnFileOpen, etc.)
8. **SetProperties** - Modify connection settings
9. **Test** - Validate connection is accessible
10. **Refresh** - Reload data from source
11. **LoadTo** - Load connection-only connection to worksheet table

### ✅ Use TEXT Connections for All Tests

```csharp
// Test helper pattern
public static async Task CreateTextFileConnectionAsync(
    string filePath, 
    string connectionName, 
    string csvFilePath)
{
    await using var batch = await ExcelSession.BeginBatchAsync(filePath);
    await batch.ExecuteAsync<int>((ctx, ct) =>
    {
        dynamic connections = ctx.Book.Connections;
        
        // Create CSV file if needed
        if (!File.Exists(csvFilePath))
        {
            File.WriteAllText(csvFilePath, "Col1,Col2\nVal1,Val2\n");
        }
        
        // Create TEXT connection
        string connectionString = $"TEXT;{csvFilePath}";
        dynamic conn = connections.Add(
            Name: connectionName,
            Description: "Test text connection",
            ConnectionString: connectionString,
            CommandText: ""
        );
        
        return ValueTask.FromResult(0);
    });
    await batch.SaveAsync();
}
```

### ❌ Don't Test OLEDB/ODBC Creation

```csharp
// DON'T DO THIS - Excel COM API fails
await ConnectionTestHelper.CreateOleDbConnectionAsync(file, "Test", oledbString);
// Throws: "Value does not fall within the expected range"
```

### ⚠️ Current Known Issues

1. **Type 3 vs Type 4 Confusion**
   - TEXT connections may be read back as WEB type
   - Investigating connection string format requirements
   - Tests currently failing: List, View, Update, SetProperties

2. **Error Message**
   ```
   'System.__ComObject' does not contain a definition for 'WebConnection'
   ```
   - Occurs when code expects type 3 (TEXT) but Excel reports type 4 (WEB)
   - UpdateConnectionProperties tries to access wrong connection object

---

## Security Considerations

### Password Handling

```csharp
// ALWAYS sanitize before displaying
string safe = ConnectionHelpers.SanitizeConnectionString(rawConnectionString);
Console.WriteLine(safe); // Passwords masked with ***
```

### Default Settings

```csharp
// Never save passwords by default
oledb.SavePassword = false;
text.SavePassword = false;
web.SavePassword = false;
```

### Export Safety

```csharp
// Warn users before exporting to .odc
if (connection contains credentials) {
    Warn("Connection may contain sensitive data. Export carefully.");
}
```

---

## Common Scenarios

### Scenario 1: User Created Connection in Excel UI

1. User opens Excel
2. Data → Get Data → From SQL Server
3. Enters server, database, credentials
4. Creates connection named "SalesDB"
5. Uses ExcelMcp: `excelcli conn-list workbook.xlsx` → sees "SalesDB"
6. Uses ExcelMcp: `excelcli conn-export workbook.xlsx SalesDB sales.odc` → exports definition

### Scenario 2: Import Connection from .odc File

1. User has `sales.odc` file (created by Excel or external tool)
2. Uses ExcelMcp: `excelcli conn-import workbook.xlsx SalesDB sales.odc`
3. Connection loaded into workbook
4. Uses ExcelMcp: `excelcli conn-refresh workbook.xlsx SalesDB` → data loaded

### Scenario 3: Manage Connection Properties

1. User has existing connection "WebData"
2. Uses ExcelMcp: `excelcli conn-properties workbook.xlsx WebData` → views settings
3. Uses ExcelMcp: `excelcli conn-set-properties workbook.xlsx WebData '{"BackgroundQuery": true, "RefreshPeriod": 60}'`
4. Connection now auto-refreshes every 60 minutes in background

### Scenario 4: Automated Testing (Current Pattern)

1. Test creates CSV file: `File.WriteAllText("data.csv", "Name,Value\nTest,100\n");`
2. Test creates TEXT connection: `CreateTextFileConnectionAsync(file, "TestConn", "data.csv");`
3. Test verifies List: `var result = await commands.ListAsync(batch);` → finds "TestConn"
4. Test verifies View: `var details = await commands.ViewAsync(batch, "TestConn");` → gets properties
5. Test exports: `await commands.ExportAsync(batch, "TestConn", "export.json");` → JSON created
6. Test deletes: `await commands.DeleteAsync(batch, "TestConn");` → removed
7. Test re-imports: `await commands.ImportAsync(batch, "TestConn", "export.json");` → restored

---

## Type Mapping Reference (ConnectionHelpers.cs)

```csharp
public static string GetConnectionTypeName(int connectionType)
{
    return connectionType switch
    {
        1 => "OLEDB",      // xlConnectionTypeOLEDB
        2 => "ODBC",       // xlConnectionTypeODBC
        3 => "TEXT",       // xlConnectionTypeTEXT
        4 => "WEB",        // xlConnectionTypeWEB
        5 => "XMLMAP",     // xlConnectionTypeXMLMAP
        6 => "DATAFEED",   // xlConnectionTypeDATAFEED
        7 => "MODEL",      // xlConnectionTypeMODEL
        8 => "WORKSHEET",  // xlConnectionTypeWORKSHEET
        9 => "NOSOURCE",   // xlConnectionTypeNOSOURCE
        _ => $"Unknown ({connectionType})"
    };
}
```

**⚠️ Critical:** This mapping matches Microsoft's official XlConnectionType enum values. However, Excel may return unexpected type values when reading back connections created via `Connections.Add()`.

---

## Troubleshooting

### Issue: "Value does not fall within the expected range"

**Cause:** Trying to create OLEDB or ODBC connection via `Connections.Add()`

**Solution:** 
- Use TEXT connections for testing
- For production, import from .odc files or let users create via Excel UI

### Issue: "'System.__ComObject' does not contain a definition for 'WebConnection'"

**Cause:** Type mapping mismatch - code expects TEXT (type 3) but Excel returned WEB (type 4)

**Solution:** 
- Verify connection string format
- Check if "TEXT;path" is being misinterpreted as "URL;path"
- Investigation ongoing (as of 2025-10-27)

### Issue: Tests Pass Locally, Fail in CI

**Cause:** CI environment doesn't have Excel installed

**Solution:**
- Mark connection tests with `[Trait("RequiresExcel", "true")]`
- CI should filter: `dotnet test --filter "RequiresExcel!=true"`
- Only run locally or on test machines with Excel

---

## Future Enhancements

### Potential Improvements

1. **WEB Connection Testing**
   - Verify if WEB connections (type 4) can be reliably created
   - May be alternative to TEXT for testing
   
2. **Connection String Format Investigation**
   - Document exact format requirements for each type
   - Clarify TEXT vs WEB ambiguity
   
3. **Enhanced Error Messages**
   - When Connections.Add() fails, provide helpful guidance
   - "This connection type requires Excel UI or .odc import"

4. **Type Detection Logic**
   - Add validation when creating connections
   - Warn if Excel returns unexpected type value

---

## Key Takeaways for LLM

1. **OLEDB/ODBC creation via COM API is unreliable** - don't use in tests
2. **TEXT connections work reliably** - use for all automated testing
3. **Type 3 vs Type 4 ambiguity exists** - investigation ongoing
4. **ConnectionCommands manages existing connections** - not for creating from scratch
5. **Always sanitize connection strings** - security first
6. **Tests require Excel installed** - mark appropriately
7. **Each connection type has specific COM object** - OLEDBConnection, TextConnection, etc.

