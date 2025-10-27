using System.ComponentModel;
using ModelContextProtocol.Server;
using Microsoft.Extensions.AI;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// MCP Prompts for Excel connection types, COM API limitations, and management strategies.
/// </summary>
[McpServerPromptType]
public static class ExcelConnectionPrompts
{
    /// <summary>
    /// Comprehensive guide to Excel connection types, supported operations, and testing strategies.
    /// </summary>
    [McpServerPrompt(Name = "excel_connection_types_guide")]
    [Description("Complete reference for Excel connection types, COM API limitations, and best practices")]
    public static ChatMessage ConnectionTypesGuide()
    {
        return new ChatMessage(ChatRole.User, @"# Excel Connection Types - Complete Reference

## Supported Connection Types

Excel supports 9 connection types via the XlConnectionType enum:

| Type | Name | Description | Creation via COM API | Recommended for Testing |
|------|------|-------------|---------------------|------------------------|
| 1 | OLEDB | OLE DB (SQL Server, Access, etc.) | ❌ UNRELIABLE | No |
| 2 | ODBC | ODBC data sources | ❌ UNRELIABLE | No |
| 3 | TEXT | Text/CSV file imports | ✅ WORKS | ✅ YES |
| 4 | WEB | Web queries/URLs | ⚠️ UNTESTED | Potential |
| 5 | XMLMAP | XML data imports | ⚠️ UNTESTED | Unknown |
| 6 | DATAFEED | Data feed connections | ⚠️ UNTESTED | Unknown |
| 7 | MODEL | Data model connections | ⚠️ UNTESTED | Unknown |
| 8 | WORKSHEET | Worksheet connections | ⚠️ UNTESTED | Unknown |
| 9 | NOSOURCE | No source connections | ⚠️ UNTESTED | Unknown |

## CRITICAL: Excel COM API Limitations

### ❌ OLEDB/ODBC Connections FAIL via Connections.Add()

**Problem**: Excel's `Connections.Add()` method throws **""Value does not fall within the expected range""** for OLEDB and ODBC connections.

**Attempted Fixes (all failed)**:
- Named parameters (doesn't work)
- Positional parameters (doesn't work)
- Different connection string formats (doesn't work)

**This is a known Excel COM API limitation** - not a bug in your code!

**Workaround for Users**:
1. Create OLEDB/ODBC connections via **Excel UI** (Data → Get Data → From Database)
2. Import connections from **.odc files** (Office Data Connection XML files)
3. Use **ConnectionCommands to MANAGE existing connections** (not create from scratch)

### ✅ TEXT Connections WORK Reliably

**Connection String Format**:
```
TEXT;C:\path\to\file.csv
```

**Why TEXT Connections are Recommended**:
- ✅ `Connections.Add()` succeeds for TEXT type
- ✅ No database or network dependencies
- ✅ Simple CSV files always available in test environments
- ✅ Perfect for testing all CRUD operations

**Current Testing Standard**: All automated tests use TEXT connections.

### ⚠️ Type 3 vs Type 4 Confusion (Known Issue)

**Problem**: When TEXT connections are created with `connectionString = ""TEXT;{filePath}""`, Excel may return **type 4 (WEB)** instead of type 3 (TEXT) when reading back `conn.Type`.

**Symptoms**:
- Create connection with `""TEXT;file.csv""`
- Excel reports `conn.Type = 4` (WEB) instead of 3 (TEXT)
- Code tries to access `conn.WebConnection` instead of `conn.TextConnection`
- Error: **""'System.__ComObject' does not contain a definition for 'WebConnection'""**

**Investigation**: Connection string format may be ambiguous. Excel might interpret ""TEXT;path"" as ""URL;path"".

## Connection Management Philosophy

**ConnectionCommands is designed to MANAGE existing connections**, not create them from scratch.

### What ConnectionCommands DOES:
- ✅ List all connections in workbook
- ✅ View connection details (type, connection string, properties)
- ✅ Export connection definitions to JSON (version control)
- ✅ Import connections from JSON/ODC files
- ✅ Update existing connection properties
- ✅ Delete connections
- ✅ Refresh connection data
- ✅ Test connection validity
- ✅ Load connection-only connections to worksheets
- ✅ Get/Set connection properties (BackgroundQuery, RefreshOnFileOpen, etc.)

### What ConnectionCommands DOES NOT:
- ❌ Generate OLEDB/ODBC connection strings from scratch
- ❌ Create database connections via COM API
- ❌ Bypass Excel COM limitations

### How Users Create Connections:
1. **Excel UI**: Data → Get Data → From Database/Web/File
2. **ODC Files**: Office Data Connection files (XML format)
3. **Power Query**: For M code-based connections (use `excel_powerquery` tool instead)

## Connection String Formats

### OLEDB (Manage Existing Only)
```
Provider=SQLOLEDB;Data Source=server;Initial Catalog=database;User ID=user;Password=pass;
```

### ODBC (Manage Existing Only)
```
DSN=MyDataSource;UID=username;PWD=password;
```

### TEXT (Can Create via COM API)
```
TEXT;C:\path\to\file.csv
```

### WEB (Untested)
```
URL;https://example.com/data.xml
```

### Power Query Connections (Use excel_powerquery Tool)
```
OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=QueryName
```
**Note**: Power Query uses special OLEDB provider. Use `excel_powerquery` tool, not `excel_connection`.

## Connection Type-Specific COM Objects

Each connection type has its own COM object:

### Type 1: OLEDB
```csharp
if (conn.Type == 1) {
    dynamic oledb = conn.OLEDBConnection;
    oledb.BackgroundQuery = true;
    oledb.RefreshOnFileOpen = false;
    oledb.CommandText = ""SELECT * FROM Table"";
}
```

### Type 2: ODBC
```csharp
if (conn.Type == 2) {
    dynamic odbc = conn.ODBCConnection;
    odbc.BackgroundQuery = true;
    odbc.CommandText = ""SELECT * FROM Table"";
}
```

### Type 3: TEXT
```csharp
if (conn.Type == 3) {
    dynamic text = conn.TextConnection;
    text.TextFilePlatform = 65001; // UTF-8
    text.TextFileCommaDelimiter = true;
    text.TextFileParseType = 1; // Delimited
}
```

### Type 4: WEB
```csharp
if (conn.Type == 4) {
    dynamic web = conn.WebConnection;
    web.URL = ""https://example.com/data.xml"";
    web.RefreshOnFileOpen = false;
}
```

## Security Best Practices

### 🔒 Password Handling
- **ALWAYS sanitize connection strings** before displaying or exporting
- **SavePassword = false** by default (never export credentials)
- **Warn users** before exporting to ODC files (may contain sensitive data)

### Example: Password Sanitization
```csharp
string safe = ConnectionHelpers.SanitizeConnectionString(rawConnectionString);
// Passwords masked with ***
```

## Common Usage Scenarios

### Scenario 1: List Connections
```
AI: excel_connection(action=""list"", excelPath=""workbook.xlsx"")
→ Returns all connections with types, names, and basic properties
```

### Scenario 2: View Connection Details
```
AI: excel_connection(action=""view"", excelPath=""workbook.xlsx"", connectionName=""SalesDB"")
→ Returns detailed connection info (sanitized connection string, command text, properties)
```

### Scenario 3: Export for Version Control
```
AI: excel_connection(action=""export"", excelPath=""workbook.xlsx"",
                     connectionName=""SalesDB"", targetPath=""salesdb.json"")
→ Saves connection definition to JSON file
```

### Scenario 4: Configure Connection Properties
```
AI: excel_connection(action=""set-properties"", excelPath=""workbook.xlsx"",
                     connectionName=""SalesDB"",
                     backgroundQuery=true, refreshPeriod=30)
→ Sets auto-refresh every 30 minutes in background
```

### Scenario 5: Load Connection-Only Query to Sheet
```
AI: excel_connection(action=""loadto"", excelPath=""workbook.xlsx"",
                     connectionName=""WebData"", targetPath=""Sheet1"")
→ Loads connection data to worksheet table
```

## Troubleshooting

### Issue: ""Value does not fall within the expected range""
**Cause**: Trying to create OLEDB or ODBC connection via `Connections.Add()`
**Solution**:
- Users must create via Excel UI or import from ODC files
- ConnectionCommands manages existing connections only

### Issue: ""'System.__ComObject' does not contain a definition for 'WebConnection'""
**Cause**: Type mapping mismatch - code expects TEXT (type 3) but Excel returned WEB (type 4)
**Solution**:
- Verify connection string format
- Check if ""TEXT;path"" is being misinterpreted as ""URL;path""
- This is under investigation (as of 2025-10-27)

### Issue: Power Query Connection Shows in List
**Cause**: Power Query connections use special OLEDB provider
**Solution**: Use `excel_powerquery` tool for Power Query connections, not `excel_connection`

## Available MCP Actions

All 11 actions for `excel_connection` tool:

1. **list** - Enumerate all connections
2. **view** - Display connection details
3. **import** - Import from JSON file
4. **export** - Export to JSON file
5. **update** - Modify from JSON file
6. **refresh** - Reload data from source
7. **delete** - Remove connection
8. **loadto** - Load to worksheet
9. **properties** - Get connection settings
10. **set-properties** - Modify connection settings
11. **test** - Validate connection

## Key Takeaways for AI Assistants

✅ **DO**:
- Use TEXT connections for testing
- Manage existing OLEDB/ODBC connections (don't try to create)
- Sanitize connection strings (automatic in ConnectionCommands)
- Export connection definitions for version control
- Configure refresh behavior and query settings

❌ **DON'T**:
- Try to create OLEDB/ODBC connections via Connections.Add()
- Assume all connection types can be created programmatically
- Export passwords without user consent
- Use `excel_connection` for Power Query connections (use `excel_powerquery` instead)

📚 **Reference**: Microsoft official documentation - https://learn.microsoft.com/en-us/office/vba/api/excel.xlconnectiontype
");
    }

    /// <summary>
    /// Quick reference for connection management operations and common workflows.
    /// </summary>
    [McpServerPrompt(Name = "excel_connection_workflow_examples")]
    [Description("Common connection management workflows and example usage patterns")]
    public static ChatMessage ConnectionWorkflowExamples()
    {
        return new ChatMessage(ChatRole.User, @"# Excel Connection Management - Workflow Examples

## Workflow 1: Discover Existing Connections

**Goal**: See what data connections exist in a workbook

```
Step 1: List all connections
→ excel_connection(action=""list"", excelPath=""report.xlsx"")

Step 2: View specific connection details
→ excel_connection(action=""view"", excelPath=""report.xlsx"", connectionName=""SalesDB"")

Step 3: Check connection properties
→ excel_connection(action=""properties"", excelPath=""report.xlsx"", connectionName=""SalesDB"")
```

**Expected Results**:
- List shows all connections with types (OLEDB, ODBC, TEXT, WEB, etc.)
- View shows connection string (sanitized), command text, description
- Properties shows BackgroundQuery, RefreshOnFileOpen, SavePassword, RefreshPeriod

## Workflow 2: Export for Version Control

**Goal**: Save connection definitions to Git repository

```
Step 1: Export connection to JSON
→ excel_connection(action=""export"", excelPath=""report.xlsx"",
                   connectionName=""SalesDB"", targetPath=""connections/salesdb.json"")

Step 2: Commit to Git
→ (AI can suggest: git add connections/salesdb.json && git commit -m ""Export SalesDB connection"")
```

**JSON Structure**:
```json
{
  ""Name"": ""SalesDB"",
  ""Type"": ""OLEDB"",
  ""Description"": ""Sales database connection"",
  ""ConnectionString"": ""Provider=SQLOLEDB;Data Source=server;..."",
  ""Properties"": {
    ""BackgroundQuery"": true,
    ""RefreshOnFileOpen"": false,
    ""SavePassword"": false,
    ""RefreshPeriod"": 0
  }
}
```

## Workflow 3: Configure Auto-Refresh

**Goal**: Set connection to refresh automatically every 30 minutes

```
Step 1: Check current settings
→ excel_connection(action=""properties"", excelPath=""report.xlsx"", connectionName=""WebData"")

Step 2: Enable background refresh
→ excel_connection(action=""set-properties"", excelPath=""report.xlsx"",
                   connectionName=""WebData"",
                   backgroundQuery=true, refreshPeriod=30)

Step 3: Verify changes
→ excel_connection(action=""properties"", excelPath=""report.xlsx"", connectionName=""WebData"")
```

**Result**: Connection refreshes data every 30 minutes without blocking Excel UI.

## Workflow 4: Load Connection-Only Query

**Goal**: Connection exists but data not loaded to any worksheet

```
Step 1: List connections to find connection-only queries
→ excel_connection(action=""list"", excelPath=""analysis.xlsx"")

Step 2: Load data to worksheet
→ excel_connection(action=""loadto"", excelPath=""analysis.xlsx"",
                   connectionName=""CustomerData"", targetPath=""Sheet1"")
```

**Result**: Connection data loaded to Sheet1 as Excel table.

## Workflow 5: Update Connection from JSON

**Goal**: Modify existing connection using saved definition

```
Step 1: Edit JSON file (e.g., change refresh period)
→ (AI can suggest editing the JSON file)

Step 2: Update connection from modified JSON
→ excel_connection(action=""update"", excelPath=""report.xlsx"",
                   connectionName=""SalesDB"", targetPath=""connections/salesdb.json"")

Step 3: Verify changes
→ excel_connection(action=""properties"", excelPath=""report.xlsx"", connectionName=""SalesDB"")
```

## Workflow 6: Test Connection Before Refresh

**Goal**: Validate connection without loading data

```
Step 1: Test connection
→ excel_connection(action=""test"", excelPath=""report.xlsx"", connectionName=""WebAPI"")

Step 2: If test succeeds, refresh data
→ excel_connection(action=""refresh"", excelPath=""report.xlsx"", connectionName=""WebAPI"")
```

**Result**: Data refreshed only if connection is valid.

## Workflow 7: Migrate Connection to New Workbook

**Goal**: Copy connection definition from one workbook to another

```
Step 1: Export from source workbook
→ excel_connection(action=""export"", excelPath=""old-report.xlsx"",
                   connectionName=""SalesDB"", targetPath=""salesdb.json"")

Step 2: Import to target workbook
→ excel_connection(action=""import"", excelPath=""new-report.xlsx"",
                   connectionName=""SalesDB"", targetPath=""salesdb.json"")
```

**Result**: SalesDB connection now exists in both workbooks.

## Workflow 8: Clean Up Obsolete Connections

**Goal**: Remove unused connections from workbook

```
Step 1: List all connections
→ excel_connection(action=""list"", excelPath=""report.xlsx"")

Step 2: Export connection for backup (optional)
→ excel_connection(action=""export"", excelPath=""report.xlsx"",
                   connectionName=""OldData"", targetPath=""backup/olddata.json"")

Step 3: Delete connection
→ excel_connection(action=""delete"", excelPath=""report.xlsx"", connectionName=""OldData"")
```

**Result**: Connection removed, workbook cleaner.

## Common Errors and Solutions

### Error: ""Connection 'XYZ' not found""
**Solution**: Run `list` action first to see available connections. Check spelling of connectionName.

### Error: ""Power Query connection detected. Use excel_powerquery tool.""
**Solution**: For Power Query connections, use `excel_powerquery` tool with actions like `pq-list`, `pq-view`, `pq-refresh`.

### Error: Connection string contains sensitive data
**Solution**: ConnectionCommands automatically sanitizes passwords. Review exported JSON before sharing.

### Error: ""Value does not fall within the expected range"" when importing
**Solution**: Connection type may not be creatable via COM API (OLEDB/ODBC). User must create via Excel UI, then you can manage it.

## Security Notes

🔒 **Password Sanitization**: Connection strings are automatically sanitized in all outputs. Passwords replaced with `***`.

🔒 **SavePassword Default**: Always `false` unless explicitly set. Never exports credentials by default.

🔒 **Export Warning**: When exporting to JSON, remind users that connection definitions may contain sensitive data.

## Integration with Other Tools

**excel_powerquery**: For M code-based connections
- List queries: `excel_powerquery(action=""list"")`
- View M code: `excel_powerquery(action=""view"")`
- Refresh: `excel_powerquery(action=""refresh"")`

**excel_worksheet**: For loading connection data
- After refresh: `excel_worksheet(action=""read"")` to get data
- Write results: `excel_worksheet(action=""write"")` from CSV

**excel_parameter**: For connection parameters
- Store connection strings: `excel_parameter(action=""create"", parameterName=""DBServer"")`
- Dynamic connections: Reference parameters in connection definitions
");
    }
}
