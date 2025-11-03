# Connection Management Feature Specification

## Overview

Add comprehensive connection management capabilities to ExcelMcp, mirroring the existing Power Query commands architecture. This enables programmatic manipulation of Excel data connections (OLEDB, ODBC, Text, Web, XML, etc.) through both CLI and MCP Server interfaces.

## Objectives

1. Provide full CRUD operations for Excel connections matching Power Query functionality
2. Support all Excel connection types via COM interop
3. Maintain architectural consistency with existing PowerQueryCommands pattern
4. Enable AI-assisted connection development workflows through MCP Server
5. Support connection definition import/export for version control

## Verified Excel COM API Capabilities

Based on official Microsoft documentation review:

### WorkbookConnection Object

- **Methods Available:**
  - `Delete()` - Remove connection ✅
  - `Refresh()` - Refresh connection data ✅

- **Properties Available (Read/Write unless noted):**
  - `Name` - Connection name ✅
  - `Description` - Connection description ✅
  - `Type` - Connection type (Read-Only) ✅
  - `OLEDBConnection` - Access to OLEDB-specific properties ✅
  - `ODBCConnection` - Access to ODBC-specific properties ✅
  - `TextConnection` - Access to Text-specific properties ✅
  - `WorksheetDataConnection` - Access to Worksheet-specific properties ✅
  - `DataFeedConnection` - Access to Data Feed-specific properties ✅
  - `ModelConnection` - Access to PowerPivot Model properties ✅
  - `RefreshWithRefreshAll` - Include in RefreshAll operations ✅
  - `InModel` - Whether connection is in Data Model (Read-Only) ✅

### OLEDBConnection Object Properties

- `Connection` - Connection string (Read/Write) ✅
- `CommandText` - SQL command or table name (Read/Write) ✅
- `CommandType` - Command type enum (Read/Write) ✅
- `BackgroundQuery` - Refresh in background (Read/Write) ✅
- `RefreshOnFileOpen` - Auto-refresh on open (Read/Write) ✅
- `RefreshPeriod` - Auto-refresh interval (Read/Write) ✅
- `SavePassword` - Store credentials (Read/Write) ⚠️ Security Risk
- `RefreshDate` - Last refresh timestamp (Read-Only) ✅
- `MaintainConnection` - Keep connection alive (Read/Write) ✅
- `EnableRefresh` - Allow refreshing (Read/Write) ✅

### ODBCConnection Object Properties

Similar to OLEDBConnection with same core properties ✅

### Connection Type Enumeration (XlConnectionType)

| Value | Name | Description |
|-------|------|-------------|
| 1 | xlConnectionTypeOLEDB | OLE DB connection |
| 2 | xlConnectionTypeODBC | ODBC connection |
| 3 | xlConnectionTypeXMLMAP | XML MAP connection |
| 4 | xlConnectionTypeTEXT | Text file connection |
| 5 | xlConnectionTypeWEB | Web query connection |
| 6 | xlConnectionTypeDATAFEED | Data Feed connection |
| 7 | xlConnectionTypeMODEL | PowerPivot Model connection |
| 8 | xlConnectionTypeWORKSHEET | Worksheet connection |
| 9 | xlConnectionTypeNOSOURCE | No source connection |

## Functional Requirements

### Core Operations (Mirroring Power Query)

#### 1. List Connections (`conn-list`)

**Purpose:** Display all data connections in a workbook

**CLI Usage:**

```powershell
excelcli conn-list "workbook.xlsx"
```

**Output:** Table showing:

- Connection Name
- Type (OLEDB, ODBC, Text, etc.)
- Description
- Last Refresh Date
- Refresh Settings (Background, On Open)

**Implementation:**

- Iterate `workbook.Connections` collection
- Map Type enum to human-readable names
- Extract type-specific properties via OLEDBConnection/ODBCConnection/etc.

---

#### 2. View Connection Details (`conn-view`)

**Purpose:** Display complete connection configuration

**CLI Usage:**

```powershell
excelcli conn-view "workbook.xlsx" "MyConnection"
```

**Output:**

- Connection name, type, description
- Connection string (sanitized - passwords masked)
- Command text (SQL query, table name, etc.)
- Command type
- All refresh settings
- Full JSON definition for export

**Security Note:** Password values must be masked/removed from output

---

#### 3. Import Connection (`conn-import`)

**Purpose:** Create new connection from JSON definition file

**CLI Usage:**

```powershell
excelcli conn-import "workbook.xlsx" "NewConnection" "connection-def.json"
```

**JSON Definition Format:**

```json
{
  "type": "OLEDB",
  "description": "Sales Database Connection",
  "connectionString": "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=Sales",
  "commandText": "SELECT * FROM Customers",
  "commandType": "SQL",
  "backgroundQuery": false,
  "refreshOnFileOpen": true,
  "savePassword": false
}
```

**Validation:**

- Verify connection name doesn't already exist
- Validate required fields for connection type
- Enforce `savePassword: false` by default (security)

**Limitations:**

- Cannot create Power Query connections (use `pq-import` instead)
- Type-specific: Different fields required for OLEDB vs ODBC vs Text

---

#### 4. Export Connection (`conn-export`)

**Purpose:** Export connection definition to JSON file

**CLI Usage:**

```powershell
excelcli conn-export "workbook.xlsx" "MyConnection" "connection-def.json"
```

**Output:** JSON file with complete connection configuration
**Use Case:** Version control, sharing, backup

---

#### 5. Update Connection (`conn-update`)

**Purpose:** Modify existing connection from JSON definition

**CLI Usage:**

```powershell
excelcli conn-update "workbook.xlsx" "MyConnection" "updated-def.json"
```

**Behavior:**

- Replace connection string, command text, etc.
- Preserve connection name unless explicitly changed
- Validate changes before applying

**Limitations:**

- Cannot change connection Type (delete and recreate instead)
- Power Query connections read-only (use `pq-update` instead)

---

#### 6. Refresh Connection (`conn-refresh`)

**Purpose:** Refresh connection data

**CLI Usage:**

```powershell
excelcli conn-refresh "workbook.xlsx" "MyConnection"
```

**Implementation:**

- Call `connection.Refresh()` method
- Handle COM exceptions (connection failures, timeout, etc.)
- Report success/failure with error details

**Note:** Does NOT use `workbook.RefreshAll()` (known to hang)

---

#### 7. Delete Connection (`conn-delete`)

**Purpose:** Remove connection from workbook

**CLI Usage:**

```powershell
excelcli conn-delete "workbook.xlsx" "MyConnection"
```

**Implementation:**

- Find connection by name
- Call `connection.Delete()`
- Remove associated QueryTables if applicable

**Warning:** Destructive operation, verify before executing

---

#### 8. Load Connection to Worksheet (`conn-loadto`)

**Purpose:** Create QueryTable to load connection data to worksheet

**CLI Usage:**

```powershell
excelcli conn-loadto "workbook.xlsx" "MyConnection" "DataSheet"
```

**Implementation:**

- Create or find target worksheet
- Create QueryTable using connection
- Configure refresh settings
- Immediate refresh to load data

**Similar to:** `pq-loadto` for Power Query

---

#### 9. Get Connection Properties (`conn-properties`)

**Purpose:** View refresh and behavior settings

**CLI Usage:**

```powershell
excelcli conn-properties "workbook.xlsx" "MyConnection"
```

**Output:**

- BackgroundQuery
- RefreshOnFileOpen
- RefreshPeriod
- SavePassword
- MaintainConnection
- EnableRefresh

---

#### 10. Set Connection Properties (`conn-set-properties`)

**Purpose:** Modify connection behavior settings

**CLI Usage:**

```powershell
excelcli conn-set-properties "workbook.xlsx" "MyConnection" --background=false --refresh-on-open=true
```

**Parameters:**

- `--background` - Enable/disable background refresh
- `--refresh-on-open` - Enable/disable refresh when file opens
- `--save-password` - Enable/disable password saving (⚠️ security warning)

---

#### 11. Test Connection (`conn-test`) ✨ NEW

**Purpose:** Validate connection without saving workbook

**CLI Usage:**

```powershell
excelcli conn-test "workbook.xlsx" "DatabaseConnection"
```

**Behavior:**

- Attempts to refresh connection
- Captures error messages if refresh fails
- Does NOT save workbook after test
- Returns detailed error information for debugging

**Output:**

```
✓ Connection 'DatabaseConnection' tested successfully
  Last Refresh: 2025-10-21 14:30:45
  
OR

✗ Connection 'DatabaseConnection' failed to connect
  Error: Login failed for user 'sa'
  Type: OLEDB
  Connection String: Provider=SQLOLEDB;Data Source=localhost;...
```

**Use Cases:**

- Validate credentials before production use
- Troubleshoot connection issues
- Verify network connectivity
- Test connection string changes

---

## Architecture & Implementation

### File Structure (Mirroring Power Query)

```
src/ExcelMcp.Core/
  Commands/
    IConnectionCommands.cs          # Interface (NEW)
    ConnectionCommands.cs           # Core implementation (NEW)
  Models/
    ResultTypes.cs                  # Add ConnectionListResult, etc. (UPDATED)
    
src/ExcelMcp.CLI/
  Commands/
    ConnectionCommands.cs           # CLI presentation layer (NEW)
  Program.cs                        # Add conn-* routing (UPDATED)

src/ExcelMcp.McpServer/
  Tools/
    ExcelConnectionTool.cs          # MCP Server tool (NEW)
```

### Model Classes Required

```csharp
// Result types
public class ConnectionListResult : ResultBase
{
    public List<ConnectionInfo> Connections { get; set; }
}

public class ConnectionInfo
{
    public string Name { get; set; }
    public string Description { get; set; }
    public string Type { get; set; }  // Human-readable: "OLEDB", "ODBC", etc.
    public DateTime? LastRefresh { get; set; }
    public bool BackgroundQuery { get; set; }
    public bool RefreshOnFileOpen { get; set; }
}

public class ConnectionViewResult : ResultBase
{
    public string ConnectionName { get; set; }
    public string Type { get; set; }
    public string ConnectionString { get; set; }  // Sanitized
    public string CommandText { get; set; }
    public string CommandType { get; set; }
    public string DefinitionJson { get; set; }
}

public class ConnectionPropertiesResult : ResultBase
{
    public string ConnectionName { get; set; }
    public bool BackgroundQuery { get; set; }
    public bool RefreshOnFileOpen { get; set; }
    public bool SavePassword { get; set; }
    public int RefreshPeriod { get; set; }
}
```

### Helper Methods Pattern

```csharp
// In ExcelHelper.cs
public static dynamic? FindConnection(dynamic workbook, string connectionName)
{
    dynamic connections = workbook.Connections;
    for (int i = 1; i <= connections.Count; i++)
    {
        dynamic conn = connections.Item(i);
        if (conn.Name == connectionName) return conn;
    }
    return null;
}

public static string GetConnectionTypeName(int typeValue)
{
    return typeValue switch
    {
        1 => "OLEDB",
        2 => "ODBC",
        3 => "XML",
        4 => "Text",
        5 => "Web",
        6 => "DataFeed",
        7 => "Model",
        8 => "Worksheet",
        9 => "NoSource",
        _ => "Unknown"
    };
}
```

## Security Considerations

### Password Handling

1. **Never log connection strings** - may contain credentials
2. **Sanitize output** - remove password values from ConnectionString before display
3. **Default SavePassword=false** - require explicit opt-in
4. **Warn users** - display security warning when SavePassword=true

### Connection String Sanitization

```csharp
private static string SanitizeConnectionString(string connString)
{
    // Remove password components
    var regex = new Regex(@"(Password|PWD)\s*=\s*[^;]*;?", RegexOptions.IgnoreCase);
    return regex.Replace(connString, "$1=***;");
}
```

## Limitations & Constraints

### Cannot Do

1. **Create Power Query connections** - Power Queries are created via `Queries` collection, connections are side-effect
2. **Modify Power Query connection strings** - These use Mashup provider, managed by M code
3. **Access Data Model internals** - Requires separate Analysis Services Tabular API
4. **Retrieve stored passwords** - Excel never exposes credentials via COM
5. **Test connection validity** - No built-in test method, only refresh and catch errors
6. **Change connection Type** - Type is read-only, must delete and recreate

### Type-Specific Handling Required

- **OLEDB:** Connection, CommandText, CommandType properties
- **ODBC:** Connection, CommandText, CommandType properties  
- **Text:** Different property set (file path, delimiters, etc.)
- **Web:** URL-based configuration
- **Power Query (Mashup):** Read-only, use `pq-*` commands instead

## MCP Server Integration

### Tool Definition

```json
{
  "name": "excel_connection",
  "description": "Manage Excel data connections (OLEDB, ODBC, Text, Web, etc.)",
  "inputSchema": {
    "type": "object",
    "properties": {
      "action": {
        "type": "string",
        "enum": [
          "list", "view", "import", "export", "update", 
          "refresh", "delete", "loadto", 
          "get-properties", "set-properties"
        ]
      },
      "excelPath": { "type": "string" },
      "connectionName": { "type": "string" },
      "definitionFile": { "type": "string" },
      "sheetName": { "type": "string" },
      "backgroundQuery": { "type": "boolean" },
      "refreshOnFileOpen": { "type": "boolean" },
      "savePassword": { "type": "boolean" }
    },
    "required": ["action", "excelPath"]
  }
}
```

### Use Cases

- **AI-assisted connection management** - LLM helps configure database connections
- **Connection troubleshooting** - View/modify connection strings to fix issues
- **Bulk connection updates** - Update multiple connections across workbooks
- **Connection documentation** - Export definitions for documentation

## Comprehensive Testing Strategy

Following the established three-tier testing architecture (Unit → Integration → RoundTrip):

### Test File Organization

```
tests/ExcelMcp.Core.Tests/
  Unit/
    Helpers/
      ConnectionTypeMapperTests.cs        # Type enum to string mapping
      ConnectionStringSanitizerTests.cs   # Password removal/masking
      ConnectionJsonParserTests.cs        # JSON definition validation
  Integration/
    Commands/
      ConnectionCommandsTests.cs          # Core connection CRUD operations
      ConnectionSecurityTests.cs          # Password sanitization in practice
      ConnectionTypeTests.cs              # Type-specific operations
  RoundTrip/
    Commands/
      ConnectionWorkflowTests.cs          # Complete workflows

tests/ExcelMcp.CLI.Tests/
  Unit/
    Commands/
      CliConnectionCommandsTests.cs       # CLI argument parsing
  Integration/
    Commands/
      CliConnectionCommandsTests.cs       # CLI execution with Excel

tests/ExcelMcp.McpServer.Tests/
  Integration/
    Tools/
      ExcelConnectionToolTests.cs         # MCP tool operations
  RoundTrip/
    McpConnectionWorkflowTests.cs         # MCP protocol workflows
```

### Unit Tests (Fast, No Excel Required)

#### 1. **ConnectionTypeMapper Tests** ✅

```csharp
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
public class ConnectionTypeMapperTests
{
    [Theory]
    [InlineData(1, "OLEDB")]
    [InlineData(2, "ODBC")]
    [InlineData(3, "XML")]
    [InlineData(4, "Text")]
    [InlineData(5, "Web")]
    [InlineData(6, "DataFeed")]
    [InlineData(7, "Model")]
    [InlineData(8, "Worksheet")]
    [InlineData(9, "NoSource")]
    [InlineData(99, "Unknown(99)")]
    public void GetConnectionTypeName_WithValidTypes_ReturnsCorrectName(int typeValue, string expected)
    
    [Fact]
    public void IsPowerQueryConnection_WithMashupProvider_ReturnsTrue()
    
    [Fact]
    public void IsPowerQueryConnection_WithRegularOLEDB_ReturnsFalse()
}
```

#### 2. **ConnectionString Sanitization Tests** ✅

```csharp
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
public class ConnectionStringSanitizerTests
{
    [Theory]
    [InlineData(
        "Provider=SQLOLEDB;Data Source=localhost;Password=secret123;User ID=admin",
        "Provider=SQLOLEDB;Data Source=localhost;Password=***;User ID=admin")]
    [InlineData(
        "Server=localhost;PWD=p@ssw0rd;Database=test",
        "Server=localhost;PWD=***;Database=test")]
    [InlineData(
        "Connection string with no password",
        "Connection string with no password")]
    public void SanitizeConnectionString_RemovesPasswords(string input, string expected)
    
    [Fact]
    public void SanitizeConnectionString_WithMultiplePasswordFormats_RemovesAll()
    
    [Fact]
    public void SanitizeConnectionString_WithEmptyString_ReturnsEmpty()
    
    [Fact]
    public void SanitizeConnectionString_WithNull_ThrowsArgumentNullException()
}
```

#### 3. **Connection JSON Parser Tests** ✅

```csharp
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
public class ConnectionJsonParserTests
{
    [Fact]
    public void ParseConnectionDefinition_WithValidOLEDB_ReturnsDefinition()
    
    [Fact]
    public void ParseConnectionDefinition_WithValidODBC_ReturnsDefinition()
    
    [Fact]
    public void ParseConnectionDefinition_WithMissingType_ThrowsValidationException()
    
    [Fact]
    public void ParseConnectionDefinition_WithInvalidJSON_ThrowsJsonException()
    
    [Theory]
    [InlineData("OLEDB", "connectionString", "commandText")]
    [InlineData("ODBC", "connectionString", "commandText")]
    public void ValidateDefinition_WithRequiredFields_ReturnsValid(string type, params string[] requiredFields)
}
```

### Integration Tests (Medium Speed, Requires Excel)

#### 4. **Connection CRUD Operations** ✅

```csharp
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "Connections")]
[Trait("RequiresExcel", "true")]
public class CoreConnectionCommandsTests : IDisposable
{
    private readonly IConnectionCommands _connectionCommands;
    private readonly IFileCommands _fileCommands;
    private readonly string _testExcelFile;
    private readonly string _testConnDefFile;
    
    [Fact]
    public void List_WithEmptyWorkbook_ReturnsEmptyList()
    
    [Fact]
    public void List_WithConnections_ReturnsAllConnections()
    
    [Fact]
    public async Task Import_WithValidOLEDBDefinition_CreatesConnection()
    
    [Fact]
    public async Task Import_WithExistingName_ReturnsError()
    
    [Fact]
    public void View_WithValidConnection_ReturnsDetails()
    
    [Fact]
    public void View_WithInvalidConnection_ReturnsError()
    
    [Fact]
    public async Task Export_WithValidConnection_CreatesJSONFile()
    
    [Fact]
    public async Task Update_WithValidDefinition_ModifiesConnection()
    
    [Fact]
    public void Refresh_WithValidConnection_SucceedsOrReturnsError()
    
    [Fact]
    public void Delete_WithValidConnection_RemovesConnection()
    
    [Fact]
    public void Delete_WithInvalidConnection_ReturnsError()
    
    [Fact]
    public void LoadTo_WithValidConnection_CreatesQueryTable()
    
    [Fact]
    public void GetProperties_WithValidConnection_ReturnsProperties()
    
    [Fact]
    public void SetProperties_WithValidValues_UpdatesProperties()
}
```

#### 5. **Connection Security Tests** ✅

```csharp
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "Connections")]
[Trait("RequiresExcel", "true")]
public class ConnectionSecurityTests : IDisposable
{
    [Fact]
    public void View_ConnectionWithPassword_SanitizesOutput()
    {
        // Create connection with password in connection string
        // View connection
        // Assert: Password is masked with ***
    }
    
    [Fact]
    public async Task Export_ConnectionWithPassword_SanitizesJSON()
    {
        // Create connection with password
        // Export to JSON
        // Assert: JSON contains Password=***
    }
    
    [Fact]
    public async Task Import_WithSavePasswordTrue_DisplaysWarning()
    {
        // Create JSON with SavePassword: true
        // Import connection
        // Assert: Result contains security warning
    }
    
    [Fact]
    public async Task Import_WithoutSavePassword_DefaultsToFalse()
    {
        // Create JSON without SavePassword property
        // Import connection
        // Assert: SavePassword is set to false
    }
}
```

#### 6. **Connection Type-Specific Tests** ✅

```csharp
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "Connections")]
[Trait("RequiresExcel", "true")]
public class ConnectionTypeTests : IDisposable
{
    [Fact]
    public async Task Import_OLEDB_WithSQLCommand_CreatesCorrectConnection()
    
    [Fact]
    public async Task Import_ODBC_WithDSN_CreatesCorrectConnection()
    
    [Fact]
    public async Task Import_Text_WithCSVFile_CreatesCorrectConnection()
    
    [Fact]
    public async Task Update_PowerQueryConnection_ReturnsErrorRedirectToPQ()
    {
        // Attempt to update a Power Query connection
        // Assert: Error message directs user to use pq-update command
    }
    
    [Fact]
    public void View_PowerQueryConnection_ShowsMashupProvider()
    {
        // View a Power Query connection
        // Assert: Identifies it as Power Query type
    }
}
```

### Round Trip Tests (Slow, Complete Workflows)

#### 7. **Connection Workflow Tests** ✅

```csharp
[Trait("Category", "RoundTrip")]
[Trait("Speed", "Slow")]
[Trait("Layer", "Core")]
[Trait("Feature", "Connections")]
[Trait("RequiresExcel", "true")]
public class ConnectionWorkflowTests : IDisposable
{
    [Fact]
    public async Task CompleteWorkflow_ImportRefreshExportModifyReimport_Success()
    {
        // 1. Import connection from JSON
        // 2. Refresh connection
        // 3. Export connection to JSON
        // 4. Modify JSON definition
        // 5. Update connection from modified JSON
        // 6. Verify changes applied
        // 7. Delete connection
    }
    
    [Fact]
    public async Task DatabaseConnection_CreateLoadRefreshVerify_Success()
    {
        // 1. Create OLEDB connection to local database
        // 2. Load to worksheet
        // 3. Verify data loaded
        // 4. Modify connection string
        // 5. Refresh
        // 6. Verify updated data
    }
    
    [Fact]
    public async Task MultipleConnections_ManageSimultaneously_Success()
    {
        // 1. Create 3 different connections (OLEDB, ODBC, Text)
        // 2. List all connections
        // 3. Update each connection
        // 4. Refresh all individually
        // 5. Delete all
    }
}
```

### CLI Integration Tests

#### 8. **CLI Connection Commands Tests** ✅

```csharp
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "CLI")]
[Trait("Feature", "Connections")]
[Trait("RequiresExcel", "true")]
public class CliConnectionCommandsTests : IDisposable
{
    [Fact]
    public void ConnList_WithValidArguments_ExitsWithZero()
    
    [Fact]
    public void ConnView_WithValidArguments_DisplaysConnectionDetails()
    
    [Fact]
    public void ConnImport_WithInvalidJSON_ExitsWithOne()
    
    [Fact]
    public void ConnRefresh_WithNonexistentConnection_ShowsError()
    
    [Fact]
    public void ConnSetProperties_WithMultipleFlags_UpdatesCorrectly()
}
```

### MCP Server Tests

#### 9. **MCP Connection Tool Tests** ✅

```csharp
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Connections")]
[Trait("RequiresExcel", "true")]
public class ExcelConnectionToolTests
{
    [Fact]
    public async Task ExcelConnection_ListAction_ReturnsConnectionList()
    
    [Fact]
    public async Task ExcelConnection_ViewAction_ReturnsConnectionDetails()
    
    [Fact]
    public async Task ExcelConnection_ImportAction_CreatesConnection()
    
    [Fact]
    public async Task ExcelConnection_UnknownAction_ThrowsMcpException()
}
```

#### 10. **MCP Connection Workflow Tests** ✅

```csharp
[Trait("Category", "RoundTrip")]
[Trait("Speed", "Slow")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "MCPProtocol")]
[Trait("RequiresExcel", "true")]
public class McpConnectionWorkflowTests : IDisposable
{
    [Fact]
    public async Task McpServer_ConnectionRoundTrip_CompleteWorkflow()
    {
        // Start MCP server process
        // Send list request via JSON-RPC
        // Send import request via JSON-RPC
        // Send refresh request via JSON-RPC
        // Send delete request via JSON-RPC
        // Verify all operations completed successfully
    }
}
```

### Test Data Management

#### Shared Test Fixtures

```csharp
public class ConnectionTestFixtures
{
    public static string GetSampleOLEDBDefinition() => 
        """
        {
          "type": "OLEDB",
          "description": "Test SQL Server Connection",
          "connectionString": "Provider=SQLOLEDB;Data Source=(local);Initial Catalog=tempdb;Integrated Security=SSPI",
          "commandText": "SELECT 1 AS TestColumn",
          "commandType": "SQL",
          "backgroundQuery": false,
          "refreshOnFileOpen": false,
          "savePassword": false
        }
        """;
    
    public static string GetSampleODBCDefinition() => 
        """
        {
          "type": "ODBC",
          "description": "Test ODBC Connection",
          "connectionString": "DSN=Excel Files;DBQ=C:\\temp\\test.xlsx",
          "commandText": "SELECT * FROM [Sheet1$]",
          "commandType": "SQL",
          "backgroundQuery": false,
          "refreshOnFileOpen": false,
          "savePassword": false
        }
        """;
}
```

### Test Execution Matrix

| Test Category | Count | Speed | Excel Required | Run in CI | Run Locally |
|---------------|-------|-------|----------------|-----------|-------------|
| Unit Tests | ~45 | < 1s | No | ✅ Always | ✅ Always |
| Integration Tests | ~80 | 1-5s each | Yes | ❌ No | ✅ Default |
| Round Trip Tests | ~15 | 10-30s each | Yes | ❌ No | ⚠️ On Request |
| **Total** | **~140** | **5-20 min** | - | - | - |

### Test Commands

```powershell
# Run all connection tests
dotnet test --filter "Feature=Connections&RunType!=OnDemand"

# Run specific test class
dotnet test --filter "FullyQualifiedName~CoreConnectionCommandsTests"
```

### Test Coverage Goals

- **Integration Tests:** 95% coverage of all connection functionality
- **Overall:** 95%+ test coverage for connection feature

**Note:** No unit tests (see `docs/ADR-001-NO-UNIT-TESTS.md` for rationale)

## CLI Command Reference

```powershell
# List all connections
excelcli conn-list "workbook.xlsx"

# View connection details
excelcli conn-view "workbook.xlsx" "SalesDB"

# Import connection from definition
excelcli conn-import "workbook.xlsx" "NewConnection" "conn-def.json"

# Export connection definition
excelcli conn-export "workbook.xlsx" "SalesDB" "sales-conn.json"

# Update connection
excelcli conn-update "workbook.xlsx" "SalesDB" "updated-conn.json"

# Refresh connection data
excelcli conn-refresh "workbook.xlsx" "SalesDB"

# Delete connection
excelcli conn-delete "workbook.xlsx" "OldConnection"

# Load connection to worksheet
excelcli conn-loadto "workbook.xlsx" "SalesDB" "DataSheet"

# View connection properties
excelcli conn-properties "workbook.xlsx" "SalesDB"

# Set connection properties
excelcli conn-set-properties "workbook.xlsx" "SalesDB" --background=false --refresh-on-open=true
```

## Documentation Updates Required

1. **README.md** - Add connection commands to capabilities list
2. **COMMANDS.md** - Document all `conn-*` commands with examples
3. **copilot-instructions.md** - Add connection management patterns
4. **MCP Server README** - Document `excel_connection` tool

## Migration Path

### Phase 1: Core Implementation ✅

- [ ] Create IConnectionCommands.cs interface
- [ ] Create ConnectionCommands.cs (Core layer)
- [ ] Add connection result types to ResultTypes.cs
- [ ] Add FindConnection helper to ExcelHelper.cs

### Phase 2: CLI Integration

- [ ] Create ConnectionCommands.cs (CLI layer)
- [ ] Update Program.cs routing
- [ ] Add connection commands to help text

### Phase 3: MCP Server Integration

- [ ] Create ExcelConnectionTool.cs
- [ ] Update server.json with tool definition
- [ ] Add connection actions to MCP server

### Phase 4: Testing & Documentation

- [ ] Unit tests for Core layer
- [ ] Integration tests for CLI
- [ ] Round trip tests for MCP Server
- [ ] Update all documentation
- [ ] Add usage examples

## Success Criteria

1. ✅ All 10 core operations implemented and tested
2. ✅ Parity with Power Query command architecture
3. ✅ Security best practices enforced (password sanitization)
4. ✅ Comprehensive test coverage (Unit + Integration + RoundTrip)
5. ✅ MCP Server tool functional for AI workflows
6. ✅ Complete documentation with examples
7. ✅ Zero build warnings
8. ✅ Follows existing code patterns and conventions

## DRY Analysis: Shared Functionality Extraction

### Identified Reusable Components from PowerQueryCommands

#### 1. **Connection Management Utilities** ✅ EXTRACT TO SHARED

Current: `RemoveQueryConnections()` in PowerQueryCommands.cs

```csharp
// CURRENT - PowerQuery specific
private static void RemoveQueryConnections(dynamic workbook, string queryName)

// PROPOSED - Generic helper in ExcelHelper.cs
public static void RemoveConnections(dynamic workbook, string connectionName)
public static void RemoveQueryTables(dynamic workbook, string connectionName)
```

#### 2. **QueryTable Creation** ✅ EXTRACT TO SHARED

Current: `CreateQueryTableConnection()` in PowerQueryCommands.cs

```csharp
// CURRENT - PowerQuery specific with Mashup provider
private static void CreateQueryTableConnection(dynamic workbook, dynamic targetSheet, string queryName)

// PROPOSED - Generic helper with connection string parameter
public static void CreateQueryTable(dynamic workbook, dynamic targetSheet, string connectionString, 
    string commandText, string queryTableName, QueryTableOptions options = null)
```

#### 3. **Privacy Level Handling** ❌ NOT APPLICABLE

Privacy levels are **Power Query specific** (M code feature)

- Regular connections (OLEDB/ODBC) do NOT have privacy levels
- This functionality stays in PowerQueryCommands
- Connections use different security model (SavePassword property)

#### 4. **Connection String Sanitization** ✅ NEW SHARED UTILITY

Currently: NOT implemented (missing!)

```csharp
// PROPOSED - New utility in ExcelHelper.cs
public static string SanitizeConnectionString(string connectionString)
{
    // Remove password values from connection strings
    var patterns = new[]
    {
        @"(Password|PWD)\s*=\s*[^;]*;?",
        @"(Uid|User\s*Id)\s*=\s*[^;]*;?" // Optional: mask usernames too
    };
    
    var result = connectionString;
    foreach (var pattern in patterns)
    {
        result = Regex.Replace(result, pattern, "$1=***;", RegexOptions.IgnoreCase);
    }
    return result;
}
```

#### 5. **Find Pattern Helpers** ✅ ADD CONNECTION VERSION

Current: `FindQuery()`, `FindName()`, `FindSheet()` exist

```csharp
// PROPOSED - Add to ExcelHelper.cs
public static dynamic? FindConnection(dynamic workbook, string connectionName)
{
    try
    {
        dynamic connections = workbook.Connections;
        for (int i = 1; i <= connections.Count; i++)
        {
            dynamic conn = connections.Item(i);
            if (conn.Name == connectionName) return conn;
        }
    }
    catch { }
    return null;
}

public static List<string> GetConnectionNames(dynamic workbook)
{
    var names = new List<string>();
    try
    {
        dynamic connections = workbook.Connections;
        for (int i = 1; i <= connections.Count; i++)
        {
            names.Add(connections.Item(i).Name);
        }
    }
    catch { }
    return names;
}
```

#### 6. **Connection Type Mapping** ✅ NEW SHARED UTILITY

```csharp
// PROPOSED - Add to ExcelHelper.cs
public static string GetConnectionTypeName(int typeValue)
{
    return typeValue switch
    {
        1 => "OLEDB",
        2 => "ODBC",
        3 => "XML",
        4 => "Text",
        5 => "Web",
        6 => "DataFeed",
        7 => "Model",
        8 => "Worksheet",
        9 => "NoSource",
        _ => $"Unknown({typeValue})"
    };
}

public static bool IsPowerQueryConnection(dynamic connection)
{
    try
    {
        // Power Query connections use the Mashup provider
        if (connection.Type == 1) // OLEDB
        {
            dynamic oledb = connection.OLEDBConnection;
            string connString = oledb.Connection?.ToString() ?? "";
            return connString.Contains("Microsoft.Mashup.OleDb.1", StringComparison.OrdinalIgnoreCase);
        }
    }
    catch { }
    return false;
}
```

### Shared Utilities to Create

```csharp
// src/ExcelMcp.Core/ExcelHelper.cs - ADD THESE

/// <summary>
/// Query table configuration options
/// </summary>
public class QueryTableOptions
{
    public bool BackgroundQuery { get; set; } = false;
    public bool PreserveColumnInfo { get; set; } = true;
    public bool PreserveFormatting { get; set; } = true;
    public bool AdjustColumnWidth { get; set; } = true;
    public bool RefreshOnFileOpen { get; set; } = false;
    public bool SavePassword { get; set; } = false; // Security: default false
}
```

### Refactoring Impact

**PowerQueryCommands.cs:**

- Replace `RemoveQueryConnections()` → call `ExcelHelper.RemoveConnections()`
- Replace `CreateQueryTableConnection()` → call `ExcelHelper.CreateQueryTable()`
- Use `ExcelHelper.SanitizeConnectionString()` in View/Export operations

**ConnectionCommands.cs (NEW):**

- Use all shared helpers from ExcelHelper
- Add connection-specific logic only
- Leverage existing patterns

**Benefits:**

- ✅ DRY compliance
- ✅ Consistent behavior across commands
- ✅ Easier testing (test shared utilities once)
- ✅ Reduced code duplication
- ✅ Easier maintenance

## Open Questions / Decisions Needed - ✅ ALL RESOLVED

### 1. ✅ Connection file formats: Support .odc, .iqy files?

**DECISION:** **YES - Phase 2 Enhancement**

- Phase 1: JSON format only (easy to read/write, version control friendly)
- Phase 2: Add `.odc` (Office Data Connection) file import
  - Use `Connections.AddFromFile(filePath)` COM method
  - Export via `connection.SaveAsODC(filePath)` method
- `.iqy` files: Legacy format, low priority

**Rationale:** JSON gives us full control and testability. `.odc` support adds enterprise compatibility later.

### 2. ✅ Connection validation: Add `conn-test` command?

**DECISION:** **YES - Essential Feature**

```powershell
excelcli conn-test "workbook.xlsx" "DatabaseConnection"
```

**Implementation:**

```csharp
public OperationResult Test(string filePath, string connectionName)
{
    // Attempt connection refresh and catch specific errors
    // Return success/failure with detailed error messages
    // Does NOT save workbook after test
}
```

**Rationale:** Users need to validate connections without risking data changes. Test-only refresh is critical.

### 3. ✅ Batch operations: Support `conn-refresh-all`?

**DECISION:** **NO - Too Risky**

**Why NOT implement:**

- `workbook.RefreshAll()` is known to hang (documented in copilot-instructions.md)
- User can script individual refresh operations if needed
- Better control with explicit per-connection refresh

**Alternative:**

```powershell
# Users can script batch refresh
foreach ($conn in (excelcli conn-list "file.xlsx" --output=json | ConvertFrom-Json).Connections) {
    excelcli conn-refresh "file.xlsx" $conn.Name
}
```

### 4. ✅ Connection templates: Provide sample JSON templates?

**DECISION:** **YES - Essential Documentation**

Create `docs/connection-templates/` with examples:

**SQL Server (OLEDB):**

```json
{
  "type": "OLEDB",
  "description": "SQL Server Production Database",
  "connectionString": "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=MyDB;Integrated Security=SSPI",
  "commandText": "SELECT * FROM Customers WHERE Active = 1",
  "commandType": "SQL",
  "backgroundQuery": false,
  "refreshOnFileOpen": false,
  "savePassword": false
}
```

**PostgreSQL (ODBC):**

```json
{
  "type": "ODBC",
  "description": "PostgreSQL Analytics Database",
  "connectionString": "Driver={PostgreSQL Unicode};Server=localhost;Port=5432;Database=analytics",
  "commandText": "SELECT * FROM sales_summary",
  "commandType": "SQL",
  "backgroundQuery": false,
  "refreshOnFileOpen": false,
  "savePassword": false
}
```

**CSV Text File:**

```json
{
  "type": "Text",
  "description": "Monthly Sales CSV Import",
  "filePath": "C:\\Data\\monthly_sales.csv",
  "delimiter": ",",
  "firstRowHasHeaders": true,
  "textFileOrigin": "Windows",
  "refreshOnFileOpen": true
}
```

### 5. ✅ Data Model: Expose ModelConnection properties?

**DECISION:** **NO - Future Enhancement**

**Why NOT in Phase 1:**

- ModelConnection (PowerPivot) is complex and requires separate expertise
- Focus on common connection types first (OLEDB, ODBC, Text, Web)
- Data Model manipulation needs comprehensive design (relationships, measures, etc.)

**Future Enhancement (Phase 3):**

- Dedicated `model-*` commands for Data Model operations
- Requires Analysis Services Tabular API integration
- Beyond scope of initial connection management

## Risk Assessment

| Risk | Likelihood | Impact | Mitigation |
|------|------------|--------|------------|
| Power Query connections interfere | Medium | Medium | Detect Mashup provider, redirect to pq-* commands |
| Password exposure | Low | High | Mandatory sanitization, security warnings |
| Connection hanging on refresh | Medium | Medium | Timeout handling, user guidance |
| Type-specific complexity | High | Medium | Clear error messages per connection type |
| COM exceptions unclear | Medium | Low | Comprehensive error handling, user-friendly messages |

---

## Implementation Phases (Updated with DRY Refactoring)

### Phase 0: Shared Utilities Extraction (DRY Compliance) 🔧

**Estimate:** 2-3 hours
**Priority:** MUST DO FIRST

- [ ] Extract shared utilities to `ExcelHelper.cs`:
  - [ ] `FindConnection()` - Connection finder pattern
  - [ ] `GetConnectionNames()` - List all connection names
  - [ ] `GetConnectionTypeName()` - Type enum to string mapper
  - [ ] `IsPowerQueryConnection()` - Detect Mashup provider
  - [ ] `SanitizeConnectionString()` - Password removal/masking
  - [ ] `RemoveConnections()` - Generic connection removal
  - [ ] `RemoveQueryTables()` - Generic QueryTable cleanup
  - [ ] `CreateQueryTable()` - Generic QueryTable creation with options
  - [ ] `QueryTableOptions` - Configuration class

- [ ] Refactor `PowerQueryCommands.cs` to use shared utilities:
  - [ ] Replace `RemoveQueryConnections()` with shared version
  - [ ] Replace `CreateQueryTableConnection()` with shared version
  - [ ] Add connection string sanitization to View/Export
  
- [ ] Create unit tests for shared utilities:
  - [ ] `ExcelHelperConnectionTests.cs`
  - [ ] `ConnectionStringSanitizerTests.cs`
  - [ ] `ConnectionTypeMapperTests.cs`

**Deliverable:** Shared utilities tested and PowerQuery refactored

---

### Phase 1: Core Implementation ✅

**Estimate:** 6-8 hours
**Dependencies:** Phase 0 complete

- [ ] Create `ConnectionCommands.cs` (Core layer)
  - [ ] Implement `List()` - Connection enumeration
  - [ ] Implement `View()` - Connection details with sanitization
  - [ ] Implement `Import()` - Create from JSON definition
  - [ ] Implement `Export()` - Save to JSON with sanitization
  - [ ] Implement `Update()` - Modify from JSON
  - [ ] Implement `Refresh()` - Refresh data
  - [ ] Implement `Delete()` - Remove connection
  - [ ] Implement `LoadTo()` - Create QueryTable
  - [ ] Implement `GetProperties()` - View settings
  - [ ] Implement `SetProperties()` - Modify settings
  - [ ] Implement `Test()` - Validate without saving

- [ ] Add connection result types to `ResultTypes.cs` (already done ✅)

- [ ] Create connection definition JSON templates:
  - [ ] `docs/connection-templates/oledb-sqlserver.json`
  - [ ] `docs/connection-templates/odbc-postgresql.json`
  - [ ] `docs/connection-templates/text-csv.json`

- [ ] Unit tests for Core layer (~45 tests):
  - [ ] `ConnectionJsonParserTests.cs`
  - [ ] Connection definition validation tests

- [ ] Integration tests for Core layer (~80 tests):
  - [ ] `CoreConnectionCommandsTests.cs` - CRUD operations
  - [ ] `ConnectionSecurityTests.cs` - Password sanitization
  - [ ] `ConnectionTypeTests.cs` - Type-specific operations

**Deliverable:** Core commands functional with 95%+ test coverage

---

### Phase 2: CLI Integration

**Estimate:** 4-5 hours
**Dependencies:** Phase 1 complete

- [ ] Create `ConnectionCommands.cs` (CLI layer)
  - [ ] `List()` - Format as Spectre.Console table
  - [ ] `View()` - Display with syntax highlighting
  - [ ] `Import()` - CLI wrapper with validation
  - [ ] `Export()` - CLI wrapper with success message
  - [ ] `Update()` - CLI wrapper with confirmation
  - [ ] `Refresh()` - Show progress indicators
  - [ ] `Delete()` - Confirmation prompt
  - [ ] `LoadTo()` - Progress + success message
  - [ ] `GetProperties()` - Format as table
  - [ ] `SetProperties()` - Parse flags + confirmation
  - [ ] `Test()` - Display test results with colors

- [ ] Update `Program.cs` routing:
  - [ ] Add `conn-*` command routing
  - [ ] Update help text with connection commands

- [ ] CLI tests (~20 tests):
  - [ ] `CliConnectionCommandsTests.cs` - Argument validation
  - [ ] `CliConnectionCommandsTests.cs` - Integration execution

**Deliverable:** CLI commands functional and documented

---

### Phase 3: MCP Server Integration

**Estimate:** 3-4 hours
**Dependencies:** Phase 2 complete

- [ ] Create `ExcelConnectionTool.cs` (MCP Server layer)
  - [ ] Implement action routing (list, view, import, etc.)
  - [ ] Add JSON serialization for results
  - [ ] Add MCP exception handling
  - [ ] Add detailed error messages for LLMs

- [ ] Update `server.json` with tool definition:
  - [ ] Add `excel_connection` tool
  - [ ] Define action enum
  - [ ] Document parameters

- [ ] MCP Server tests (~25 tests):
  - [ ] `ExcelConnectionToolTests.cs` - Tool actions
  - [ ] `McpConnectionWorkflowTests.cs` - Protocol workflows

**Deliverable:** MCP Server tool functional

---

### Phase 4: Testing & Documentation

**Estimate:** 3-4 hours
**Dependencies:** Phase 3 complete

- [ ] Round trip tests (~15 tests):
  - [ ] `ConnectionWorkflowTests.cs` - End-to-end workflows
  - [ ] Multi-connection scenarios
  - [ ] Error recovery scenarios

- [ ] Update documentation:
  - [ ] `README.md` - Add connection commands to capabilities
  - [ ] `COMMANDS.md` - Document all `conn-*` commands
  - [ ] `docs/DEVELOPMENT.md` - Add connection development patterns
  - [ ] `.github/copilot-instructions.md` - Connection management patterns
  - [ ] `src/ExcelMcp.McpServer/README.md` - Document `excel_connection` tool

- [ ] Create usage examples:
  - [ ] SQL Server connection workflow
  - [ ] PostgreSQL ODBC workflow
  - [ ] CSV text file import
  - [ ] Connection troubleshooting guide

**Deliverable:** Complete documentation + 95%+ test coverage

---

### Phase 5: .ODC File Support (Optional Enhancement)

**Estimate:** 2-3 hours
**Dependencies:** Phase 4 complete

- [ ] Implement `.odc` file import:
  - [ ] `conn-import-odc` command
  - [ ] Use `Connections.AddFromFile()` COM method

- [ ] Implement `.odc` file export:
  - [ ] `conn-export-odc` command
  - [ ] Use `connection.SaveAsODC()` COM method

- [ ] Add tests for `.odc` operations

**Deliverable:** Enterprise-compatible connection file format support

---

## Approval Checklist

Before implementation, please review and confirm:

- [x] **Functional requirements** - 11 operations defined (including test command) ✅
- [x] **Architecture alignment** - Mirrors PowerQuery patterns exactly ✅
- [x] **DRY compliance** - Shared utilities extracted, no duplication ✅
- [x] **Security considerations** - Password sanitization mandatory ✅
- [x] **Limitations documented** - Clear what cannot be done ✅
- [x] **Testing strategy** - 140+ tests planned, 95%+ coverage goal ✅
- [x] **Documentation scope** - Comprehensive with templates ✅
- [x] **Open questions resolved** - All 5 questions answered ✅
- [x] **Success criteria** - Clear deliverables per phase ✅

**Status:** � **READY FOR IMPLEMENTATION**

**Estimated Total Time:** 20-27 hours (5 phases)

**Next Step:** Begin Phase 0 (Shared Utilities Extraction)
