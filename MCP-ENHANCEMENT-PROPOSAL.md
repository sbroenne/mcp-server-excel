# MCP Server Enhancement Proposal

**Date**: October 27, 2025  
**Status**: Proposal  
**Target**: ExcelMcp.McpServer MCP Protocol Compliance & Feature Enhancement

---

## Executive Summary

This proposal outlines enhancements to bring ExcelMcp.McpServer to full MCP specification compliance and improve LLM interaction patterns. Based on the [MCP Specification 2025-06-18](https://modelcontextprotocol.io/specification/2025-06-18/) and [VS Code MCP Integration Guide](https://code.visualstudio.com/api/extension-guides/ai/mcp), we recommend implementing:

1. **Resources** - Expose Excel data as dynamic resources for context attachment
2. **Completions** - Provide autocomplete suggestions for file paths, query names, etc.
3. **Enhanced Prompts** - Add more educational prompts with parameters
4. **Icons** - Visual branding for tools and resources
5. **Sampling** - (Future) Allow server to request LLM help for complex operations

**Impact**: Better VS Code integration, improved discoverability, enhanced user experience

---

## Current Implementation Status

### ✅ What We Have

| Feature | Status | Implementation |
|---------|--------|----------------|
| **Tools** | ✅ Complete | 9 tools with batch session support |
| **Prompts** | ✅ Basic | 2 prompts for batch session education |
| **Stdio Transport** | ✅ Complete | Standard input/output communication |
| **HTTP Transport** | ❌ Not implemented | Could enable remote scenarios |
| **Resources** | ❌ Not implemented | Missing file/data exposure |
| **Completions** | ❌ Not implemented | No autocomplete support |
| **Sampling** | ❌ Not implemented | No LLM request capability |
| **Icons** | ❌ Not implemented | No visual branding |

### 📊 Our Tools (9 total)

1. `excel_file` - File creation/management
2. `excel_powerquery` - Power Query M code (11 actions)
3. `excel_connection` - Data connections (11 actions)
4. `excel_datamodel` - Data Model & DAX (8 actions)
5. `excel_worksheet` - Worksheet operations (10 actions)
6. `excel_parameter` - Named ranges (5 actions)
7. `excel_cell` - Cell operations (3 actions)
8. `excel_vba` - VBA management (7 actions)
9. `excel_version` - Version checking (1 action)

**Plus Batch Management**:
- `begin_excel_batch`
- `commit_excel_batch`
- `list_excel_batches`

---

## Recommended Enhancements

## 1. Resources 📚

### What Are Resources?

Resources expose data/content that users can:
- Attach as context to chat prompts
- Browse via MCP Resources Quick Pick in VS Code
- Access directly or via resource templates (with parameters)

### Proposed Resources

#### 1.1 Excel File Metadata Resource

**Resource URI**: `excel://file/{filePath}`

**Purpose**: Expose workbook structure and metadata

**Example Content**:
```json
{
  "name": "Sales Report",
  "path": "C:\\data\\sales.xlsx",
  "worksheets": ["Sheet1", "Summary", "RawData"],
  "powerQueries": ["SalesData", "CustomerLookup"],
  "namedRanges": ["ReportDate", "TotalSales"],
  "connections": ["SQL_Connection"],
  "hasMacros": true,
  "lastModified": "2025-10-27T10:30:00Z"
}
```

**Use Case**: LLM can understand workbook structure before suggesting operations

#### 1.2 Power Query Code Resource

**Resource URI**: `excel://query/{filePath}/{queryName}`

**Purpose**: Expose M code for analysis/refactoring

**Example Content**:
```m
let
    Source = Sql.Database("localhost", "Sales"),
    FilteredRows = Table.SelectRows(Source, each [Date] >= #date(2025,1,1)),
    RemovedColumns = Table.RemoveColumns(FilteredRows, {"Internal_ID"})
in
    RemovedColumns
```

**Use Case**: LLM can analyze query performance, suggest optimizations, detect issues

#### 1.3 Worksheet Data Resource

**Resource URI**: `excel://worksheet/{filePath}/{sheetName}`

**Purpose**: Expose worksheet data (first 100 rows for preview)

**Example Content**:
```json
{
  "worksheet": "Sheet1",
  "rows": 1500,
  "columns": 8,
  "preview": [
    ["Date", "Product", "Sales", "Region"],
    ["2025-01-01", "Widget A", "1234.56", "North"],
    ["2025-01-02", "Widget B", "2345.67", "South"]
  ],
  "truncated": true
}
```

**Use Case**: LLM can understand data structure before writing formulas/queries

#### 1.4 Data Model Structure Resource

**Resource URI**: `excel://datamodel/{filePath}`

**Purpose**: Expose Data Model tables, relationships, measures

**Example Content**:
```json
{
  "tables": [
    {"name": "Sales", "rowCount": 10000},
    {"name": "Products", "rowCount": 50}
  ],
  "relationships": [
    {
      "from": "Sales[ProductID]",
      "to": "Products[ProductID]",
      "type": "many-to-one"
    }
  ],
  "measures": [
    {"name": "TotalSales", "expression": "SUM(Sales[Amount])"}
  ]
}
```

**Use Case**: LLM can suggest relationship improvements, identify missing measures

#### 1.5 VBA Modules Resource

**Resource URI**: `excel://vba/{filePath}/{moduleName}`

**Purpose**: Expose VBA code for analysis

**Example Content**:
```vba
Sub ProcessData()
    ' Auto-generated documentation by LLM possible
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Data")
    ' ... macro code ...
End Sub
```

**Use Case**: LLM can document VBA, suggest refactoring, find bugs

### Implementation Approach

```csharp
// In a new ResourceProvider.cs file
[McpServerResourceType]
public static class ExcelResourceProvider
{
    [McpServerResource("excel://file/{filePath}")]
    [Description("Get Excel workbook structure and metadata")]
    public static async Task<ResourceContents> GetFileMetadata(
        [Description("Path to Excel file")] string filePath)
    {
        // Implementation using batch API
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        
        var metadata = new
        {
            name = Path.GetFileNameWithoutExtension(filePath),
            path = filePath,
            worksheets = await GetWorksheetsAsync(batch),
            powerQueries = await GetQueriesAsync(batch),
            // ... gather metadata
        };
        
        return new ResourceContents
        {
            Uri = $"excel://file/{filePath}",
            MimeType = "application/json",
            Text = JsonSerializer.Serialize(metadata, new JsonSerializerOptions { WriteIndented = true })
        };
    }
    
    // Resource templates for parameterized access
    [McpServerResourceTemplate("excel://query/{filePath}/{queryName}")]
    [Description("Get Power Query M code")]
    public static async Task<ResourceContents> GetQueryCode(
        [Description("Path to Excel file")] string filePath,
        [Description("Power Query name")] string queryName)
    {
        // Implementation
    }
}
```

### VS Code Integration

Users can:
1. Run "MCP: Browse Resources" command
2. See all Excel resources listed
3. Attach to chat prompt as context
4. Resources update in real-time (if we implement resource updates)

---

## 2. Completions 🔧

### What Are Completions?

Autocomplete suggestions for prompt/resource parameters - like IDE code completion.

### Proposed Completions

#### 2.1 File Path Completion

**Scenario**: User typing `excelPath` parameter

**Completion Source**: Recently opened files, current directory .xlsx/.xlsm files

**Example**:
```
User types: "C:\data\sa"
Completions:
  - C:\data\sales.xlsx
  - C:\data\sample.xlsm
  - C:\data\sales_2025.xlsx
```

#### 2.2 Power Query Name Completion

**Scenario**: User specifying `queryName` parameter

**Completion Source**: Queries in the specified workbook

**Example**:
```
User types: "Sal"
Completions (from sales.xlsx):
  - SalesData
  - SalesTransformed
  - SalesSummary
```

#### 2.3 Worksheet Name Completion

**Scenario**: User specifying `sheetName` parameter

**Completion Source**: Worksheets in the workbook

**Example**:
```
User types: "Sum"
Completions:
  - Summary
  - SummaryQ1
  - SummaryQ2
```

#### 2.4 Action Completion

**Scenario**: User specifying `action` parameter

**Completion Source**: Valid actions for the tool

**Example**:
```
Tool: excel_powerquery
User types: "im"
Completions:
  - import
  - (other actions not shown since they don't match)
```

### Implementation Approach

```csharp
// In Program.cs or new CompletionProvider.cs
builder.Services.AddMcpServer()
    .WithStdioServerTransport()
    .WithToolsFromAssembly()
    .WithCompletions(); // Enable completion capability

// Implement completion handlers
public class ExcelCompletionProvider
{
    public async Task<CompletionResult> CompleteAsync(CompletionRequest request)
    {
        if (request.Ref.Type == "ref/prompt" || request.Ref.Type == "ref/resource")
        {
            var argumentName = request.Argument.Name;
            var currentValue = request.Argument.Value;
            
            return argumentName switch
            {
                "excelPath" => await CompleteFilePathAsync(currentValue),
                "queryName" => await CompleteQueryNameAsync(currentValue, request.Context),
                "sheetName" => await CompleteSheetNameAsync(currentValue, request.Context),
                "action" => CompleteActionName(currentValue, request.Context),
                _ => CompletionResult.Empty
            };
        }
        
        return CompletionResult.Empty;
    }
    
    private async Task<CompletionResult> CompleteFilePathAsync(string partialPath)
    {
        var directory = Path.GetDirectoryName(partialPath) ?? Environment.CurrentDirectory;
        var searchPattern = Path.GetFileName(partialPath) + "*";
        
        var files = Directory.GetFiles(directory, searchPattern)
            .Where(f => f.EndsWith(".xlsx") || f.EndsWith(".xlsm"))
            .Take(100)
            .ToList();
            
        return new CompletionResult
        {
            Values = files,
            Total = files.Count,
            HasMore = false
        };
    }
}
```

### VS Code Integration

When user types in a parameter field:
- Dropdown appears with suggestions
- Filter as user types
- Select to autocomplete
- Works in tool confirmation dialog, prompt parameter dialog

---

## 3. Enhanced Prompts 📝

### Current Prompts

We have 2 prompts:
1. `excel_batch_guide` - Batch session comprehensive guide
2. `excel_batch_reference` - Quick reference

### Proposed Additional Prompts

#### 3.1 Excel Automation Patterns Prompt

**Name**: `excel_automation_patterns`

**Purpose**: Teach LLMs common Excel automation workflows

**Parameters**:
- `scenario` (completable): "data-import", "reporting", "validation", "transformation"

**Example**:
```csharp
[McpServerPrompt(Name = "excel_automation_patterns")]
[Description("Learn common Excel automation patterns and best practices")]
public static ChatMessage AutomationPatterns(
    [Description("Automation scenario: data-import, reporting, validation, transformation")]
    string scenario = "data-import")
{
    var content = scenario switch
    {
        "data-import" => @"# Data Import Pattern
1. Create workbook: excel_file create-empty
2. Add query: excel_powerquery import
3. Configure load: excel_powerquery set-load-to-table
4. Refresh: excel_powerquery refresh
5. Validate: excel_worksheet read",
        
        "reporting" => @"# Automated Reporting Pattern
1. Begin batch session
2. Update parameters: excel_parameter set
3. Refresh queries: excel_powerquery refresh
4. Update summary: excel_worksheet write
5. Commit batch",
        // ... other scenarios
    };
    
    return new ChatMessage(ChatRole.User, content);
}
```

**Completion Support**:
```csharp
// Provide completions for 'scenario' parameter
public static IEnumerable<string> CompleteScenario(string partialValue)
{
    return new[] { "data-import", "reporting", "validation", "transformation" }
        .Where(s => s.StartsWith(partialValue));
}
```

#### 3.2 Power Query Best Practices Prompt

**Name**: `excel_powerquery_guide`

**Purpose**: Educate about M code best practices

**Parameters**:
- `topic`: "performance", "error-handling", "data-types", "functions"

#### 3.3 Troubleshooting Prompt

**Name**: `excel_troubleshooting`

**Purpose**: Help diagnose common issues

**Parameters**:
- `issue`: "refresh-failed", "connection-error", "vba-error", "formula-error"

### Implementation

```csharp
// Update ExcelBatchPrompts.cs or create new prompt classes
namespace Sbroenne.ExcelMcp.McpServer.Prompts;

[McpServerPromptType]
public static class ExcelAutomationPrompts
{
    [McpServerPrompt(Name = "excel_automation_patterns")]
    [Description("Learn common Excel automation patterns for different scenarios")]
    public static ChatMessage AutomationPatterns(
        [Description("Scenario type")] 
        string scenario = "overview")
    {
        // Implementation
    }
    
    // Add more prompts...
}
```

---

## 4. Icons 🎨

### What Are Icons?

Visual indicators for tools/resources in VS Code UI:
- MCP servers list
- Tool picker in agent mode
- Resource picker
- Chat view

### Proposed Icons

#### 4.1 Server Icon

**Location**: .mcp folder or embedded resource

**Format**: PNG or Data URI

**Design**: Excel "X" with circuit board pattern (suggesting automation)

#### 4.2 Tool Icons

Different icons for different tool categories:

- 📊 `excel_worksheet` - Spreadsheet grid
- 🔌 `excel_connection` - Database/plug icon  
- 🧠 `excel_powerquery` - Flow diagram
- 📈 `excel_datamodel` - Relationship diagram
- 🎯 `excel_parameter` - Variable symbol
- 📝 `excel_vba` - Code brackets
- 📁 `excel_file` - Document icon

#### 4.3 Resource Icons

- 📄 File metadata
- 📊 Worksheet data
- 💾 Query code
- 🔗 Relationships

### Implementation

```csharp
// In Program.cs
builder.Services.AddMcpServer()
    .WithServerIcon(new Uri("file:///path/to/icon.png"))
    .WithStdioServerTransport()
    .WithToolsFromAssembly();

// In tool attributes
[McpServerTool(Name = "excel_powerquery")]
[Icon("data:image/png;base64,iVBORw0KGgoAAAANS...")]
public static async Task<string> ExcelPowerQuery(...)
{
    // Implementation
}

// Or use icon property if SDK supports
public class ExcelPowerQueryTool
{
    public static Icon ToolIcon => new Icon 
    { 
        Src = new Uri("file:///icons/powerquery.png") 
    };
}
```

### VS Code Integration

Icons appear:
- In "MCP: List Servers" view
- In tool picker dropdown (agent mode)
- In resource browser
- Next to tool names in chat confirmations

---

## 5. Sampling (Future Enhancement) 🤖

### What Is Sampling?

Allows MCP server to **request LLM help** for complex operations.

### Use Cases for ExcelMcp

#### 5.1 Complex M Code Generation

**Scenario**: User asks to "create query that joins sales and products"

**Server behavior**:
1. Receives request via tool
2. Sends sampling request to LLM: "Generate M code to join these tables: [metadata]"
3. LLM returns M code
4. Server validates and imports query
5. Returns result to user

#### 5.2 Data Transformation Analysis

**Scenario**: User uploads CSV, asks "clean this data"

**Server behavior**:
1. Reads CSV structure
2. Samples LLM: "Suggest Power Query transformations for: [data sample]"
3. LLM suggests steps
4. Server applies transformations
5. Returns cleaned data

#### 5.3 DAX Formula Generation

**Scenario**: "Create measure for year-over-year growth"

**Server behavior**:
1. Gets Data Model structure
2. Samples LLM: "Generate DAX for YoY growth given: [tables/columns]"
3. Validates DAX syntax
4. Creates measure
5. Returns result

### Implementation (When SDK Supports)

```csharp
// Hypothetical API
public async Task<string> GeneratePowerQuery(string userRequest, TableMetadata[] tables)
{
    var samplingRequest = new SamplingRequest
    {
        Messages = 
        [
            new ChatMessage(ChatRole.System, "You are a Power Query expert."),
            new ChatMessage(ChatRole.User, 
                $"Generate M code to: {userRequest}\nAvailable tables: {JsonSerializer.Serialize(tables)}")
        ],
        MaxTokens = 1000
    };
    
    var response = await _mcpServer.SampleAsync(samplingRequest);
    return response.Content.Text;
}
```

**Security Consideration**: Users must authorize sampling in VS Code settings.

---

## Implementation Roadmap

### Phase 1: Resources (High Priority) 🎯

**Effort**: 2-3 days  
**Impact**: High - enables context-aware chat

**Tasks**:
1. ✅ Research C# MCP SDK resource API
2. 🔲 Implement `ExcelResourceProvider` class
3. 🔲 Add file metadata resource
4. 🔲 Add Power Query code resource
5. 🔲 Add worksheet data resource (preview)
6. 🔲 Register resources in `Program.cs`
7. 🔲 Test in VS Code with "MCP: Browse Resources"
8. 🔲 Document in README.md

**Acceptance Criteria**:
- Resources appear in VS Code MCP Resources picker
- Resources can be attached to chat prompts
- Resources contain accurate, well-formatted data
- Resources update when workbook changes (via resource update notifications)

### Phase 2: Completions (Medium Priority) 🎯

**Effort**: 3-4 days  
**Impact**: Medium - improves UX, reduces errors

**Tasks**:
1. 🔲 Research C# MCP SDK completion API
2. 🔲 Implement `ExcelCompletionProvider` class
3. 🔲 Add file path completion
4. 🔲 Add query name completion
5. 🔲 Add worksheet name completion
6. 🔲 Add action completion
7. 🔲 Declare completion capability in Program.cs
8. 🔲 Test in VS Code tool confirmation dialogs
9. 🔲 Document in README.md

**Acceptance Criteria**:
- Completions appear when typing in parameter fields
- Completions filter based on partial input
- Completions are contextual (based on file/workbook)
- Max 100 items returned per completion
- Fast response (<500ms)

### Phase 3: Enhanced Prompts (Low Priority) 📝

**Effort**: 1-2 days  
**Impact**: Medium - educates LLMs about patterns

**Tasks**:
1. 🔲 Create `ExcelAutomationPrompts.cs`
2. 🔲 Add automation patterns prompt
3. 🔲 Add Power Query best practices prompt
4. 🔲 Add troubleshooting prompt
5. 🔲 Add completions for prompt parameters
6. 🔲 Test prompts in VS Code chat (slash commands)
7. 🔲 Document in README.md

**Acceptance Criteria**:
- Prompts available as slash commands (e.g., `/mcp.excel.excel_automation_patterns`)
- Prompts accept parameters with completions
- Prompts provide useful, actionable guidance
- Prompts integrate smoothly into chat flow

### Phase 4: Icons (Low Priority) 🎨

**Effort**: 1 day  
**Impact**: Low - improves branding, UX polish

**Tasks**:
1. 🔲 Design server icon (Excel + automation theme)
2. 🔲 Create tool category icons (8 icons)
3. 🔲 Create resource icons (4 icons)
4. 🔲 Add icons to server configuration
5. 🔲 Add icons to tool attributes
6. 🔲 Add icons to resource definitions
7. 🔲 Test icon display in VS Code
8. 🔲 Document in README.md

**Acceptance Criteria**:
- Icons appear in VS Code MCP server list
- Icons appear in tool picker
- Icons appear in resource browser
- Icons are consistent with brand
- Icons are clear at small sizes (16x16)

### Phase 5: Sampling (Future) 🚀

**Effort**: TBD (depends on SDK support)  
**Impact**: High - enables intelligent automation

**Tasks**:
1. 🔲 Wait for C# MCP SDK sampling support
2. 🔲 Design sampling use cases
3. 🔲 Implement sampling for M code generation
4. 🔲 Implement sampling for DAX generation
5. 🔲 Add user authorization checks
6. 🔲 Test sampling in various scenarios
7. 🔲 Document in README.md

**Acceptance Criteria**:
- Server can request LLM assistance
- User authorization required (VS Code setting)
- Sampling responses are validated before use
- Sampling enhances, not replaces, direct tool calls

---

## Technical Considerations

### C# MCP SDK Support

**Research Needed**:
1. Does current SDK version support resources? ✅ **YES** (based on MS docs)
2. Does current SDK version support completions? ⚠️ **VERIFY** (mentioned in spec)
3. Does current SDK version support icons? ⚠️ **VERIFY**
4. Does current SDK version support sampling? ❌ **NO** (not in current preview)

**SDK Version Check**:
```bash
dotnet list package | grep ModelContextProtocol
# Current: Check version in ExcelMcp.McpServer.csproj
```

**Action**: Review SDK release notes and GitHub repo for feature availability.

### Performance Impact

**Resources**:
- Opening workbooks to read metadata has overhead
- **Mitigation**: Cache resource content, use batch sessions, lazy loading

**Completions**:
- File system scans can be slow
- **Mitigation**: Limit to 100 results, use recent file cache, timeout after 500ms

**Icons**:
- Minimal impact (static files)

### Security Considerations

**Resources**:
- Don't expose sensitive data (passwords in connections)
- **Mitigation**: Already implemented in connection export (password sanitization)

**Completions**:
- Don't suggest files outside allowed directories
- **Mitigation**: Restrict to current directory + recent files

**Sampling** (future):
- User must authorize LLM access
- Don't send sensitive data to LLM without consent
- **Mitigation**: VS Code authorization dialog, configurable allowed models

### Breaking Changes

**None expected** - all enhancements are additive:
- Existing tools continue to work
- Existing prompts remain unchanged
- Resources/completions are new capabilities
- Backward compatible with older MCP clients

---

## Success Metrics

### Adoption Metrics
- Number of resources accessed per session
- Completion acceptance rate (% of completions used)
- Prompt invocation count

### Performance Metrics
- Resource generation time (<2 seconds)
- Completion response time (<500ms)
- Server startup time (no regression)

### Quality Metrics
- Resource accuracy (correct metadata)
- Completion relevance (user accepts suggestion)
- Prompt usefulness (user feedback)

### Integration Metrics
- VS Code MCP features utilized
- GitHub Copilot agent mode usage
- Claude Desktop integration success

---

## Alternatives Considered

### 1. HTTP Transport Instead of Stdio

**Pros**:
- Enables remote MCP server scenarios
- Better for cloud deployment
- Supports SSE (Server-Sent Events)

**Cons**:
- Adds complexity (authentication, CORS)
- Not needed for local Excel automation
- Current stdio works well for primary use case

**Decision**: Keep stdio as primary, add HTTP as optional later

### 2. Real-time Resource Updates

**Pros**:
- Resources stay in sync with workbook changes
- Better UX for long-running sessions

**Cons**:
- Requires file watching
- Added complexity
- MCP SDK support unclear

**Decision**: Start with static resources, add updates if SDK supports

### 3. Tool-Specific Subservers

**Pros**:
- Cleaner separation of concerns
- Independent versioning

**Cons**:
- More servers to manage
- Worse discoverability
- Unnecessary for current scope

**Decision**: Keep single unified server

---

## Documentation Updates Required

### README.md
- Add "Resources" section explaining available resources
- Add "Completions" section with examples
- Update "Prompts" section with new prompts
- Add screenshots of VS Code integration

### BATCH-SESSION-GUIDE.md
- Add resource usage examples
- Show how to attach resources to prompts

### New: RESOURCES-GUIDE.md
- Comprehensive resource documentation
- URI patterns and examples
- Use cases for each resource
- Performance tips

### New: COMPLETIONS-GUIDE.md
- How completions work
- Supported parameters
- Customization options

---

## Conclusion

Implementing **Resources** and **Completions** will significantly enhance ExcelMcp's MCP compliance and integration with VS Code/GitHub Copilot. These features align with the official MCP specification and provide tangible UX improvements.

**Recommended Priority**:
1. **Phase 1: Resources** (High Impact, Foundational)
2. **Phase 2: Completions** (High Value, UX Improvement)
3. **Phase 3: Enhanced Prompts** (Educational, Lower Priority)
4. **Phase 4: Icons** (Polish, Low Priority)
5. **Phase 5: Sampling** (Future, SDK-dependent)

**Next Steps**:
1. Verify C# MCP SDK support for resources/completions in current version
2. Create GitHub issue for Phase 1 implementation
3. Set up development branch: `feature/mcp-resources`
4. Begin implementation following test-driven approach

---

## References

- [MCP Specification 2025-06-18](https://modelcontextprotocol.io/specification/2025-06-18/)
- [VS Code MCP Developer Guide](https://code.visualstudio.com/api/extension-guides/ai/mcp)
- [MCP Completion Spec](https://modelcontextprotocol.io/specification/2025-06-18/server/utilities/completion)
- [MCP C# SDK](https://github.com/modelcontextprotocol/csharp-sdk)
- [Microsoft MCP Documentation](https://learn.microsoft.com/en-us/dotnet/ai/get-started-mcp)
- [Azure MCP Server Reference](https://github.com/Azure/azure-mcp)

---

**Author**: GitHub Copilot  
**Reviewed By**: [Your Name]  
**Approval Date**: [TBD]
