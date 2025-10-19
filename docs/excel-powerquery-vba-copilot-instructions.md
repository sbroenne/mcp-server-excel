# ExcelMcp MCP Server - Excel PowerQuery VBA Copilot Instructions

Copy this file to your project's `.github/`-directory to enable GitHub Copilot support for ExcelMcp MCP Server - AI-powered Excel development with PowerQuery and VBA capabilities.

---

## ExcelMcp MCP Server Integration

This project uses the ExcelMcp MCP Server for AI-assisted Excel development. The MCP server provides conversational access to Excel operations through 6 resource-based tools, enabling AI assistants like GitHub Copilot, Claude, and ChatGPT to perform Excel development tasks through natural language.

### Available MCP Server Tools

The MCP server provides 6 resource-based tools for Excel automation:

**1. excel_file** - File management
- Actions: `create-empty`, `validate`, `check-exists`
- Create Excel workbooks (.xlsx) or macro-enabled workbooks (.xlsm)
- Validate file format and accessibility

**2. excel_powerquery** - Power Query operations
- Actions: `list`, `view`, `import`, `export`, `update`, `refresh`, `delete`
- Manage Power Query M code for data transformations
- List, view, and edit Power Query connections
- Refresh queries and load results to worksheets

**3. excel_worksheet** - Worksheet operations
- Actions: `list`, `read`, `write`, `create`, `rename`, `copy`, `delete`, `clear`, `append`
- Manage worksheets and data ranges
- Read/write data from/to Excel worksheets
- Create, rename, copy, and delete worksheets

**4. excel_parameter** - Named range management
- Actions: `list`, `get`, `set`, `create`, `delete`
- Manage Excel named ranges as configuration parameters
- Get and set parameter values

**5. excel_cell** - Cell operations
- Actions: `get-value`, `set-value`, `get-formula`, `set-formula`
- Read and write individual cell values
- Manage cell formulas

**6. excel_vba** - VBA script management ⚠️ **Requires .xlsm files!**
- Actions: `list`, `export`, `import`, `update`, `run`, `delete`
- Manage VBA modules and procedures
- Execute VBA macros with parameters
- Export/import VBA code for version control

### Conversational Workflow Examples

**Power Query Refactoring:**

Ask Copilot: "Review the Power Query 'WebData' in analysis.xlsx and optimize it for performance"

The AI will:
1. View the current M code
2. Analyze for performance issues
3. Suggest and apply optimizations
4. Update the query with improved code

**VBA Enhancement:**

Ask Copilot: "Add comprehensive error handling to the DataProcessor module in automation.xlsm"

The AI will:
1. Export the current VBA module
2. Analyze the code structure
3. Add try-catch patterns and logging
4. Update the module with enhanced code

**Combined Power Query + VBA Workflow:**

Ask Copilot: "Set up a data pipeline in report.xlsm that loads data via Power Query, then generates charts with VBA"

The AI will:
1. Import or create Power Query for data loading
2. Refresh the query to get latest data
3. Import or create VBA module for chart generation
4. Execute the VBA macro to create visualizations

**Excel Development Automation:**

Ask Copilot: "Create a macro-enabled workbook with a data loader query and a formatting macro"

The AI will:
1. Create a new .xlsm file
2. Import Power Query M code for data loading
3. Import VBA module for formatting
4. Set up named ranges for parameters

### When to Use ExcelMcp MCP Server

The MCP server is ideal for AI-assisted Excel development workflows:

- **Power Query Refactoring** - AI analyzes and optimizes M code for performance
- **VBA Code Enhancement** - Add error handling, logging, and best practices to VBA modules
- **Code Review** - AI reviews existing Power Query/VBA code for issues and improvements
- **Development Automation** - AI creates and configures Excel workbooks with queries and macros
- **Documentation Generation** - Auto-generate comments and documentation for Excel code
- **Debugging Assistance** - AI helps troubleshoot Power Query and VBA issues
- **Best Practices Implementation** - AI applies Excel development patterns and standards

### AI Development vs. Scripted Automation

**Use MCP Server (AI-Assisted Development):**
- Conversational workflows with AI guidance
- Code refactoring and optimization
- Adding features to existing Excel solutions
- Learning and discovering Excel capabilities
- Complex multi-step development tasks

**Use CLI (Scripted Automation):**
- Available separately for scripted workflows
- See CLI.md documentation for command-line automation
- Ideal for CI/CD pipelines and batch processing

### File Format Requirements

- **Standard Excel files (.xlsx)**: For Power Query, worksheets, parameters, and cell operations
- **Macro-enabled files (.xlsm)**: Required for all VBA script operations
- **VBA Trust**: Must be enabled for VBA operations (MCP server can guide setup)

### Requirements

- Windows operating system
- Microsoft Excel installed
- .NET 10 runtime
- MCP server running (via `dnx Sbroenne.ExcelMcp.McpServer@latest`)
- For VBA operations: VBA trust must be enabled

### Getting Started

1. **Install the MCP Server:**
   ```powershell
   # Install .NET 10 SDK
   winget install Microsoft.DotNet.SDK.10
   
   # Run MCP server
   dnx Sbroenne.ExcelMcp.McpServer@latest --yes
   ```

2. **Configure AI Assistant:**
   - Add MCP server to your AI assistant configuration
   - See README.md for GitHub Copilot, Claude, and ChatGPT setup

3. **Start Conversational Development:**
   - Ask AI to perform Excel operations naturally
   - AI uses MCP tools to interact with Excel files
   - Get real-time feedback and suggestions

### Example Prompts for Copilot

- "Review this Power Query and suggest performance improvements"
- "Add comprehensive error handling to the VBA module"
- "Create a macro-enabled workbook with a data loader query"
- "Optimize this M code for better query folding"
- "Export all VBA modules from this workbook for version control"
- "Set up named ranges for report parameters"
- "Debug why this Power Query isn't refreshing properly"
